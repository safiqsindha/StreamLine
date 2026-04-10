/**
 * Streamline Backend — Copilot action endpoints
 *
 * These four routes correspond 1:1 to operationIds declared in
 * copilot-package/streamline-actions.json:
 *   POST /createGantt
 *   POST /importFromM365
 *   POST /updateTasks
 *   POST /describeGantt
 *
 * Architecture note:
 *   Streamline's agentActions.js is host-agnostic and expects a context with
 *   { refreshController, templateManager, graphClient }. The first two
 *   operate on the active PowerPoint slide and ONLY exist client-side. The
 *   server therefore can't actually render shapes — instead it:
 *
 *     1. Validates the request
 *     2. For data-fetch actions, calls Graph via OBO and returns the rows
 *     3. For render actions, returns a "render-pending" payload that the
 *        Office add-in (or Copilot's host) consumes when the user opens
 *        PowerPoint to actually draw the chart
 *     4. For description actions, returns a stub summary (the real summary
 *        comes from the client-side describeGantt action because only the
 *        client knows what's currently on the slide)
 *
 * For dev/sideload testing, this is sufficient to verify the auth path,
 * the request schemas, and the Graph integration. A production deployment
 * would queue render-pending payloads in a per-user store the add-in polls.
 */

const express = require("express");
const { validateBootstrapToken, exchangeForGraphToken } = require("../auth");

const router = express.Router();

// Apply auth to every route in this router.
router.use(validateBootstrapToken);

// ── Helpers ────────────────────────────────────────────────────────────────

function badRequest(res, msg) {
  return res.status(400).json({ error: "bad_request", message: msg });
}

async function callGraph(req, path, scopes) {
  const token = await exchangeForGraphToken(req.user.rawJwt, scopes);
  const url = `https://graph.microsoft.com/v1.0${path}`;
  const r = await fetch(url, {
    headers: { Authorization: `Bearer ${token}`, Accept: "application/json" },
  });
  if (!r.ok) {
    const body = await r.text().catch(() => "");
    const e = new Error(`Graph ${path} → ${r.status}: ${body.slice(0, 200)}`);
    e.status = r.status;
    throw e;
  }
  return r.json();
}

// ── /createGantt ───────────────────────────────────────────────────────────
//
// Accepts a structured task list and returns a "render-pending" envelope.
// The Office add-in picks this up via long-polling or via a Copilot host
// callback to actually render shapes on the active slide.
router.post("/createGantt", (req, res) => {
  const body = req.body || {};
  if (!Array.isArray(body.tasks)) {
    return badRequest(res, "request.tasks (array) is required");
  }

  // Compute a quick summary so Copilot can describe what will be rendered
  // even before the client picks up the render envelope.
  const taskCount = body.tasks.filter((t) => (t.type || "Task") === "Task").length;
  const milestoneCount = body.tasks.filter((t) => t.type === "Milestone").length;
  const lanes = new Set(body.tasks.map((t) => t.swimLane).filter(Boolean));

  const dates = body.tasks
    .flatMap((t) => [t.startDate, t.endDate])
    .filter(Boolean)
    .sort();

  res.json({
    status: "render-pending",
    projectName: body.projectName || "Copilot Gantt",
    summary: {
      swimLaneCount: lanes.size,
      taskCount,
      milestoneCount,
      startDate: dates[0] || null,
      endDate: dates[dates.length - 1] || null,
    },
    renderRequest: {
      action: "createGantt",
      projectName: body.projectName || "Copilot Gantt",
      tasks: body.tasks,
      templateKey: body.templateKey || null,
    },
  });
});

// ── /importFromM365 ────────────────────────────────────────────────────────
//
// This one DOES call Graph server-side, because the data fetch is exactly
// the kind of work that benefits from being centralized. The returned rows
// are still rendered by the client.
router.post("/importFromM365", async (req, res) => {
  const body = req.body || {};
  const source = body.source;
  if (!source) return badRequest(res, "source is required");

  try {
    let rows = null;

    if (source === "planner") {
      if (!body.planId) return badRequest(res, "planId is required for source=planner");
      const [tasks, buckets] = await Promise.all([
        callGraph(req, `/planner/plans/${encodeURIComponent(body.planId)}/tasks`, [
          "https://graph.microsoft.com/Tasks.Read",
        ]),
        callGraph(req, `/planner/plans/${encodeURIComponent(body.planId)}/buckets`, [
          "https://graph.microsoft.com/Tasks.Read",
        ]),
      ]);
      rows = { tasks: tasks.value || [], buckets: buckets.value || [] };
    } else if (source === "todo") {
      if (!body.listId) return badRequest(res, "listId is required for source=todo");
      const tasks = await callGraph(
        req,
        `/me/todo/lists/${encodeURIComponent(body.listId)}/tasks`,
        ["https://graph.microsoft.com/Tasks.Read"]
      );
      rows = { tasks: tasks.value || [] };
    } else if (source === "calendar") {
      const from = body.fromDate || new Date().toISOString();
      const to =
        body.toDate ||
        new Date(Date.now() + 90 * 24 * 60 * 60 * 1000).toISOString();
      const events = await callGraph(
        req,
        `/me/calendar/calendarView?startDateTime=${encodeURIComponent(
          from
        )}&endDateTime=${encodeURIComponent(to)}`,
        ["https://graph.microsoft.com/Calendars.Read"]
      );
      rows = { events: events.value || [] };
    } else if (source === "onedrive") {
      if (!body.driveItemId) return badRequest(res, "driveItemId is required for source=onedrive");
      // Just return the item metadata; the client downloads the actual file
      // via its own graphClient (avoids streaming binary through this hop).
      const item = await callGraph(
        req,
        `/me/drive/items/${encodeURIComponent(body.driveItemId)}`,
        ["https://graph.microsoft.com/Files.Read"]
      );
      rows = { item };
    } else if (source === "sharepoint") {
      if (!body.siteId || !body.listId) {
        return badRequest(res, "siteId and listId are required for source=sharepoint");
      }
      const items = await callGraph(
        req,
        `/sites/${encodeURIComponent(body.siteId)}/lists/${encodeURIComponent(
          body.listId
        )}/items?expand=fields`,
        ["https://graph.microsoft.com/Sites.Read.All"]
      );
      rows = { items: items.value || [] };
    } else {
      return badRequest(res, `unknown source: ${source}`);
    }

    res.json({
      status: "render-pending",
      source,
      rows,
      renderRequest: {
        action: "importFromM365",
        source,
        rows,
        templateKey: body.templateKey || null,
      },
    });
  } catch (err) {
    const status = err.status || 500;
    console.error(`[copilot] importFromM365(${source}) failed:`, err.message);
    res.status(status).json({ error: "import_failed" });
  }
});

// ── /updateTasks ───────────────────────────────────────────────────────────
//
// Server can't reach into the active slide, so this returns a queued
// update envelope the client picks up.
router.post("/updateTasks", (req, res) => {
  const body = req.body || {};
  if (!Array.isArray(body.updates)) {
    return badRequest(res, "request.updates (array) is required");
  }

  res.json({
    status: "update-pending",
    matched: 0, // client will fill in real numbers after running the update
    updated: 0,
    cascaded: 0,
    notFound: [],
    renderRequest: {
      action: "updateTasks",
      updates: body.updates,
    },
  });
});

// ── /describeGantt ─────────────────────────────────────────────────────────
//
// The real description requires reading the active slide, which only the
// client can do. Server returns a stub plus a flag the client uses to know
// it should run the local describeGantt and respond to Copilot directly.
router.post("/describeGantt", (_req, res) => {
  res.json({
    status: "describe-pending",
    note:
      "Description must be generated client-side from the active slide. " +
      "The Office add-in will pick up this request and post the real summary back to Copilot.",
    renderRequest: { action: "describeGantt" },
  });
});

module.exports = router;
