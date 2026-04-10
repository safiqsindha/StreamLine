/**
 * Streamline Copilot Agent Actions
 *
 * Implementation of the actions declared in copilot-package/streamline-actions.json.
 * Each action is a pure function over inputs + an injectable context (RefreshController,
 * GraphClient, TemplateManager) so it can be unit-tested without PowerPoint.
 *
 * The actions are designed to be host-agnostic:
 *   - Called directly from the task pane (client-side) for inline testing
 *   - Called from a server endpoint (/api/copilot/*) when Copilot invokes them
 *   - Called from the Teams message extension handler
 *
 * All three contexts share the same code path.
 */

const {
  plannerToRows,
  todoToRows,
  calendarToRows,
  sharePointListToRows,
} = require("../core/m365Importers");

// ── createGantt ─────────────────────────────────────────────────────────

/**
 * Build a Gantt chart from a caller-supplied list of tasks.
 *
 * @param {object} req            Request body matching CreateGanttRequest schema
 * @param {object} context
 * @param {object} context.refreshController
 * @param {object} context.templateManager
 * @returns {Promise<object>} GanttSummary
 */
async function createGantt(req, context) {
  if (!req || !Array.isArray(req.tasks)) {
    throw new Error("createGantt: request.tasks array is required");
  }
  const { refreshController, templateManager } = context;
  if (!refreshController) throw new Error("createGantt: refreshController required in context");

  // Apply template if requested
  if (req.templateKey && templateManager) {
    try { templateManager.setActiveTemplate(req.templateKey); } catch (_) { /* fall back to active */ }
  }

  // Normalize task inputs → Streamline row shape
  const rows = req.tasks.map((t) => normalizeTaskInput(t));

  const result = await refreshController.generateFromRows(
    rows,
    req.projectName || "Copilot Gantt",
    context.config || {}
  );

  if (!result.success) {
    const err = new Error(`createGantt failed: ${(result.errors || []).join("; ")}`);
    err.validationErrors = result.errors;
    throw err;
  }

  return summarizeLayout(result, req.projectName || "Copilot Gantt");
}

// ── importFromM365 ──────────────────────────────────────────────────────

/**
 * Pull tasks from a Microsoft 365 source and render them as a Gantt chart.
 *
 * @param {object} req     ImportFromM365Request
 * @param {object} context { refreshController, templateManager, graphClient }
 */
async function importFromM365(req, context) {
  if (!req || !req.source) throw new Error("importFromM365: source required");
  const { refreshController, templateManager, graphClient } = context;
  if (!refreshController) throw new Error("importFromM365: refreshController required");
  if (!graphClient) throw new Error("importFromM365: graphClient required");
  if (!graphClient.hasAccessToken()) {
    throw new Error("importFromM365: Graph client has no access token. Sign in first.");
  }

  let rows = [];
  let label = req.source;

  switch (req.source) {
    case "planner": {
      if (!req.planId) throw new Error("planId required for source=planner");
      const [tasks, buckets] = await Promise.all([
        graphClient.getPlanTasks(req.planId),
        graphClient.getPlanBuckets(req.planId),
      ]);
      rows = plannerToRows(tasks, buckets);
      label = `Planner plan ${req.planId.slice(0, 8)}`;
      break;
    }
    case "todo": {
      if (!req.listId) throw new Error("listId required for source=todo");
      const tasks = await graphClient.getTodoTasks(req.listId);
      rows = todoToRows(tasks, req.listName || "To Do");
      label = req.listName || "To Do";
      break;
    }
    case "calendar": {
      const from = req.fromDate || new Date().toISOString();
      const to = req.toDate || new Date(Date.now() + 90 * 86400e3).toISOString();
      const events = await graphClient.getCalendarEvents(from, to);
      rows = calendarToRows(events, "Calendar");
      label = "Calendar";
      break;
    }
    case "onedrive": {
      if (!req.driveItemId) throw new Error("driveItemId required for source=onedrive");
      // OneDrive import pulls the file bytes and routes through the Excel/MPP pipeline.
      // Copilot agent callers should prefer this for .xlsx / .xml schedule files.
      const buffer = await graphClient.downloadDriveItem(req.driveItemId);
      const result = await refreshController.generate(
        buffer,
        req.fileName || "OneDrive import",
        context.config || {}
      );
      if (!result.success) {
        throw new Error(`OneDrive import failed: ${(result.errors || []).join("; ")}`);
      }
      return summarizeLayout(result, req.fileName || "OneDrive import");
    }
    case "sharepoint": {
      if (!req.siteId || !req.listId) throw new Error("siteId and listId required for source=sharepoint");
      const items = await graphClient.getListItems(req.siteId, req.listId);
      rows = sharePointListToRows(items, { defaultLane: req.defaultLane });
      label = "SharePoint list";
      break;
    }
    default:
      throw new Error(`Unknown source: ${req.source}`);
  }

  if (rows.length === 0) {
    throw new Error(`No tasks found in ${req.source}.`);
  }

  if (req.templateKey && templateManager) {
    try { templateManager.setActiveTemplate(req.templateKey); } catch (_) { /* ignore */ }
  }

  const result = await refreshController.generateFromRows(rows, label, context.config || {});
  if (!result.success) {
    throw new Error(`Import failed: ${(result.errors || []).join("; ")}`);
  }
  return summarizeLayout(result, label);
}

// ── updateTasks ─────────────────────────────────────────────────────────

/**
 * Apply one or more updates to tasks on the currently-rendered Gantt.
 * Matches tasks by name (case-insensitive) and cascades dependent dates.
 */
async function updateTasks(req, context) {
  if (!req || !Array.isArray(req.updates)) {
    throw new Error("updateTasks: request.updates array is required");
  }
  const { refreshController, lastTasks, autoShift } = context;
  if (!lastTasks) throw new Error("updateTasks: no active Gantt. Create one first.");
  if (!autoShift) throw new Error("updateTasks: autoShift function required in context");

  const byName = new Map();
  for (const t of lastTasks) byName.set(String(t.name).toLowerCase(), t);

  let matched = 0;
  let updated = 0;
  let cascaded = 0;
  const notFound = [];

  for (const u of req.updates) {
    const key = String(u.taskName || "").toLowerCase();
    const task = byName.get(key);
    if (!task) {
      notFound.push(u.taskName);
      continue;
    }
    matched++;
    let changed = false;

    if (u.newName) { task.name = u.newName; changed = true; }
    if (u.newStartDate) {
      task.startDate = parseIsoDate(u.newStartDate);
      changed = true;
    }
    if (u.newEndDate && task.type !== "MILESTONE") {
      task.endDate = parseIsoDate(u.newEndDate);
      changed = true;
    }
    if (u.newStatus) {
      task.status = labelToStatusKey(u.newStatus);
      changed = true;
    }
    if (typeof u.newPercentComplete === "number") {
      task.percentComplete = Math.max(0, Math.min(100, u.newPercentComplete));
      changed = true;
    }
    if (changed) updated++;

    // Cascade
    if (changed && (u.newStartDate || u.newEndDate)) {
      const shifted = autoShift(lastTasks, task.id, {
        startDate: task.startDate,
        endDate: task.endDate,
      });
      cascaded += shifted.size;
    }
  }

  // Re-render with the mutated tasks
  if (updated > 0 && refreshController) {
    const rows = tasksToRows(lastTasks);
    await refreshController.generateFromRows(rows, "Copilot update", context.config || {});
  }

  return { matched, updated, cascaded, notFound };
}

// ── describeGantt ───────────────────────────────────────────────────────

/**
 * Return a natural-language-friendly summary of the active Gantt chart.
 */
function describeGantt(req, context) {
  const { lastLayout, lastTasks } = context;
  if (!lastLayout || !lastTasks) {
    return {
      projectName: null,
      swimLaneCount: 0,
      taskCount: 0,
      milestoneCount: 0,
      dependencyCount: 0,
      criticalPathLength: 0,
      atRiskTasks: [],
    };
  }

  const tasks = lastTasks.filter((t) => t.type !== "MILESTONE");
  const milestones = lastTasks.filter((t) => t.type === "MILESTONE");
  const deps = lastTasks.reduce((sum, t) => sum + (t.dependencies || []).length, 0);
  const lanes = new Set(lastTasks.map((t) => t.swimLane)).size;
  const atRisk = lastTasks
    .filter((t) => t.status === "AT_RISK" || t.status === "DELAYED")
    .map((t) => t.name);

  const dates = lastTasks
    .flatMap((t) => [t.startDate, t.endDate])
    .filter((d) => d instanceof Date);
  const startDate = dates.length ? new Date(Math.min(...dates.map((d) => d.getTime()))) : null;
  const endDate = dates.length ? new Date(Math.max(...dates.map((d) => d.getTime()))) : null;

  return {
    projectName: context.projectName || "Streamline Gantt",
    swimLaneCount: lanes,
    taskCount: tasks.length,
    milestoneCount: milestones.length,
    dependencyCount: deps,
    criticalPathLength: (lastLayout.criticalPathIds && lastLayout.criticalPathIds.size) || 0,
    startDate: startDate ? isoDate(startDate) : null,
    endDate: endDate ? isoDate(endDate) : null,
    atRiskTasks: atRisk,
  };
}

// ── Helpers ─────────────────────────────────────────────────────────────

function normalizeTaskInput(t) {
  const type = t.type === "Milestone" || t.type === "MILESTONE" ? "Milestone" : "Task";
  return {
    swimLane: t.swimLane || "General",
    subSwimLane: t.subSwimLane || null,
    taskName: t.taskName || t.name || "Untitled",
    type,
    startDate: parseIsoDate(t.startDate),
    endDate: type === "Milestone" ? null : parseIsoDate(t.endDate || t.startDate),
    percentComplete: typeof t.percentComplete === "number" ? t.percentComplete : null,
    status: t.status || "",
    dependency: t.dependency || "",
    owner: t.owner || null,
    notes: t.notes || "",
  };
}

function parseIsoDate(val) {
  if (!val) return null;
  if (val instanceof Date) return val;
  const d = new Date(val);
  return isNaN(d.getTime()) ? null : d;
}

function isoDate(d) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

function labelToStatusKey(label) {
  const map = {
    "On Track": "ON_TRACK",
    "At Risk": "AT_RISK",
    Delayed: "DELAYED",
    Complete: "COMPLETE",
  };
  return map[label] || null;
}

function tasksToRows(tasks) {
  return tasks.map((t) => ({
    swimLane: t.swimLane,
    subSwimLane: t.subSwimLane,
    taskName: t.name,
    type: t.type === "MILESTONE" ? "Milestone" : "Task",
    startDate: t.startDate,
    endDate: t.endDate,
    plannedStartDate: t.plannedStartDate,
    plannedEndDate: t.plannedEndDate,
    percentComplete: t.percentComplete,
    status: statusKeyToLabel(t.status),
    dependency: (t.dependencyNames || []).join(", "),
    milestoneShape: t.milestoneShape,
    owner: t.owner,
    notes: t.notes,
  }));
}

function statusKeyToLabel(key) {
  const map = {
    ON_TRACK: "On Track",
    AT_RISK: "At Risk",
    DELAYED: "Delayed",
    COMPLETE: "Complete",
  };
  return map[key] || "";
}

function summarizeLayout(result, projectName) {
  const stats = result.stats || {};
  const tasks = result.tasks || [];
  const atRisk = tasks
    .filter((t) => t.status === "AT_RISK" || t.status === "DELAYED")
    .map((t) => t.name);

  const dates = tasks
    .flatMap((t) => [t.startDate, t.endDate])
    .filter((d) => d instanceof Date);
  const startDate = dates.length ? new Date(Math.min(...dates.map((d) => d.getTime()))) : null;
  const endDate = dates.length ? new Date(Math.max(...dates.map((d) => d.getTime()))) : null;

  return {
    projectName,
    swimLaneCount: stats.swimLanes || 0,
    taskCount: stats.taskBars || 0,
    milestoneCount: stats.milestones || 0,
    dependencyCount: stats.dependencies || 0,
    criticalPathLength: stats.criticalPathLength || 0,
    startDate: startDate ? isoDate(startDate) : null,
    endDate: endDate ? isoDate(endDate) : null,
    atRiskTasks: atRisk,
  };
}

// ── Action dispatcher ───────────────────────────────────────────────────

const ACTIONS = {
  createGantt,
  importFromM365,
  updateTasks,
  describeGantt,
};

/**
 * Dispatch a named action with the given request body and context.
 * Used by the server endpoint and the Teams message extension.
 */
async function dispatchAction(actionName, req, context) {
  const handler = ACTIONS[actionName];
  if (!handler) {
    const err = new Error(`Unknown action: ${actionName}`);
    err.code = "UNKNOWN_ACTION";
    throw err;
  }
  return handler(req, context);
}

module.exports = {
  createGantt,
  importFromM365,
  updateTasks,
  describeGantt,
  dispatchAction,
  ACTIONS,
  // exported for test visibility
  normalizeTaskInput,
  summarizeLayout,
};
