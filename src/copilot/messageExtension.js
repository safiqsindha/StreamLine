/**
 * Streamline Teams Message Extension
 *
 * Implements the three message extension commands declared in
 * teams-package/manifest.json:
 *
 *   1. searchGantts        - search-type command; returns a list of stored
 *                            Gantt charts matching the user's query as
 *                            adaptive cards suitable for Teams compose box
 *                            and Microsoft 365 Copilot chat.
 *   2. createGanttFromMessage - action-type command; extracts structured
 *                            tasks from the currently-selected message body
 *                            and runs the createGantt agent action.
 *   3. summarizeGantt      - action-type command; returns a stats summary
 *                            adaptive card for a previously-saved Gantt.
 *
 * The handlers here are transport-agnostic. They take a JSON payload and
 * return { type, attachments } objects that an Express/Azure Functions bot
 * adapter can hand back to the Teams Bot Framework. The bot adapter itself
 * lives outside this repo (it's a deploy-time concern), but the business
 * logic is here and is fully unit-testable.
 */

const {
  createGantt,
  describeGantt,
  importFromM365,
} = require("./agentActions");
const { classifyDriveItems } = require("../core/m365Importers");

// ── Handler: searchGantts ───────────────────────────────────────────────

/**
 * Search for Streamline-compatible files in OneDrive and return them as
 * adaptive card previews. Accepts parameters in Bot Framework compose
 * extension query format.
 *
 * @param {object} req      { query: string }
 * @param {object} context  { graphClient }
 * @returns {Promise<{ type: "result", attachmentLayout: "list", attachments: Array }>}
 */
async function handleSearchGantts(req, context) {
  const { graphClient } = context;
  if (!graphClient) throw new Error("searchGantts: graphClient required");
  if (!graphClient.hasAccessToken()) {
    return buildSignInResponse();
  }

  const query = (req && req.query) || "";
  if (!query || query.trim() === "") {
    // Initial run: list recent drive items that look like Streamline files
    const items = await graphClient.getOneDriveChildren("Streamline");
    return buildSearchResults(classifyDriveItems(items));
  }

  const hits = await graphClient.searchDrive(query);
  return buildSearchResults(classifyDriveItems(hits));
}

function buildSearchResults(items) {
  const attachments = items.slice(0, 25).map((item) => ({
    contentType: "application/vnd.microsoft.card.adaptive",
    content: buildGanttPreviewCard(item),
    preview: {
      contentType: "application/vnd.microsoft.card.thumbnail",
      content: {
        title: item.name,
        subtitle: formatKind(item.kind),
        text: item.lastModifiedDateTime ? `Modified ${item.lastModifiedDateTime.slice(0, 10)}` : "",
      },
    },
  }));
  return {
    type: "result",
    attachmentLayout: "list",
    attachments,
  };
}

function formatKind(kind) {
  switch (kind) {
    case "excel": return "Excel schedule";
    case "mpp-xml": return "MS Project XML";
    case "mpp-binary": return "MS Project file";
    default: return "Document";
  }
}

// ── Handler: createGanttFromMessage ─────────────────────────────────────

/**
 * Extract tasks from a message's body and build a Gantt chart from them.
 * The Teams adapter passes the selected message via req.commandContext.message
 * (in Bot Framework invoke payloads).
 *
 * Uses a simple line-based parser: each line becomes a task if it matches
 * "<name> - <start> to <end>" or "<name> [on <date>]" for milestones.
 */
async function handleCreateGanttFromMessage(req, context) {
  const message = req && req.messagePayload && req.messagePayload.body && req.messagePayload.body.content;
  if (!message) {
    return buildErrorCardResponse("No message body to extract tasks from.");
  }

  const tasks = extractTasksFromText(stripHtml(message));
  if (tasks.length === 0) {
    return buildErrorCardResponse(
      "Couldn't find any task lines in that message. Try lines like \"Task name - 2026-04-15 to 2026-05-01\"."
    );
  }

  const summary = await createGantt(
    { tasks, projectName: req.projectName || "From message" },
    context
  );

  return {
    type: "result",
    attachmentLayout: "list",
    attachments: [
      {
        contentType: "application/vnd.microsoft.card.adaptive",
        content: buildSummaryCard(summary),
      },
    ],
  };
}

function extractTasksFromText(text) {
  const lines = String(text).split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
  const tasks = [];
  // Match: "Name - 2026-04-15 to 2026-05-01"
  const taskRe = /^(.+?)\s*[-–]\s*(\d{4}-\d{2}-\d{2})\s+to\s+(\d{4}-\d{2}-\d{2})\s*(?:\((.+?)\))?$/i;
  // Match: "Name [on 2026-04-15]" or "Name: 2026-04-15" (milestones)
  const msRe = /^(.+?)\s*(?:\[on\s+|:\s+)(\d{4}-\d{2}-\d{2})\]?$/i;
  // Swim-lane delimiter: "## Lane name"
  let currentLane = "General";
  const laneRe = /^#{1,3}\s*(.+)$/;

  for (const line of lines) {
    const laneMatch = line.match(laneRe);
    if (laneMatch) { currentLane = laneMatch[1]; continue; }

    const taskMatch = line.match(taskRe);
    if (taskMatch) {
      tasks.push({
        swimLane: currentLane,
        taskName: taskMatch[1].trim(),
        type: "Task",
        startDate: taskMatch[2],
        endDate: taskMatch[3],
        owner: taskMatch[4] || null,
      });
      continue;
    }

    const msMatch = line.match(msRe);
    if (msMatch) {
      tasks.push({
        swimLane: currentLane,
        taskName: msMatch[1].trim(),
        type: "Milestone",
        startDate: msMatch[2],
        endDate: null,
      });
    }
  }
  return tasks;
}

// ── Handler: summarizeGantt ─────────────────────────────────────────────

/**
 * Return a stats summary adaptive card for a previously-saved Gantt.
 * In practice the Gantt ID is a drive item ID; we download it, parse it,
 * and call describeGantt against the loaded state.
 */
async function handleSummarizeGantt(req, context) {
  const ganttId = req && req.parameters && req.parameters.ganttId;
  if (!ganttId) return buildErrorCardResponse("ganttId parameter required.");

  const { graphClient, refreshController } = context;
  if (!graphClient || !graphClient.hasAccessToken()) return buildSignInResponse();

  // Pull the file, run it through the pipeline (generateFromRows happens inside importFromM365).
  await importFromM365(
    { source: "onedrive", driveItemId: ganttId, fileName: "Shared Gantt" },
    context
  );

  // describeGantt needs lastTasks/lastLayout in context
  const ctx = {
    ...context,
    lastTasks: refreshController && refreshController._lastTasks,
    lastLayout: refreshController && refreshController.lastLayout,
  };
  const summary = describeGantt({}, ctx);

  return {
    type: "result",
    attachmentLayout: "list",
    attachments: [
      {
        contentType: "application/vnd.microsoft.card.adaptive",
        content: buildSummaryCard(summary),
      },
    ],
  };
}

// ── Adaptive Card Builders ──────────────────────────────────────────────

/**
 * Build an adaptive card representing a single Gantt chart file.
 * Rendered inline in Teams messages and Copilot chat replies.
 */
function buildGanttPreviewCard(item) {
  return {
    type: "AdaptiveCard",
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    version: "1.5",
    body: [
      {
        type: "TextBlock",
        size: "Large",
        weight: "Bolder",
        text: item.name,
        wrap: true,
      },
      {
        type: "TextBlock",
        text: formatKind(item.kind),
        isSubtle: true,
        spacing: "None",
      },
      item.lastModifiedDateTime
        ? {
            type: "TextBlock",
            text: `Modified ${item.lastModifiedDateTime.slice(0, 10)}`,
            isSubtle: true,
            spacing: "Small",
          }
        : null,
      item.size
        ? {
            type: "TextBlock",
            text: `${formatSize(item.size)}`,
            isSubtle: true,
            spacing: "Small",
          }
        : null,
    ].filter(Boolean),
    actions: [
      {
        type: "Action.OpenUrl",
        title: "Open in Streamline",
        url: item.webUrl || "https://localhost:3000",
      },
      {
        type: "Action.Submit",
        title: "Import into Gantt",
        data: { msteams: { type: "invoke" }, action: "importDriveItem", driveItemId: item.id },
      },
    ],
  };
}

/**
 * Summary card: shows project stats for a Gantt. Used by both
 * createGanttFromMessage and summarizeGantt.
 */
function buildSummaryCard(summary) {
  return {
    type: "AdaptiveCard",
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    version: "1.5",
    body: [
      {
        type: "TextBlock",
        size: "Large",
        weight: "Bolder",
        text: summary.projectName || "Gantt Summary",
        wrap: true,
      },
      {
        type: "FactSet",
        facts: [
          { title: "Swim lanes", value: String(summary.swimLaneCount || 0) },
          { title: "Tasks", value: String(summary.taskCount || 0) },
          { title: "Milestones", value: String(summary.milestoneCount || 0) },
          { title: "Dependencies", value: String(summary.dependencyCount || 0) },
          summary.criticalPathLength
            ? { title: "Critical path", value: `${summary.criticalPathLength} task(s)` }
            : null,
          summary.startDate ? { title: "Start", value: summary.startDate } : null,
          summary.endDate ? { title: "Finish", value: summary.endDate } : null,
        ].filter(Boolean),
      },
      summary.atRiskTasks && summary.atRiskTasks.length
        ? {
            type: "TextBlock",
            text: `At risk: ${summary.atRiskTasks.slice(0, 5).join(", ")}${
              summary.atRiskTasks.length > 5 ? ` (+${summary.atRiskTasks.length - 5} more)` : ""
            }`,
            color: "Warning",
            wrap: true,
          }
        : null,
    ].filter(Boolean),
  };
}

function buildErrorCardResponse(message) {
  return {
    type: "result",
    attachmentLayout: "list",
    attachments: [
      {
        contentType: "application/vnd.microsoft.card.adaptive",
        content: {
          type: "AdaptiveCard",
          $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
          version: "1.5",
          body: [
            {
              type: "TextBlock",
              text: message,
              wrap: true,
              color: "Attention",
            },
          ],
        },
      },
    ],
  };
}

function buildSignInResponse() {
  return {
    type: "auth",
    suggestedActions: {
      actions: [
        {
          type: "openUrl",
          title: "Sign in to Microsoft 365",
          value: "https://localhost:3000/api/auth/signin",
        },
      ],
    },
  };
}

// ── Utilities ───────────────────────────────────────────────────────────

function stripHtml(html) {
  return String(html).replace(/<[^>]*>/g, " ").replace(/\s+/g, " ").trim();
}

function formatSize(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

// ── Dispatcher ──────────────────────────────────────────────────────────

const HANDLERS = {
  searchGantts: handleSearchGantts,
  createGanttFromMessage: handleCreateGanttFromMessage,
  summarizeGantt: handleSummarizeGantt,
};

/**
 * Bot framework entry: map a commandId from the compose extension invoke
 * payload to the correct handler.
 */
async function handleMessageExtensionInvoke(commandId, req, context) {
  const handler = HANDLERS[commandId];
  if (!handler) {
    return buildErrorCardResponse(`Unknown command: ${commandId}`);
  }
  return handler(req, context);
}

module.exports = {
  handleSearchGantts,
  handleCreateGanttFromMessage,
  handleSummarizeGantt,
  handleMessageExtensionInvoke,
  buildGanttPreviewCard,
  buildSummaryCard,
  extractTasksFromText,
  HANDLERS,
};
