/**
 * Streamline PowerPoint Function Commands
 *
 * Registered via Office.actions.associate and invoked by Copilot for
 * PowerPoint (or by custom ribbon buttons in the manifest). Each function
 * runs in a hidden HTML runtime - it has access to PowerPoint.run() and the
 * Office JS API but NO visible UI. Commands MUST call event.completed()
 * before returning.
 *
 * Copilot for PowerPoint calls these when the user asks things like:
 *   "Add a milestone to my Gantt for next Friday"
 *   "Refresh the Streamline chart"
 *   "Change the Streamline theme to dark mode"
 *
 * The manifest declares a FunctionFile ExtensionPoint that points at
 * commands.html, which loads this bundle and calls registerCommands().
 */

const { TemplateManager } = require("../core/templateManager");
const { RefreshController } = require("../core/refreshController");

// Lazy-initialized singletons - created on first command invocation so we
// don't pay the cost on every function-file load.
let _templateManager = null;
let _refreshController = null;
let _lastTasks = null;
let _lastLayout = null;
let _projectName = null;

function ensureControllers() {
  if (!_templateManager) _templateManager = new TemplateManager();
  if (!_refreshController) _refreshController = new RefreshController(_templateManager);
  return { templateManager: _templateManager, refreshController: _refreshController };
}

/**
 * Seed the command runtime with the current task pane state. The task pane
 * calls this via cross-frame messaging so the function commands can operate
 * on the chart the user is looking at.
 */
function seedState({ tasks, layout, projectName }) {
  _lastTasks = tasks || null;
  _lastLayout = layout || null;
  _projectName = projectName || null;
}

// ── refreshGantt ────────────────────────────────────────────────────────

/**
 * Refresh the currently-linked Gantt chart. Invoked by Copilot when the user
 * says "refresh the Streamline chart" or similar.
 */
async function refreshGantt(event) {
  try {
    const { refreshController } = ensureControllers();
    if (!refreshController.hasLinkedFile()) {
      await showNotification("No chart to refresh. Open Streamline and import data first.", "warning");
      event.completed();
      return;
    }
    const result = await refreshController.refresh(null, {});
    if (result.success) {
      _lastLayout = result.layout;
      _lastTasks = result.tasks;
      await showNotification(
        `Refreshed: ${result.stats.swimLanes} lanes, ${result.stats.totalTasks} tasks.`,
        "success"
      );
    } else {
      await showNotification(`Refresh failed: ${(result.errors || []).join("; ")}`, "error");
    }
  } catch (err) {
    await showNotification(`Refresh error: ${err.message}`, "error");
  }
  event.completed();
}

// ── applyTemplate ───────────────────────────────────────────────────────

/**
 * Apply a named template and re-render. Copilot calls this for "make my
 * Gantt look professional" or "switch to dark mode".
 */
async function applyTemplate(event) {
  try {
    const { templateManager, refreshController } = ensureControllers();
    // Template key comes in via event.source.id - the button/command sets
    // its own id which maps to a template key via the manifest's Control
    // definitions. For Copilot-initiated invocations, the template key is
    // supplied in event.source.parameters.
    const key =
      (event.source && event.source.parameters && event.source.parameters.templateKey) ||
      (event.source && event.source.id) ||
      "standard";

    try {
      templateManager.setActiveTemplate(key);
    } catch (_) {
      await showNotification(`Unknown template: ${key}`, "error");
      event.completed();
      return;
    }

    if (refreshController.hasLinkedFile()) {
      const result = await refreshController.refresh(null, {});
      if (result.success) {
        _lastLayout = result.layout;
        _lastTasks = result.tasks;
        await showNotification(`Applied template "${key}".`, "success");
      } else {
        await showNotification(`Template applied but refresh failed.`, "warning");
      }
    } else {
      await showNotification(`Template "${key}" selected. Import data to apply it.`, "info");
    }
  } catch (err) {
    await showNotification(`Template error: ${err.message}`, "error");
  }
  event.completed();
}

// ── toggleTodayMarker ───────────────────────────────────────────────────

/**
 * Toggle the "today" vertical line marker. Copilot calls this for "show the
 * today line on the Gantt" / "hide the today marker".
 */
async function toggleTodayMarker(event) {
  try {
    const { refreshController } = ensureControllers();
    if (!refreshController.hasLinkedFile()) {
      await showNotification("No Gantt chart to modify.", "warning");
      event.completed();
      return;
    }
    const show = !(event.source && event.source.parameters && event.source.parameters.hide);
    const result = await refreshController.refresh(null, { showTodayMarker: show });
    if (result.success) {
      _lastLayout = result.layout;
      await showNotification(`Today marker ${show ? "shown" : "hidden"}.`, "success");
    }
  } catch (err) {
    await showNotification(`Toggle error: ${err.message}`, "error");
  }
  event.completed();
}

// ── exportPng ──────────────────────────────────────────────────────────

/**
 * Export the current Gantt as a PNG. Runs in the command runtime, so we
 * can't actually trigger a browser download (no DOM). Instead we push a
 * message to the task pane via Office.context.ui.messageParent; the task
 * pane then runs the existing exportManager.downloadPNG() handler.
 */
async function exportPng(event) {
  try {
    if (!_lastLayout) {
      await showNotification("No chart to export.", "warning");
      event.completed();
      return;
    }
    // Function commands can't access the DOM directly; we signal the task
    // pane to perform the export.
    if (typeof Office !== "undefined" && Office.context && Office.context.ui && Office.context.ui.messageParent) {
      Office.context.ui.messageParent(JSON.stringify({ command: "exportPng" }));
    }
    await showNotification("Exporting PNG...", "info");
  } catch (err) {
    await showNotification(`Export error: ${err.message}`, "error");
  }
  event.completed();
}

// ── addMilestone ────────────────────────────────────────────────────────

/**
 * Add a milestone to the existing Gantt chart. Copilot calls this when the
 * user says "add a milestone called Launch Day on June 1".
 *
 * Parameters come via event.source.parameters:
 *   { name: string, date: "YYYY-MM-DD", swimLane?: string }
 */
async function addMilestone(event) {
  try {
    const params = (event.source && event.source.parameters) || {};
    if (!params.name || !params.date) {
      await showNotification("Milestone name and date required.", "error");
      event.completed();
      return;
    }
    if (!_lastTasks) {
      await showNotification("No Gantt to add milestone to. Create one first.", "warning");
      event.completed();
      return;
    }
    const { refreshController } = ensureControllers();

    const newRow = {
      swimLane: params.swimLane || _lastTasks[0].swimLane || "General",
      taskName: params.name,
      type: "Milestone",
      startDate: new Date(params.date),
      endDate: null,
      status: "",
      dependency: "",
      owner: params.owner || null,
      notes: params.notes || "",
    };

    const existingRows = _lastTasks.map(taskToRow);
    existingRows.push(newRow);

    const result = await refreshController.generateFromRows(
      existingRows,
      _projectName || "Copilot edit",
      {}
    );
    if (result.success) {
      _lastLayout = result.layout;
      _lastTasks = result.tasks;
      await showNotification(`Added milestone "${params.name}".`, "success");
    } else {
      await showNotification(`Add failed: ${(result.errors || []).join("; ")}`, "error");
    }
  } catch (err) {
    await showNotification(`Add milestone error: ${err.message}`, "error");
  }
  event.completed();
}

// ── showNotification ────────────────────────────────────────────────────

/**
 * Display a toast via Office's notification API. Falls back to console if
 * the API isn't available (e.g. when running under a test harness).
 */
async function showNotification(message, kind = "info") {
  if (typeof Office === "undefined" || !Office.context) {
    // eslint-disable-next-line no-console
    console.log(`[Streamline ${kind}] ${message}`);
    return;
  }
  try {
    // Office.NotificationMessageType: informationalMessage / errorMessage / progressIndicator
    const type =
      kind === "error" ? "errorMessage" :
      kind === "success" ? "informationalMessage" :
      kind === "warning" ? "informationalMessage" :
      "informationalMessage";
    const ppt = Office.context.document;
    if (ppt && ppt.notifications && ppt.notifications.addAsync) {
      ppt.notifications.addAsync("streamline-command", {
        type,
        message,
        icon: "Icon.16x16",
        persistent: false,
      });
    }
  } catch (_) { /* swallow - notifications are nice-to-have */ }
}

function taskToRow(t) {
  return {
    swimLane: t.swimLane,
    subSwimLane: t.subSwimLane,
    taskName: t.name,
    type: t.type === "MILESTONE" ? "Milestone" : "Task",
    startDate: t.startDate,
    endDate: t.endDate,
    percentComplete: t.percentComplete,
    status: t.status ? statusKeyToLabel(t.status) : "",
    dependency: (t.dependencyNames || []).join(", "),
    milestoneShape: t.milestoneShape,
    owner: t.owner,
    notes: t.notes,
  };
}

function statusKeyToLabel(key) {
  const map = { ON_TRACK: "On Track", AT_RISK: "At Risk", DELAYED: "Delayed", COMPLETE: "Complete" };
  return map[key] || "";
}

// ── Registration ────────────────────────────────────────────────────────

/**
 * Register all commands with Office. Called from commands.html once Office
 * is ready. Each function is addressable from the manifest by its key.
 */
function registerCommands() {
  if (typeof Office === "undefined" || !Office.actions || !Office.actions.associate) {
    // eslint-disable-next-line no-console
    console.warn("Streamline commands: Office.actions.associate not available in this host.");
    return;
  }
  Office.actions.associate("refreshGantt", refreshGantt);
  Office.actions.associate("applyTemplate", applyTemplate);
  Office.actions.associate("toggleTodayMarker", toggleTodayMarker);
  Office.actions.associate("exportPng", exportPng);
  Office.actions.associate("addMilestone", addMilestone);
}

module.exports = {
  registerCommands,
  seedState,
  // Exported directly for tests
  refreshGantt,
  applyTemplate,
  toggleTodayMarker,
  exportPng,
  addMilestone,
};
