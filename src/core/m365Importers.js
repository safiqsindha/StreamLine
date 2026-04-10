/**
 * Streamline M365 Importers
 *
 * Converts Microsoft Graph payloads (Planner tasks, To Do tasks, SharePoint
 * list items, Outlook calendar events) into Streamline's canonical row shape:
 *
 *   { swimLane, taskName, type, startDate, endDate, percentComplete, status,
 *     dependency, owner, notes }
 *
 * These rows feed straight into RefreshController.generateFromRows() - the
 * same pipeline used by Excel import, MPP import, and the data editor. This
 * module contains zero Graph auth / zero network code; it takes parsed JSON
 * payloads and returns Streamline rows. That keeps it pure-testable.
 */

// ── Planner ────────────────────────────────────────────────────────────

/**
 * Convert Planner tasks into Streamline rows.
 * Buckets become swim lanes; tasks become task bars.
 *
 * @param {Array} plannerTasks   Tasks from GET /planner/plans/{id}/tasks
 * @param {Array} plannerBuckets Buckets from GET /planner/plans/{id}/buckets (for lane naming)
 * @returns {Array<Row>}
 */
function plannerToRows(plannerTasks, plannerBuckets = []) {
  if (!Array.isArray(plannerTasks)) {
    throw new Error("plannerToRows: plannerTasks must be an array");
  }

  const bucketMap = new Map();
  for (const b of plannerBuckets) {
    if (b && b.id) bucketMap.set(b.id, b.name || "General");
  }

  const rows = [];
  for (const t of plannerTasks) {
    if (!t || !t.title) continue;

    const bucketName = (t.bucketId && bucketMap.get(t.bucketId)) || "General";
    const start = parseDateSafe(t.startDateTime);
    const due = parseDateSafe(t.dueDateTime);

    // Milestone heuristic: if it has a due date but no start, treat it as a milestone.
    const isMilestone = !start && !!due;

    rows.push({
      swimLane: bucketName,
      taskName: t.title,
      type: isMilestone ? "Milestone" : "Task",
      startDate: start || due || null,
      endDate: isMilestone ? null : (due || start || null),
      percentComplete: typeof t.percentComplete === "number" ? t.percentComplete : null,
      status: mapPlannerStatus(t.percentComplete),
      dependency: "",
      owner: extractPlannerOwner(t),
      notes: "",
    });
  }
  return rows;
}

function mapPlannerStatus(percent) {
  if (percent === null || percent === undefined) return "";
  if (percent >= 100) return "Complete";
  if (percent > 0) return "On Track";
  return "On Track";
}

function extractPlannerOwner(task) {
  if (!task.assignments) return null;
  // assignments is { userId: { ... } } - return first user ID as a stand-in name
  const ids = Object.keys(task.assignments);
  return ids.length > 0 ? ids[0] : null;
}

// ── To Do ──────────────────────────────────────────────────────────────

/**
 * Convert To Do tasks into Streamline rows.
 * The list name becomes the swim lane; each task becomes a row.
 *
 * @param {Array} todoTasks Tasks from GET /me/todo/lists/{id}/tasks
 * @param {string} listName The owning list's displayName
 */
function todoToRows(todoTasks, listName = "To Do") {
  if (!Array.isArray(todoTasks)) {
    throw new Error("todoToRows: todoTasks must be an array");
  }

  const rows = [];
  for (const t of todoTasks) {
    if (!t || !t.title) continue;

    // To Do uses { dateTime, timeZone } objects
    const start = parseTodoDate(t.startDateTime);
    const due = parseTodoDate(t.dueDateTime);
    const isMilestone = !start && !!due;

    let status = "On Track";
    if (t.status === "completed") status = "Complete";
    else if (t.status === "waitingOnOthers") status = "At Risk";
    else if (t.status === "deferred") status = "Delayed";

    const percentComplete = t.status === "completed" ? 100 : null;

    rows.push({
      swimLane: listName,
      taskName: t.title,
      type: isMilestone ? "Milestone" : "Task",
      startDate: start || due || null,
      endDate: isMilestone ? null : (due || start || null),
      percentComplete,
      status,
      dependency: "",
      owner: null,
      notes: t.body && t.body.content ? stripHtml(t.body.content).slice(0, 200) : "",
    });
  }
  return rows;
}

function parseTodoDate(obj) {
  if (!obj) return null;
  // { dateTime: "2026-04-15T00:00:00.0000000", timeZone: "UTC" }
  if (typeof obj === "string") return parseDateSafe(obj);
  if (obj.dateTime) return parseDateSafe(obj.dateTime);
  return null;
}

// ── Outlook Calendar ───────────────────────────────────────────────────

/**
 * Convert Outlook calendar events into Streamline rows. Each event becomes a
 * milestone (single point in time) or a task (multi-day event). Great for
 * overlaying meetings, deadlines, and project kickoffs onto a Gantt.
 */
function calendarToRows(events, laneName = "Calendar") {
  if (!Array.isArray(events)) {
    throw new Error("calendarToRows: events must be an array");
  }
  const rows = [];
  for (const ev of events) {
    if (!ev || !ev.subject) continue;
    const start = parseTodoDate(ev.start);
    const end = parseTodoDate(ev.end);
    if (!start && !end) continue;

    // If <= 1 day, treat as milestone
    const isMilestone =
      !start || !end || (end.getTime() - start.getTime() <= 24 * 60 * 60 * 1000);

    rows.push({
      swimLane: laneName,
      taskName: ev.subject,
      type: isMilestone ? "Milestone" : "Task",
      startDate: start || end,
      endDate: isMilestone ? null : end,
      percentComplete: null,
      status: "",
      dependency: "",
      owner: (ev.organizer && ev.organizer.emailAddress && ev.organizer.emailAddress.name) || null,
      notes: ev.bodyPreview || "",
    });
  }
  return rows;
}

// ── SharePoint Lists ───────────────────────────────────────────────────

/**
 * Convert SharePoint list items into Streamline rows. This is conventional:
 * we look for common field names (Title, StartDate, DueDate, Status, Owner,
 * PercentComplete, Dependency, SwimLane) in the expanded `fields` bag. Users
 * can pass a `fieldMap` override to remap any of these.
 */
function sharePointListToRows(items, options = {}) {
  if (!Array.isArray(items)) {
    throw new Error("sharePointListToRows: items must be an array");
  }
  const fm = {
    swimLane: "SwimLane",
    taskName: "Title",
    type: "TaskType",
    startDate: "StartDate",
    endDate: "DueDate",
    percentComplete: "PercentComplete",
    status: "Status",
    dependency: "Dependency",
    owner: "AssignedTo",
    notes: "Notes",
    ...options.fieldMap,
  };

  const defaultLane = options.defaultLane || "Tasks";
  const rows = [];

  for (const item of items) {
    const f = (item && item.fields) || {};
    const name = f[fm.taskName];
    if (!name) continue;

    const explicitType = f[fm.type];
    const end = parseDateSafe(f[fm.endDate]);
    const start = parseDateSafe(f[fm.startDate]);
    const isMilestone =
      (explicitType && String(explicitType).toLowerCase() === "milestone") ||
      (!start && !!end) ||
      (!end && !!start && !explicitType);

    rows.push({
      swimLane: f[fm.swimLane] || defaultLane,
      taskName: name,
      type: isMilestone ? "Milestone" : "Task",
      startDate: start || end || null,
      endDate: isMilestone ? null : (end || start || null),
      percentComplete: parseNumberSafe(f[fm.percentComplete]),
      status: f[fm.status] || "",
      dependency: f[fm.dependency] || "",
      owner: extractSharePointOwner(f[fm.owner]),
      notes: f[fm.notes] || "",
    });
  }
  return rows;
}

function extractSharePointOwner(val) {
  if (!val) return null;
  if (typeof val === "string") return val;
  // SP Person fields arrive as objects or arrays of objects
  if (Array.isArray(val)) {
    return val.map((u) => (u && (u.LookupValue || u.Title)) || "").filter(Boolean).join(", ") || null;
  }
  return val.LookupValue || val.Title || null;
}

// ── OneDrive file picker helpers ───────────────────────────────────────

/**
 * Filter a OneDrive children listing to items that look like Streamline can
 * open them (Excel, MS Project XML). Returns items with a `kind` annotation.
 */
function classifyDriveItems(items) {
  if (!Array.isArray(items)) return [];
  const out = [];
  for (const it of items) {
    if (!it || !it.name) continue;
    const lower = it.name.toLowerCase();
    let kind = null;
    if (lower.endsWith(".xlsx") || lower.endsWith(".xlsm") || lower.endsWith(".xls")) {
      kind = "excel";
    } else if (lower.endsWith(".xml")) {
      kind = "mpp-xml";
    } else if (lower.endsWith(".mpp")) {
      kind = "mpp-binary";
    }
    if (kind) {
      out.push({
        id: it.id,
        name: it.name,
        size: it.size || 0,
        lastModifiedDateTime: it.lastModifiedDateTime || null,
        webUrl: it.webUrl || null,
        kind,
      });
    }
  }
  return out;
}

// ── Utilities ──────────────────────────────────────────────────────────

function parseDateSafe(val) {
  if (!val) return null;
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
  if (typeof val === "object" && val.dateTime) return parseDateSafe(val.dateTime);
  const d = new Date(val);
  return isNaN(d.getTime()) ? null : d;
}

function parseNumberSafe(val) {
  if (val === null || val === undefined || val === "") return null;
  const n = Number(val);
  return isNaN(n) ? null : n;
}

function stripHtml(html) {
  return String(html).replace(/<[^>]*>/g, "").replace(/\s+/g, " ").trim();
}

module.exports = {
  plannerToRows,
  todoToRows,
  calendarToRows,
  sharePointListToRows,
  classifyDriveItems,
};
