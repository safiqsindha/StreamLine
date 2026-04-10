/**
 * Streamline Data Model
 * Internal representations for tasks, milestones, swim lanes, and dependencies.
 */

const TaskType = Object.freeze({
  TASK: "TASK",
  MILESTONE: "MILESTONE",
});

const TaskStatus = Object.freeze({
  ON_TRACK: "ON_TRACK",
  AT_RISK: "AT_RISK",
  DELAYED: "DELAYED",
  COMPLETE: "COMPLETE",
});

const STATUS_MAP = {
  "on track": TaskStatus.ON_TRACK,
  "at risk": TaskStatus.AT_RISK,
  delayed: TaskStatus.DELAYED,
  complete: TaskStatus.COMPLETE,
  completed: TaskStatus.COMPLETE,
};

// Dependency link types
const DepType = Object.freeze({
  FS: "FS", // Finish-to-Start (default)
  FF: "FF", // Finish-to-Finish
  SS: "SS", // Start-to-Start
  SF: "SF", // Start-to-Finish
});

// Milestone shape options
const MilestoneShape = Object.freeze({
  DIAMOND: "diamond",
  CIRCLE: "circle",
  TRIANGLE: "triangle",
  STAR: "star",
  FLAG: "flag",
  SQUARE: "square",
});

let _nextId = 1;

function generateId(prefix) {
  return `${prefix}_${String(_nextId++).padStart(3, "0")}`;
}

function resetIdCounter() {
  _nextId = 1;
}

/**
 * Parse dependency string with optional type and lag/lead.
 * Formats:
 *   "Task A"                -> { name: "Task A", type: FS, lag: 0 }
 *   "Task A [FF]"           -> { name: "Task A", type: FF, lag: 0 }
 *   "Task A [FS+5d]"        -> { name: "Task A", type: FS, lag: 5 }
 *   "Task A [SS-3d]"        -> { name: "Task A", type: SS, lag: -3 }
 */
function parseDependencyEntry(raw) {
  const trimmed = raw.trim();
  const bracketMatch = trimmed.match(/^(.+?)\s*\[([A-Z]{2})([+-]\d+[dwm])?\]$/i);

  if (bracketMatch) {
    const name = bracketMatch[1].trim();
    const type = bracketMatch[2].toUpperCase();
    let lagDays = 0;

    if (bracketMatch[3]) {
      const lagStr = bracketMatch[3];
      const num = parseInt(lagStr, 10);
      const unit = lagStr.slice(-1).toLowerCase();
      if (unit === "w") lagDays = num * 7;
      else if (unit === "m") lagDays = num * 30;
      else lagDays = num; // days
    }

    return {
      name,
      type: DepType[type] || DepType.FS,
      lagDays,
    };
  }

  return { name: trimmed, type: DepType.FS, lagDays: 0 };
}

function createTask(row) {
  const type =
    row.type && row.type.toUpperCase() === "MILESTONE"
      ? TaskType.MILESTONE
      : TaskType.TASK;

  const statusKey = row.status ? row.status.toLowerCase().trim() : null;
  const status = statusKey ? STATUS_MAP[statusKey] || null : null;

  const depString = row.dependency ? String(row.dependency).trim() : "";
  const rawDeps = depString
    ? depString.split(",").map((d) => d.trim()).filter(Boolean)
    : [];

  const parsedDeps = rawDeps.map(parseDependencyEntry);

  // Percent complete (0-100)
  let percentComplete = null;
  if (row.percentComplete !== undefined && row.percentComplete !== null && row.percentComplete !== "") {
    const val = parseFloat(row.percentComplete);
    if (!isNaN(val)) percentComplete = Math.max(0, Math.min(100, val));
  }

  // Milestone shape
  let milestoneShape = MilestoneShape.DIAMOND;
  if (row.milestoneShape) {
    const key = String(row.milestoneShape).toLowerCase().trim();
    milestoneShape = Object.values(MilestoneShape).includes(key) ? key : MilestoneShape.DIAMOND;
  }

  return {
    id: generateId(type === TaskType.MILESTONE ? "ms" : "task"),
    swimLane: row.swimLane ? String(row.swimLane).trim() : "",
    subSwimLane: row.subSwimLane ? String(row.subSwimLane).trim() : null,
    name: row.taskName ? String(row.taskName).trim() : "",
    type,
    startDate: row.startDate,
    endDate: type === TaskType.TASK ? row.endDate : null,
    plannedStartDate: row.plannedStartDate || null,
    plannedEndDate: row.plannedEndDate || null,
    percentComplete,
    milestoneShape,
    dependencyNames: rawDeps.map((d) => parseDependencyEntry(d).name),
    dependencyLinks: parsedDeps, // { name, type, lagDays }
    dependencies: [], // resolved task IDs
    dependencyTypes: new Map(), // depTaskId -> { type, lagDays }
    notes: row.notes ? String(row.notes).trim() : "",
    status,
    owner: row.owner ? String(row.owner).trim() : null,
  };
}

function createSwimLane(name, tasks) {
  // Group tasks by sub-swimlane
  const subLaneMap = new Map();
  const topLevelTasks = [];

  for (const task of tasks) {
    if (task.subSwimLane) {
      if (!subLaneMap.has(task.subSwimLane)) {
        subLaneMap.set(task.subSwimLane, []);
      }
      subLaneMap.get(task.subSwimLane).push(task);
    } else {
      topLevelTasks.push(task);
    }
  }

  const subLanes = [];
  for (const [subName, subTasks] of subLaneMap) {
    subLanes.push({
      id: generateId("sublane"),
      name: subName,
      tasks: subTasks,
      yOffset: 0,
      height: 0,
    });
  }

  return {
    id: generateId("lane"),
    name,
    tasks, // all tasks (for backward compat)
    topLevelTasks, // tasks not in a sub-lane
    subLanes,
    yOffset: 0,
    height: 0,
  };
}

function resolveDependencies(tasks) {
  const nameToId = new Map();
  for (const task of tasks) {
    nameToId.set(task.name.toLowerCase(), task.id);
  }

  for (const task of tasks) {
    task.dependencies = [];
    task.dependencyTypes = new Map();

    for (const link of task.dependencyLinks) {
      const depId = nameToId.get(link.name.toLowerCase());
      if (depId) {
        task.dependencies.push(depId);
        task.dependencyTypes.set(depId, { type: link.type, lagDays: link.lagDays });
      }
    }
  }
}

function groupIntoSwimLanes(tasks) {
  const laneMap = new Map();

  for (const task of tasks) {
    const laneName = task.swimLane || "Ungrouped";
    if (!laneMap.has(laneName)) {
      laneMap.set(laneName, []);
    }
    laneMap.get(laneName).push(task);
  }

  const lanes = [];
  for (const [name, laneTasks] of laneMap) {
    lanes.push(createSwimLane(name, laneTasks));
  }

  return lanes;
}

/**
 * Calculate schedule variance in days (actual vs planned).
 * Negative = ahead of schedule, Positive = behind schedule.
 */
function calculateVariance(task) {
  if (task.type === TaskType.MILESTONE) {
    if (!task.plannedStartDate || !task.startDate) return null;
    return (task.startDate - task.plannedStartDate) / (1000 * 60 * 60 * 24);
  }
  if (!task.plannedEndDate || !task.endDate) return null;
  return (task.endDate - task.plannedEndDate) / (1000 * 60 * 60 * 24);
}

/**
 * Calculate task duration in days.
 */
function getTaskDuration(task) {
  if (task.type === TaskType.MILESTONE) return 0;
  if (!task.startDate || !task.endDate) return 0;
  return Math.round((task.endDate - task.startDate) / (1000 * 60 * 60 * 24));
}

module.exports = {
  TaskType,
  TaskStatus,
  DepType,
  MilestoneShape,
  createTask,
  createSwimLane,
  resolveDependencies,
  groupIntoSwimLanes,
  calculateVariance,
  getTaskDuration,
  parseDependencyEntry,
  resetIdCounter,
};
