/**
 * Streamline Auto-Scheduler
 * Cascading dependency scheduling: when a task's dates change,
 * linked tasks auto-shift based on their dependency type + lag/lead.
 */

const { TaskType, DepType } = require("./dataModel");

/**
 * Auto-shift all downstream tasks based on dependency links.
 * Performs a topological traversal — each task is processed after all its predecessors.
 *
 * @param {Array} tasks - All tasks with resolved dependencies and dependencyTypes
 * @param {string} changedTaskId - The task that was moved/resized
 * @param {Object} newDates - { startDate, endDate } for the changed task
 * @returns {Map<string, {startDate, endDate}>} Map of taskId -> new dates for all shifted tasks
 */
function autoShift(tasks, changedTaskId, newDates) {
  const taskMap = new Map();
  for (const t of tasks) taskMap.set(t.id, t);

  // Build forward adjacency: predecessorId -> list of { successorId, type, lagDays }
  const forward = new Map();
  for (const t of tasks) {
    for (const depId of t.dependencies) {
      if (!forward.has(depId)) forward.set(depId, []);
      const info = t.dependencyTypes.get(depId) || { type: DepType.FS, lagDays: 0 };
      forward.get(depId).push({ successorId: t.id, type: info.type, lagDays: info.lagDays });
    }
  }

  // Compute updated dates via BFS from the changed task
  const updatedDates = new Map();
  updatedDates.set(changedTaskId, { ...newDates });

  // Apply the new dates to the changed task (for constraint computation)
  const changedTask = taskMap.get(changedTaskId);
  if (changedTask) {
    changedTask.startDate = newDates.startDate;
    changedTask.endDate = newDates.endDate;
  }

  const queue = [changedTaskId];
  const visited = new Set([changedTaskId]);

  while (queue.length > 0) {
    const currentId = queue.shift();
    const current = taskMap.get(currentId);
    if (!current) continue;

    const successors = forward.get(currentId) || [];
    for (const { successorId, type, lagDays } of successors) {
      if (visited.has(successorId)) continue;

      const successor = taskMap.get(successorId);
      if (!successor) continue;

      const predDates = updatedDates.get(currentId) || {
        startDate: current.startDate,
        endDate: current.endDate,
      };

      const newSuccDates = computeSuccessorDates(
        predDates, successor, type, lagDays
      );

      if (newSuccDates) {
        updatedDates.set(successorId, newSuccDates);
        successor.startDate = newSuccDates.startDate;
        successor.endDate = newSuccDates.endDate;
        visited.add(successorId);
        queue.push(successorId);
      }
    }
  }

  // Remove the originally changed task from result — caller already knows its dates
  updatedDates.delete(changedTaskId);
  return updatedDates;
}

/**
 * Compute new successor dates based on dependency type and lag.
 */
function computeSuccessorDates(predDates, successor, depType, lagDays) {
  const lagMs = lagDays * 24 * 60 * 60 * 1000;

  if (successor.type === TaskType.MILESTONE) {
    let anchorDate;
    switch (depType) {
      case DepType.FS:
        anchorDate = new Date((predDates.endDate || predDates.startDate).getTime() + lagMs);
        break;
      case DepType.FF:
        anchorDate = new Date((predDates.endDate || predDates.startDate).getTime() + lagMs);
        break;
      case DepType.SS:
        anchorDate = new Date(predDates.startDate.getTime() + lagMs);
        break;
      case DepType.SF:
        anchorDate = new Date(predDates.startDate.getTime() + lagMs);
        break;
      default:
        anchorDate = new Date((predDates.endDate || predDates.startDate).getTime() + lagMs);
    }
    return { startDate: anchorDate, endDate: null };
  }

  // Task: preserve duration, shift start
  const duration = successor.endDate && successor.startDate
    ? successor.endDate.getTime() - successor.startDate.getTime()
    : 0;

  let newStart;
  switch (depType) {
    case DepType.FS:
      // Successor starts after predecessor ends + lag
      newStart = new Date((predDates.endDate || predDates.startDate).getTime() + lagMs);
      break;
    case DepType.SS:
      // Successor starts when predecessor starts + lag
      newStart = new Date(predDates.startDate.getTime() + lagMs);
      break;
    case DepType.FF:
      // Successor ends when predecessor ends + lag → derive start from duration
      {
        const newEnd = new Date((predDates.endDate || predDates.startDate).getTime() + lagMs);
        newStart = new Date(newEnd.getTime() - duration);
      }
      break;
    case DepType.SF:
      // Successor ends when predecessor starts + lag → derive start from duration
      {
        const newEnd = new Date(predDates.startDate.getTime() + lagMs);
        newStart = new Date(newEnd.getTime() - duration);
      }
      break;
    default:
      newStart = new Date((predDates.endDate || predDates.startDate).getTime() + lagMs);
  }

  return {
    startDate: newStart,
    endDate: new Date(newStart.getTime() + duration),
  };
}

module.exports = { autoShift, computeSuccessorDates };
