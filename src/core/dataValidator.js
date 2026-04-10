/**
 * Streamline Data Validator
 * Validates parsed Excel rows for required fields, date integrity, and dependency cycles.
 */

const { TaskType } = require("./dataModel");

function validateRows(rows) {
  const errors = [];
  const warnings = [];
  const taskNames = new Set();
  const duplicateNames = new Set();

  // First pass: collect all task names and check for duplicates
  for (const row of rows) {
    const name = row.taskName ? String(row.taskName).trim() : "";
    if (name) {
      if (taskNames.has(name.toLowerCase())) {
        duplicateNames.add(name);
      }
      taskNames.add(name.toLowerCase());
    }
  }

  if (duplicateNames.size > 0) {
    errors.push(
      `Duplicate task names found: ${[...duplicateNames].join(", ")}. Each task must have a unique name.`
    );
  }

  // Second pass: validate each row
  for (const row of rows) {
    const rowLabel = `Row ${row._rowIndex}`;

    // Required: Swim Lane
    if (!row.swimLane || !String(row.swimLane).trim()) {
      errors.push(`${rowLabel}: Missing required "Swim Lane" value.`);
    }

    // Required: Task Name
    if (!row.taskName || !String(row.taskName).trim()) {
      errors.push(`${rowLabel}: Missing required "Task Name" value.`);
      continue;
    }

    // Required: Type
    const type = row.type ? String(row.type).trim().toUpperCase() : "";
    if (!type) {
      errors.push(`${rowLabel} ("${row.taskName}"): Missing required "Type" value. Must be "Task" or "Milestone".`);
    } else if (type !== "TASK" && type !== "MILESTONE") {
      errors.push(`${rowLabel} ("${row.taskName}"): Invalid Type "${row.type}". Must be "Task" or "Milestone".`);
    }

    // Required: Start Date
    if (!row.startDate) {
      errors.push(`${rowLabel} ("${row.taskName}"): Missing or invalid "Start Date".`);
    }

    // Conditional: End Date required for Tasks
    if (type === "TASK") {
      if (!row.endDate) {
        errors.push(`${rowLabel} ("${row.taskName}"): Tasks require an "End Date".`);
      } else if (row.startDate && row.endDate < row.startDate) {
        errors.push(`${rowLabel} ("${row.taskName}"): End Date is before Start Date.`);
      }
    }

    // Validate % Complete range
    if (row.percentComplete !== undefined && row.percentComplete !== null && row.percentComplete !== "") {
      const val = parseFloat(row.percentComplete);
      if (isNaN(val)) {
        warnings.push(`${rowLabel} ("${row.taskName}"): "% Complete" is not a number — will be ignored.`);
      } else if (val < 0 || val > 100) {
        warnings.push(`${rowLabel} ("${row.taskName}"): "% Complete" is ${val} — clamped to 0-100.`);
      }
    }

    // Validate baseline dates
    if (row.plannedStartDate && row.plannedEndDate && row.plannedEndDate < row.plannedStartDate) {
      warnings.push(`${rowLabel} ("${row.taskName}"): Planned End Date is before Planned Start Date.`);
    }

    // Validate dependency references exist (strip bracket notation for lookup)
    if (row.dependency) {
      const deps = String(row.dependency)
        .split(",")
        .map((d) => d.trim())
        .filter(Boolean);

      for (const dep of deps) {
        // Strip [FS+5d] style notation for name lookup
        const nameOnly = dep.replace(/\s*\[[^\]]*\]\s*$/, "").trim();
        if (!taskNames.has(nameOnly.toLowerCase())) {
          warnings.push(
            `${rowLabel} ("${row.taskName}"): Dependency "${nameOnly}" does not match any task name.`
          );
        }
      }
    }
  }

  // Check for circular dependencies
  const circularErrors = detectCircularDependencies(rows);
  errors.push(...circularErrors);

  return {
    isValid: errors.length === 0,
    errors,
    warnings,
  };
}

function detectCircularDependencies(rows) {
  const errors = [];

  const graph = new Map();

  for (const row of rows) {
    const name = row.taskName ? String(row.taskName).trim().toLowerCase() : "";
    if (!name) continue;
    graph.set(name, []);
  }

  for (const row of rows) {
    const name = row.taskName ? String(row.taskName).trim().toLowerCase() : "";
    if (!name || !row.dependency) continue;

    const deps = String(row.dependency)
      .split(",")
      .map((d) => {
        // Strip bracket notation
        const nameOnly = d.replace(/\s*\[[^\]]*\]\s*$/, "").trim().toLowerCase();
        return nameOnly;
      })
      .filter(Boolean);

    for (const dep of deps) {
      if (graph.has(dep)) {
        graph.get(name).push(dep);
      }
    }
  }

  const visited = new Set();
  const inStack = new Set();

  function dfs(node, path) {
    if (inStack.has(node)) {
      const cycleStart = path.indexOf(node);
      const cycle = path.slice(cycleStart).concat(node);
      errors.push(`Circular dependency detected: ${cycle.join(" -> ")}`);
      return;
    }
    if (visited.has(node)) return;

    visited.add(node);
    inStack.add(node);
    path.push(node);

    for (const neighbor of graph.get(node) || []) {
      dfs(neighbor, path);
    }

    path.pop();
    inStack.delete(node);
  }

  for (const node of graph.keys()) {
    if (!visited.has(node)) {
      dfs(node, []);
    }
  }

  return errors;
}

module.exports = {
  validateRows,
  detectCircularDependencies,
};
