/**
 * Streamline Refresh Controller
 * Manages the full lifecycle: parse -> validate -> layout -> render.
 * Handles both initial generation and refresh with shape clearing.
 */

const { parseExcelFile } = require("./excelParser");
const { validateRows } = require("./dataValidator");
const {
  createTask,
  resolveDependencies,
  groupIntoSwimLanes,
  resetIdCounter,
} = require("./dataModel");
const { calculateLayout } = require("./layoutEngine");
const { renderGantt, clearStreamlineShapes } = require("./powerpointRenderer");

class RefreshController {
  constructor(templateManager) {
    this.templateManager = templateManager;
    this.lastExcelData = null;
    this.lastLayout = null;
    this.linkedFileName = null;
  }

  /**
   * Full generation pipeline: Excel ArrayBuffer -> rendered Gantt chart.
   * @param {ArrayBuffer} excelArrayBuffer - Raw Excel file data
   * @param {string} fileName - Name of the imported file
   * @param {object} config - Optional layout config overrides
   * @returns {object} Result with status, warnings, and stats
   */
  async generate(excelArrayBuffer, fileName, config = {}) {
    resetIdCounter();

    // Step 1: Parse Excel
    const parsed = parseExcelFile(excelArrayBuffer);

    // Step 2: Validate
    const validation = validateRows(parsed.rows);

    if (!validation.isValid) {
      return {
        success: false,
        phase: "validation",
        errors: validation.errors,
        warnings: validation.warnings,
      };
    }

    // Step 3: Build task model
    const tasks = parsed.rows.map((row) => createTask(row));
    resolveDependencies(tasks);
    const swimLanes = groupIntoSwimLanes(tasks);

    // Step 4: Calculate layout
    const template = this.templateManager.getActiveTemplate();
    const layout = calculateLayout(swimLanes, template, config);

    // Step 5: Clear existing shapes and render
    await clearStreamlineShapes();
    await renderGantt(layout);

    // Store state for refresh
    this.lastExcelData = excelArrayBuffer;
    this.lastLayout = layout;
    this.linkedFileName = fileName;

    return {
      success: true,
      phase: "complete",
      warnings: validation.warnings,
      layout,
      tasks,
      parsedRows: parsed.rows,
      stats: {
        sheetName: parsed.sheetName,
        totalTasks: tasks.length,
        swimLanes: swimLanes.length,
        milestones: tasks.filter((t) => t.type === "MILESTONE").length,
        taskBars: tasks.filter((t) => t.type === "TASK").length,
        dependencies: tasks.reduce(
          (sum, t) => sum + (t.dependencies ? t.dependencies.length : 0),
          0
        ),
        criticalPathLength: layout.criticalPathIds.size,
        scalingMode: layout.scaling.labelMode,
      },
    };
  }

  /**
   * Refresh: re-read the provided Excel data and regenerate the Gantt chart.
   * @param {ArrayBuffer} excelArrayBuffer - Updated Excel file data
   * @param {object} config - Optional layout config overrides
   */
  async refresh(excelArrayBuffer, config = {}) {
    const data = excelArrayBuffer || this.lastExcelData;
    if (!data) {
      return {
        success: false,
        phase: "refresh",
        errors: ["No Excel data available. Please import a file first."],
        warnings: [],
      };
    }

    return this.generate(data, this.linkedFileName, config);
  }

  /**
   * Preview: parse and validate without rendering.
   * Returns layout data for inspection.
   */
  preview(excelArrayBuffer, config = {}) {
    resetIdCounter();

    const parsed = parseExcelFile(excelArrayBuffer);
    const validation = validateRows(parsed.rows);

    if (!validation.isValid) {
      return {
        success: false,
        phase: "validation",
        errors: validation.errors,
        warnings: validation.warnings,
      };
    }

    const tasks = parsed.rows.map((row) => createTask(row));
    resolveDependencies(tasks);
    const swimLanes = groupIntoSwimLanes(tasks);

    const template = this.templateManager.getActiveTemplate();
    const layout = calculateLayout(swimLanes, template, config);

    return {
      success: true,
      phase: "preview",
      warnings: validation.warnings,
      layout,
      tasks,
      swimLanes,
    };
  }

  /**
   * Generate from pre-parsed rows (for data editor and MPP import).
   * Bypasses Excel parsing, goes straight to validate -> layout -> render.
   */
  async generateFromRows(rows, sourceName, config = {}) {
    resetIdCounter();

    const validation = validateRows(rows);

    if (!validation.isValid) {
      return {
        success: false,
        phase: "validation",
        errors: validation.errors,
        warnings: validation.warnings,
      };
    }

    const tasks = rows.map((row) => createTask(row));
    resolveDependencies(tasks);
    const swimLanes = groupIntoSwimLanes(tasks);

    const template = this.templateManager.getActiveTemplate();
    const layout = calculateLayout(swimLanes, template, config);

    await clearStreamlineShapes();
    await renderGantt(layout);

    this.lastLayout = layout;
    this.linkedFileName = sourceName;
    this._lastRows = rows;

    return {
      success: true,
      phase: "complete",
      warnings: validation.warnings,
      layout,
      tasks,
      stats: {
        sheetName: sourceName,
        totalTasks: tasks.length,
        swimLanes: swimLanes.length,
        milestones: tasks.filter((t) => t.type === "MILESTONE").length,
        taskBars: tasks.filter((t) => t.type === "TASK").length,
        dependencies: tasks.reduce(
          (sum, t) => sum + (t.dependencies ? t.dependencies.length : 0),
          0
        ),
        criticalPathLength: layout.criticalPathIds.size,
        scalingMode: layout.scaling.labelMode,
      },
    };
  }

  getLinkedFileName() {
    return this.linkedFileName;
  }

  hasLinkedFile() {
    return this.lastExcelData !== null || this._lastRows !== undefined;
  }
}

module.exports = { RefreshController };
