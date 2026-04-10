/**
 * Streamline Task Pane Controller
 * Wires UI events to the core engine. Runs inside the PowerPoint task pane web view.
 */

const { TemplateManager } = require("../core/templateManager");
const { RefreshController } = require("../core/refreshController");
const { parseClipboardData } = require("../core/excelParser");
const { parseMppXml, isMppXml, parseMppFile, isMppBinary } = require("../core/mppParser");
const { downloadMppXml } = require("../core/mppExporter");
const { DataEditor } = require("./dataEditor");
const { autoShift } = require("../core/autoScheduler");
const { WORKING_DAY_PRESETS, DEFAULT_WORKING_DAYS } = require("../core/workingDays");
const { KeyboardShortcutManager } = require("./keyboardShortcuts");
const { GraphClient } = require("../core/graphClient");
const {
  plannerToRows,
  todoToRows,
  calendarToRows,
  sharePointListToRows,
  classifyDriveItems,
} = require("../core/m365Importers");
const {
  createGantt: copilotCreateGantt,
  importFromM365: copilotImportFromM365,
  describeGantt: copilotDescribeGantt,
} = require("../copilot/agentActions");

// Conditionally require modules that need browser/Office context
let exportManager = null;
let shapeInteraction = null;
try { exportManager = require("../core/exportManager"); } catch (e) { /* not in browser */ }
try { shapeInteraction = require("../core/shapeInteraction"); } catch (e) { /* not in Office */ }

class TaskPaneController {
  constructor() {
    this.templateManager = new TemplateManager();
    this.refreshController = new RefreshController(this.templateManager);
    this.dataEditor = null;
    this.lastLayout = null;
    this.lastTasks = null;          // task objects from most recent generation
    this.selectedTask = null;       // currently selected task object
    this.selectedShapeTag = null;   // tag of currently selected shape
    this.positionTracker = null;    // ShapePositionTracker instance
    this.selectionWatcher = null;   // SelectionWatcher instance
    this.reverseMapper = null;      // X → Date function
    this._suppressNextPoll = false; // flag to ignore programmatic moves
    this.keyboardShortcuts = null;  // KeyboardShortcutManager
    this.workingDaysConfig = JSON.parse(JSON.stringify(DEFAULT_WORKING_DAYS));
    this.graphClient = new GraphClient();
    this.m365User = null;           // signed-in profile { displayName, mail, id }
  }

  init() {
    this.cacheElements();
    this.bindEvents();
    this.updateSliderLabels();
    this.initDataEditor();
    this.initTabs();
    this.initKeyboardShortcuts();
    this.initTextStyleButtons();
  }

  initKeyboardShortcuts() {
    this.keyboardShortcuts = new KeyboardShortcutManager(this);
    this.keyboardShortcuts.start();
  }

  initTextStyleButtons() {
    // Initialize from active template and wire click handlers
    this._updateTextStyleButtons();
    document.querySelectorAll(".ts-btn").forEach((btn) => {
      btn.addEventListener("click", (e) => {
        const element = e.currentTarget.dataset.element;
        const style = e.currentTarget.dataset.style;
        this._toggleTextStyle(element, style);
      });
    });
  }

  _updateTextStyleButtons() {
    const tmpl = this.templateManager.getActiveTemplate();
    const styles = (tmpl.fonts && tmpl.fonts.styles) || {};
    document.querySelectorAll(".ts-btn").forEach((btn) => {
      const element = btn.dataset.element;
      const style = btn.dataset.style;
      const val = (styles[element] || {})[style];
      btn.classList.toggle("active", !!val);
    });
  }

  _toggleTextStyle(element, style) {
    const tmpl = this.templateManager.getActiveTemplate();
    if (!tmpl.fonts.styles) tmpl.fonts.styles = {};
    if (!tmpl.fonts.styles[element]) tmpl.fonts.styles[element] = {};
    tmpl.fonts.styles[element][style] = !tmpl.fonts.styles[element][style];
    this._updateTextStyleButtons();
    if (this.refreshController.hasLinkedFile()) {
      this.showStatus("Text style changed. Click Refresh to apply.", "loading");
    }
  }

  cacheElements() {
    this.el = {
      // Import tab
      btnImport: document.getElementById("btn-import"),
      btnImportMpp: document.getElementById("btn-import-mpp"),
      btnPaste: document.getElementById("btn-paste"),
      btnRefresh: document.getElementById("btn-refresh"),
      fileInput: document.getElementById("file-input"),
      mppInput: document.getElementById("mpp-input"),
      btnExportPng: document.getElementById("btn-export-png"),
      btnExportJpg: document.getElementById("btn-export-jpg"),
      btnExportPdf: document.getElementById("btn-export-pdf"),
      btnExportMpp: document.getElementById("btn-export-mpp"),
      // M365 / Copilot
      m365SignedOut: document.getElementById("m365-signed-out"),
      m365SignedIn: document.getElementById("m365-signed-in"),
      btnM365SignIn: document.getElementById("btn-m365-signin"),
      btnM365SignOut: document.getElementById("btn-m365-signout"),
      m365UserName: document.getElementById("m365-user-name"),
      m365Source: document.getElementById("m365-source"),
      btnM365Import: document.getElementById("btn-m365-import"),
      copilotPrompt: document.getElementById("copilot-prompt"),
      btnCopilotGenerate: document.getElementById("btn-copilot-generate"),
      // Style tab
      templateCategory: document.getElementById("template-category"),
      templateSelect: document.getElementById("template-select"),
      btnImportTemplate: document.getElementById("btn-import-template"),
      btnExportTemplate: document.getElementById("btn-export-template"),
      templateInput: document.getElementById("template-input"),
      styleMode: document.getElementById("style-mode"),
      labelPosition: document.getElementById("label-position"),
      labelAlign: document.getElementById("label-align"),
      labelShowDates: document.getElementById("label-show-dates"),
      labelShowOwner: document.getElementById("label-show-owner"),
      optToday: document.getElementById("opt-today"),
      optElapsed: document.getElementById("opt-elapsed"),
      optDuration: document.getElementById("opt-duration"),
      optPercent: document.getElementById("opt-percent"),
      optBaselines: document.getElementById("opt-baselines"),
      // Settings tab
      renderWidth: document.getElementById("render-width"),
      renderWidthValue: document.getElementById("render-width-value"),
      sortBy: document.getElementById("sort-by"),
      timescaleTiers: document.getElementById("timescale-tiers"),
      timescaleGranularity: document.getElementById("timescale-granularity"),
      fiscalYear: document.getElementById("fiscal-year"),
      fyLabelFormat: document.getElementById("fy-label-format"),
      fyPrefix: document.getElementById("fy-prefix"),
      workingDaysPreset: document.getElementById("working-days-preset"),
      customDaysRow: document.getElementById("custom-days-row"),
      highlightWeekends: document.getElementById("highlight-weekends"),
      btnShortcuts: document.getElementById("btn-shortcuts"),
      shortcutsModal: document.getElementById("shortcuts-modal"),
      shortcutsList: document.getElementById("shortcuts-list"),
      btnCloseShortcuts: document.getElementById("btn-close-shortcuts"),
      // Editor tab
      editorContainer: document.getElementById("data-editor-container"),
      btnEditorGenerate: document.getElementById("btn-editor-generate"),
      btnEditorClear: document.getElementById("btn-editor-clear"),
      btnApplyShapeEdit: document.getElementById("btn-apply-shape-edit"),
      shapeEmptyState: document.getElementById("shape-empty-state"),
      shapeEditPanel: document.getElementById("shape-edit-panel"),
      editShapeTypeBadge: document.getElementById("edit-shape-type-badge"),
      editShapeId: document.getElementById("edit-shape-id"),
      editShapeName: document.getElementById("edit-shape-name"),
      editShapeStart: document.getElementById("edit-shape-start"),
      editShapeEnd: document.getElementById("edit-shape-end"),
      editEndRow: document.getElementById("edit-end-row"),
      editShapeStatus: document.getElementById("edit-shape-status"),
      editShapePct: document.getElementById("edit-shape-pct"),
      editShapePctValue: document.getElementById("edit-shape-pct-value"),
      editPctRow: document.getElementById("edit-pct-row"),
      editShapeColor: document.getElementById("edit-shape-color"),
      editShapeX: document.getElementById("edit-shape-x"),
      editShapeY: document.getElementById("edit-shape-y"),
      ctxRename: document.getElementById("ctx-rename"),
      ctxRecolor: document.getElementById("ctx-recolor"),
      ctxDelete: document.getElementById("ctx-delete"),
      optTrackShapes: document.getElementById("opt-track-shapes"),
      dragStatus: document.getElementById("drag-status"),
      // Common
      statusBar: document.getElementById("status-bar"),
      statusIcon: document.getElementById("status-icon"),
      statusMessage: document.getElementById("status-message"),
      fileInfo: document.getElementById("file-info"),
      fileName: document.getElementById("file-name"),
      statsPanel: document.getElementById("stats-panel"),
      statLanes: document.getElementById("stat-lanes"),
      statTasks: document.getElementById("stat-tasks"),
      statMilestones: document.getElementById("stat-milestones"),
      statDeps: document.getElementById("stat-deps"),
      errorPanel: document.getElementById("error-panel"),
      errorList: document.getElementById("error-list"),
      warningPanel: document.getElementById("warning-panel"),
      warningList: document.getElementById("warning-list"),
    };
  }

  initTabs() {
    const tabs = document.querySelectorAll(".tab");
    const contents = document.querySelectorAll(".tab-content");

    tabs.forEach((tab) => {
      tab.addEventListener("click", () => {
        tabs.forEach((t) => t.classList.remove("active"));
        contents.forEach((c) => c.classList.remove("active"));
        tab.classList.add("active");
        const target = document.getElementById(`tab-${tab.dataset.tab}`);
        if (target) target.classList.add("active");
      });
    });
  }

  initDataEditor() {
    if (!this.el.editorContainer) return;
    this.dataEditor = new DataEditor(this.el.editorContainer);
    this.dataEditor.init();
  }

  bindEvents() {
    // Import tab
    this.el.btnImport.addEventListener("click", () => this.el.fileInput.click());
    this.el.fileInput.addEventListener("change", (e) => this.handleFileImport(e));
    this.el.btnImportMpp.addEventListener("click", () => this.el.mppInput.click());
    this.el.mppInput.addEventListener("change", (e) => this.handleMppImport(e));
    this.el.btnPaste.addEventListener("click", () => this.handleClipboardPaste());
    this.el.btnRefresh.addEventListener("click", () => this.handleRefresh());
    this.el.btnExportPng.addEventListener("click", () => this.handleExportPng());
    this.el.btnExportJpg.addEventListener("click", () => this.handleExportJpg());
    this.el.btnExportPdf.addEventListener("click", () => this.handleExportPdf());
    this.el.btnExportMpp.addEventListener("click", () => this.handleExportMpp());

    // M365 sign-in / import
    this.el.btnM365SignIn.addEventListener("click", () => this.handleM365SignIn());
    this.el.btnM365SignOut.addEventListener("click", () => this.handleM365SignOut());
    this.el.btnM365Import.addEventListener("click", () => this.handleM365Import());

    // Copilot-style NL generation
    this.el.btnCopilotGenerate.addEventListener("click", () => this.handleCopilotGenerate());

    // Working days preset
    this.el.workingDaysPreset.addEventListener("change", (e) => this.handleWorkingDaysPreset(e.target.value));

    // Day toggle buttons
    document.querySelectorAll(".day-toggle").forEach((btn) => {
      btn.addEventListener("click", (e) => this.handleDayToggle(parseInt(e.currentTarget.dataset.day, 10)));
    });

    // Template category filter
    this.el.templateCategory.addEventListener("change", (e) => this.handleTemplateCategoryChange(e.target.value));

    // Keyboard shortcuts modal
    this.el.btnShortcuts.addEventListener("click", () => this.showKeyboardShortcuts());
    this.el.btnCloseShortcuts.addEventListener("click", () => this.hideKeyboardShortcuts());
    this.el.shortcutsModal.addEventListener("click", (e) => {
      if (e.target === this.el.shortcutsModal) this.hideKeyboardShortcuts();
    });

    // Style tab
    this.el.templateSelect.addEventListener("change", (e) => {
      this.templateManager.setActiveTemplate(e.target.value);
      this._updateTextStyleButtons();
      if (this.refreshController.hasLinkedFile()) {
        this.showStatus("Template changed. Click Refresh to apply.", "loading");
      }
    });
    this.el.btnImportTemplate.addEventListener("click", () => this.el.templateInput.click());
    this.el.templateInput.addEventListener("change", (e) => this.handleTemplateImport(e));
    this.el.btnExportTemplate.addEventListener("click", () => this.handleTemplateExport());

    // Settings tab
    this.el.renderWidth.addEventListener("input", () => this.updateSliderLabels());

    // Editor tab
    this.el.btnEditorGenerate.addEventListener("click", () => this.handleEditorGenerate());
    this.el.btnEditorClear.addEventListener("click", () => {
      if (this.dataEditor) {
        this.dataEditor.init([]);
        this.showStatus("Editor cleared.", "success");
      }
    });

    // Live shape editing
    this.el.btnApplyShapeEdit.addEventListener("click", () => this.handleApplyShapeEdit());
    this.el.editShapePct.addEventListener("input", (e) => {
      this.el.editShapePctValue.textContent = e.target.value + "%";
    });

    // Contextual actions
    this.el.ctxRename.addEventListener("click", () => this.el.editShapeName.focus());
    this.el.ctxRecolor.addEventListener("click", () => this.el.editShapeColor.click());
    this.el.ctxDelete.addEventListener("click", () => this.handleDeleteSelectedShape());

    // Shape tracking toggle
    this.el.optTrackShapes.addEventListener("change", (e) => {
      if (e.target.checked) {
        this.startShapeTracking();
      } else {
        this.stopShapeTracking();
      }
    });
  }

  // ═══════════════════════════════════════════════════════════
  // Shape Tracking (Selection Watcher + Position Tracker)
  // ═══════════════════════════════════════════════════════════

  startShapeTracking() {
    if (!shapeInteraction) return;
    if (!this.lastLayout) return;

    // Start selection watcher
    if (!this.selectionWatcher) {
      this.selectionWatcher = new shapeInteraction.SelectionWatcher();
    }
    this.selectionWatcher.start((shape) => this.handleSelectionChanged(shape));

    // Start position tracker
    if (!this.positionTracker) {
      this.positionTracker = new shapeInteraction.ShapePositionTracker();
    }
    this.positionTracker.start(
      (tag, type, id, oldPos, newPos) => this.handleShapeMoved(tag, type, id, oldPos, newPos),
      (tag, type, id, oldSize, newSize) => this.handleShapeResized(tag, type, id, oldSize, newSize)
    );

    // Build reverse mapper for X → Date conversion
    this.reverseMapper = shapeInteraction.createReverseMapper(
      this.lastLayout.ganttArea,
      this.lastLayout.dateRange
    );

    this.el.dragStatus.textContent = "Tracking active. Drag any task bar on the slide to update its dates.";
    this.el.dragStatus.classList.add("drag-active");
  }

  stopShapeTracking() {
    if (this.selectionWatcher) this.selectionWatcher.stop();
    if (this.positionTracker) this.positionTracker.stop();
    this.el.dragStatus.textContent = "Tracking disabled. Enable to auto-sync shape movements.";
    this.el.dragStatus.classList.remove("drag-active");
  }

  /**
   * Selection change: user clicked a Streamline shape on the slide.
   * Auto-populate the Live Shape Editor with that task's properties.
   */
  handleSelectionChanged(shape) {
    if (!shape) {
      this.showEmptyShapeState();
      return;
    }

    // Find the underlying task object
    const task = this.findTaskById(shape.id);
    if (!task) {
      this.showEmptyShapeState();
      return;
    }

    this.selectedTask = task;
    this.selectedShapeTag = shape.tag;

    // Populate edit panel
    this.el.shapeEmptyState.classList.add("hidden");
    this.el.shapeEditPanel.classList.remove("hidden");

    this.el.editShapeTypeBadge.textContent = task.type === "MILESTONE" ? "MILESTONE" : "TASK";
    this.el.editShapeTypeBadge.classList.toggle("milestone", task.type === "MILESTONE");
    this.el.editShapeId.textContent = task.id;
    this.el.editShapeName.value = task.name || "";
    this.el.editShapeStart.value = formatDateInput(task.startDate);

    if (task.type === "MILESTONE") {
      this.el.editEndRow.classList.add("hidden");
      this.el.editPctRow.classList.add("hidden");
    } else {
      this.el.editEndRow.classList.remove("hidden");
      this.el.editPctRow.classList.remove("hidden");
      this.el.editShapeEnd.value = formatDateInput(task.endDate);
      const pct = task.percentComplete !== null ? task.percentComplete : 0;
      this.el.editShapePct.value = pct;
      this.el.editShapePctValue.textContent = pct + "%";
    }

    this.el.editShapeStatus.value = statusKeyToLabel(task.status) || "";
    this.el.editShapeX.value = shape.left.toFixed(2);
    this.el.editShapeY.value = shape.top.toFixed(2);

    // Switch to Editor tab so user sees the panel
    this.switchToTab("editor");
  }

  /**
   * Shape moved by user (drag detected via polling).
   * Reverse-map the new X position to a date, update the task, cascade deps.
   */
  async handleShapeMoved(tag, type, id, oldPos, newPos) {
    if (this._suppressNextPoll) return;
    if (!this.reverseMapper) return;

    const task = this.findTaskById(id);
    if (!task) return;

    // Read the shape's full geometry (we only have left/top in oldPos)
    const shapes = await shapeInteraction.readAllShapePositions();
    const current = shapes.find((s) => s.tag === tag);
    if (!current) return;

    // Compute new dates
    const newDates = shapeInteraction.computeDatesFromPosition(
      type,
      { left: current.left, width: current.width },
      this.reverseMapper
    );

    const oldStart = task.startDate ? task.startDate.toISOString().slice(0, 10) : "?";
    const newStart = newDates.startDate ? newDates.startDate.toISOString().slice(0, 10) : "?";

    // Apply to task
    task.startDate = newDates.startDate;
    if (type !== "milestone") task.endDate = newDates.endDate;

    // Cascade via auto-shifter
    if (this.lastTasks) {
      const shifted = autoShift(this.lastTasks, task.id, newDates);
      const shiftedCount = shifted.size;

      this.showStatus(
        `${task.name}: ${oldStart} → ${newStart}${shiftedCount > 0 ? ` (+ ${shiftedCount} dependent${shiftedCount > 1 ? "s" : ""})` : ""}`,
        "success"
      );

      // If dependent tasks shifted, we need to re-render them
      if (shiftedCount > 0) {
        // Schedule a debounced refresh
        this._debouncedRefreshAfterDrag();
      }
    }

    // Update live editor panel if this is the selected shape
    if (this.selectedTask && this.selectedTask.id === id) {
      this.el.editShapeStart.value = formatDateInput(task.startDate);
      if (type !== "milestone") {
        this.el.editShapeEnd.value = formatDateInput(task.endDate);
      }
      this.el.editShapeX.value = current.left.toFixed(2);
      this.el.editShapeY.value = current.top.toFixed(2);
    }
  }

  /**
   * Shape resized by user — update task duration.
   */
  async handleShapeResized(tag, type, id, oldSize, newSize) {
    if (this._suppressNextPoll) return;
    if (type !== "taskbar") return; // only resize makes sense for task bars
    if (!this.reverseMapper) return;

    const task = this.findTaskById(id);
    if (!task) return;

    const shapes = await shapeInteraction.readAllShapePositions();
    const current = shapes.find((s) => s.tag === tag);
    if (!current) return;

    // New end date based on left + width
    task.endDate = this.reverseMapper(current.left + current.width);

    if (this.lastTasks) {
      const shifted = autoShift(this.lastTasks, task.id, {
        startDate: task.startDate,
        endDate: task.endDate,
      });

      this.showStatus(
        `${task.name} resized (duration: ${Math.round((task.endDate - task.startDate) / (1000 * 60 * 60 * 24))}d)${shifted.size > 0 ? ` (+ ${shifted.size} dependent${shifted.size > 1 ? "s" : ""})` : ""}`,
        "success"
      );

      if (shifted.size > 0) {
        this._debouncedRefreshAfterDrag();
      }
    }

    if (this.selectedTask && this.selectedTask.id === id) {
      this.el.editShapeEnd.value = formatDateInput(task.endDate);
    }
  }

  /**
   * Debounced full re-render after a drag — triggered when dependents cascade.
   * Waits 1.5s after the last drag event to avoid re-rendering mid-drag.
   */
  _debouncedRefreshAfterDrag() {
    if (this._dragRefreshTimer) clearTimeout(this._dragRefreshTimer);
    this._dragRefreshTimer = setTimeout(async () => {
      if (!this.lastTasks) return;
      // Suppress tracking during programmatic re-render
      this._suppressNextPoll = true;
      try {
        // Convert task objects back to rows for the generation pipeline
        const rows = this._tasksToRows(this.lastTasks);
        const result = await this.refreshController.generateFromRows(
          rows, "LiveEdit", this.getConfig()
        );
        if (result.success) {
          this.lastLayout = result.layout;
          this.reverseMapper = shapeInteraction.createReverseMapper(
            this.lastLayout.ganttArea,
            this.lastLayout.dateRange
          );
          // Reset the tracker baseline so it doesn't detect its own re-render
          if (this.positionTracker) {
            await this.positionTracker.resetSnapshot();
          }
        }
      } catch (err) {
        console.error("Drag refresh error:", err);
      } finally {
        this._suppressNextPoll = false;
      }
    }, 1500);
  }

  _tasksToRows(tasks) {
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

  findTaskById(id) {
    if (!this.lastTasks) return null;
    return this.lastTasks.find((t) => t.id === id) || null;
  }

  showEmptyShapeState() {
    this.el.shapeEmptyState.classList.remove("hidden");
    this.el.shapeEditPanel.classList.add("hidden");
    this.selectedTask = null;
    this.selectedShapeTag = null;
  }

  switchToTab(tabName) {
    document.querySelectorAll(".tab").forEach((t) => t.classList.remove("active"));
    document.querySelectorAll(".tab-content").forEach((c) => c.classList.remove("active"));
    const tab = document.querySelector(`.tab[data-tab="${tabName}"]`);
    const content = document.getElementById(`tab-${tabName}`);
    if (tab) tab.classList.add("active");
    if (content) content.classList.add("active");
  }

  updateSliderLabels() {
    this.el.renderWidthValue.textContent = this.el.renderWidth.value + "%";
  }

  getConfig() {
    // Apply label position/alignment overrides to the active template before generating
    this._applyLabelConfigToTemplate();

    return {
      renderWidthPercent: parseInt(this.el.renderWidth.value, 10) / 100,
      sortBy: this.el.sortBy.value,
      timescaleTiers: parseInt(this.el.timescaleTiers.value, 10),
      timescaleGranularity: this.el.timescaleGranularity.value,
      fiscalYearStartMonth: parseInt(this.el.fiscalYear.value, 10),
      fiscalYearLabelFormat: this.el.fyLabelFormat.value,
      fiscalYearPrefix: this.el.fyPrefix.value || "FY",
      showTodayMarker: this.el.optToday.checked,
      showElapsedShading: this.el.optElapsed.checked,
      showDurationLabels: this.el.optDuration.checked,
      showPercentComplete: this.el.optPercent.checked,
      showBaselines: this.el.optBaselines.checked,
      styleMode: this.el.styleMode.value,
      taskLabelPositionOverride: this.el.labelPosition.value,
      showWeekendHighlighting: this.el.highlightWeekends.checked,
      workingDays: this.workingDaysConfig,
    };
  }

  _applyLabelConfigToTemplate() {
    const tmpl = this.templateManager.getActiveTemplate();
    if (!tmpl.labelConfig) tmpl.labelConfig = {};
    tmpl.labelConfig.taskLabelPosition = this.el.labelPosition.value;
    tmpl.labelConfig.taskLabelAlign = this.el.labelAlign.value;
    tmpl.labelConfig.showTaskDates = this.el.labelShowDates.checked;
    tmpl.labelConfig.showOwner = this.el.labelShowOwner.checked;
  }

  async handleFileImport(event) {
    const file = event.target.files[0];
    if (!file) return;

    this.showStatus("Importing and generating...", "loading");
    this.clearPanels();

    try {
      const arrayBuffer = await file.arrayBuffer();
      const result = await this.refreshController.generate(
        arrayBuffer, file.name, this.getConfig()
      );

      if (result.success) {
        this.showFileInfo(file.name);
        this.el.btnRefresh.disabled = false;
        this.enableExport();
        this.showStats(result.stats);
        this.showWarnings(result.warnings);
        this.lastLayout = result.layout || null;
        this.lastTasks = result.tasks || null;
        // Load into data editor
        if (this.dataEditor && result.parsedRows) {
          this.dataEditor.loadRows(result.parsedRows);
        }
        // Start live shape tracking
        if (this.el.optTrackShapes.checked) {
          this.startShapeTracking();
        }
        this.showStatus(
          `Generated: ${result.stats.swimLanes} lanes, ${result.stats.totalTasks} items`,
          "success"
        );
      } else {
        this.showErrors(result.errors);
        this.showWarnings(result.warnings);
        this.showStatus("Validation failed. See issues below.", "error");
      }
    } catch (err) {
      this.showStatus("Error: " + err.message, "error");
      console.error("Streamline import error:", err);
    }

    event.target.value = "";
  }

  async handleMppImport(event) {
    const file = event.target.files[0];
    if (!file) return;

    this.showStatus("Importing MS Project file...", "loading");
    this.clearPanels();

    try {
      // Check if it's a binary .mpp file first
      const buffer = await file.arrayBuffer();
      if (isMppBinary(buffer)) {
        this.showStatus(
          "Binary .mpp files cannot be parsed directly. In MS Project, use File > Export > Save As XML, then re-import the .xml file here.",
          "error"
        );
        event.target.value = "";
        return;
      }

      // Parse as text/XML
      const text = new TextDecoder("utf-8").decode(buffer);
      if (!isMppXml(text)) {
        this.showStatus("Not a valid MS Project XML file. In MS Project: File > Export > Save As > XML Format.", "error");
        event.target.value = "";
        return;
      }

      const { rows, projectName } = parseMppXml(text);
      if (rows.length === 0) {
        this.showStatus("No tasks found in XML file.", "error");
        return;
      }

      // Load into data editor and generate
      if (this.dataEditor) this.dataEditor.loadRows(rows);

      // Process through the pipeline
      const result = await this.refreshController.generateFromRows(
        rows, projectName, this.getConfig()
      );

      if (result.success) {
        this.showFileInfo(projectName + ".xml");
        this.el.btnRefresh.disabled = false;
        this.enableExport();
        this.showStats(result.stats);
        this.lastLayout = result.layout || null;
        this.lastTasks = result.tasks || null;
        if (this.el.optTrackShapes.checked) this.startShapeTracking();
        this.showStatus(
          `Imported: ${rows.length} tasks from ${projectName}`,
          "success"
        );
      } else {
        this.showErrors(result.errors);
        this.showStatus("Validation failed. See issues below.", "error");
      }
    } catch (err) {
      this.showStatus("Import error: " + err.message, "error");
      console.error("MPP import error:", err);
    }

    event.target.value = "";
  }

  async handleClipboardPaste() {
    try {
      const text = await navigator.clipboard.readText();
      if (!text || !text.trim()) {
        this.showStatus("Clipboard is empty.", "error");
        return;
      }

      const result = parseClipboardData(text);
      if (result.rows.length === 0) {
        this.showStatus("No valid data found in clipboard. Copy tabular data with headers.", "error");
        return;
      }

      // Load into data editor
      if (this.dataEditor) this.dataEditor.loadRows(result.rows);

      this.showStatus(
        `Pasted ${result.rows.length} rows (${result.mappedColumns.length} columns mapped). Switch to Editor tab to review, then Generate.`,
        "success"
      );
    } catch (err) {
      this.showStatus("Clipboard access denied. Try using Import Excel instead.", "error");
    }
  }

  async handleEditorGenerate() {
    if (!this.dataEditor) return;

    const rows = this.dataEditor.getRows();
    if (rows.length === 0) {
      this.showStatus("No data in editor. Add some tasks first.", "error");
      return;
    }

    this.showStatus("Generating from editor data...", "loading");
    this.clearPanels();

    try {
      const result = await this.refreshController.generateFromRows(
        rows, "Editor", this.getConfig()
      );

      if (result.success) {
        this.showFileInfo("Editor (manual)");
        this.el.btnRefresh.disabled = false;
        this.enableExport();
        this.showStats(result.stats);
        this.showWarnings(result.warnings);
        this.lastLayout = result.layout || null;
        this.lastTasks = result.tasks || null;
        if (this.el.optTrackShapes.checked) this.startShapeTracking();
        this.showStatus(
          `Generated: ${result.stats.swimLanes} lanes, ${result.stats.totalTasks} items`,
          "success"
        );
      } else {
        this.showErrors(result.errors);
        this.showWarnings(result.warnings);
        this.showStatus("Validation failed. See issues below.", "error");
      }
    } catch (err) {
      this.showStatus("Error: " + err.message, "error");
    }
  }

  async handleRefresh() {
    if (!this.refreshController.hasLinkedFile()) {
      this.showStatus("No file linked. Import data first.", "error");
      return;
    }

    const confirmed = await this.showConfirmation(
      "This will regenerate the Gantt chart. Any manual edits to Streamline shapes will be overwritten. Continue?"
    );
    if (!confirmed) return;

    this.showStatus("Refreshing timeline...", "loading");
    this.clearPanels();

    try {
      const result = await this.refreshController.refresh(null, this.getConfig());

      if (result.success) {
        this.showStats(result.stats);
        this.showWarnings(result.warnings);
        this.lastLayout = result.layout || null;
        this.lastTasks = result.tasks || null;
        // Reset tracker snapshot after re-render
        if (this.positionTracker && this.positionTracker.isRunning()) {
          this.reverseMapper = shapeInteraction.createReverseMapper(
            this.lastLayout.ganttArea, this.lastLayout.dateRange
          );
          await this.positionTracker.resetSnapshot();
        } else if (this.el.optTrackShapes.checked) {
          this.startShapeTracking();
        }
        this.showStatus(
          `Refreshed: ${result.stats.swimLanes} lanes, ${result.stats.totalTasks} items`,
          "success"
        );
      } else {
        this.showErrors(result.errors);
        this.showWarnings(result.warnings);
        this.showStatus("Refresh failed. See issues below.", "error");
      }
    } catch (err) {
      this.showStatus("Error: " + err.message, "error");
      console.error("Streamline refresh error:", err);
    }
  }

  handleExportPng() {
    if (!this.lastLayout || !exportManager) {
      this.showStatus("Generate a chart first.", "error");
      return;
    }
    try {
      const template = this.templateManager.getActiveTemplate();
      exportManager.downloadPNG(this.lastLayout, template);
      this.showStatus("PNG exported.", "success");
    } catch (err) {
      this.showStatus("Export failed: " + err.message, "error");
    }
  }

  handleExportJpg() {
    if (!this.lastLayout || !exportManager) {
      this.showStatus("Generate a chart first.", "error");
      return;
    }
    try {
      const template = this.templateManager.getActiveTemplate();
      exportManager.downloadJPG(this.lastLayout, template);
      this.showStatus("JPG exported.", "success");
    } catch (err) {
      this.showStatus("Export failed: " + err.message, "error");
    }
  }

  handleExportMpp() {
    if (!this.lastTasks) {
      this.showStatus("Generate a chart first.", "error");
      return;
    }
    try {
      const { groupIntoSwimLanes } = require("../core/dataModel");
      const swimLanes = groupIntoSwimLanes(this.lastTasks);
      downloadMppXml(this.lastTasks, swimLanes, "Streamline_Export");
      this.showStatus("MS Project XML exported. Open it in MS Project to sync back.", "success");
    } catch (err) {
      this.showStatus("MS Project export failed: " + err.message, "error");
    }
  }

  handleExportPdf() {
    if (!this.lastLayout || !exportManager) {
      this.showStatus("Generate a chart first.", "error");
      return;
    }
    try {
      const template = this.templateManager.getActiveTemplate();
      exportManager.downloadPDF(this.lastLayout, template);
      this.showStatus("PDF export opened.", "success");
    } catch (err) {
      this.showStatus("Export failed: " + err.message, "error");
    }
  }

  // ═══════════════════════════════════════════════════════════
  // Microsoft 365 / Graph integration
  // ═══════════════════════════════════════════════════════════

  /**
   * Sign in to M365 via Office SSO. Acquires an access token through
   * Office.auth.getAccessToken() and caches it on the GraphClient. On
   * success, loads the user's display name and swaps the UI to the
   * signed-in state.
   */
  async handleM365SignIn() {
    this.showStatus("Signing in to Microsoft 365...", "loading");
    try {
      await this.graphClient.acquireTokenViaOfficeSSO();
      const me = await this.graphClient.getMe();
      this.m365User = me;
      this.el.m365UserName.textContent = me.displayName || me.userPrincipalName || "Signed in";
      this.el.m365SignedOut.classList.add("hidden");
      this.el.m365SignedIn.classList.remove("hidden");
      this.showStatus(`Signed in as ${me.displayName || me.userPrincipalName}.`, "success");
    } catch (err) {
      this.showStatus(`Sign-in failed: ${err.message}`, "error");
    }
  }

  handleM365SignOut() {
    this.graphClient = new GraphClient();
    this.m365User = null;
    this.el.m365SignedOut.classList.remove("hidden");
    this.el.m365SignedIn.classList.add("hidden");
    this.showStatus("Signed out of Microsoft 365.", "success");
  }

  /**
   * Import tasks from the selected M365 source and render them. Dispatches
   * to the appropriate Graph endpoint based on the source dropdown.
   */
  async handleM365Import() {
    if (!this.graphClient.hasAccessToken()) {
      this.showStatus("Sign in to Microsoft 365 first.", "error");
      return;
    }
    const source = this.el.m365Source.value;
    this.showStatus(`Importing from ${source}...`, "loading");
    this.clearPanels();

    try {
      const context = {
        refreshController: this.refreshController,
        templateManager: this.templateManager,
        graphClient: this.graphClient,
        config: this.getConfig(),
      };

      // For picker-based sources (Planner plan, To Do list, OneDrive file,
      // SharePoint site+list) we prompt the user to choose. For simplicity
      // the first step lists their options and picks the top entry; a
      // fuller UI would render these as selectable cards.
      let req;
      if (source === "planner") {
        const plans = await this.graphClient.getMyPlans();
        if (plans.length === 0) { this.showStatus("No Planner plans found.", "error"); return; }
        req = { source, planId: plans[0].id, templateKey: this.el.templateSelect.value };
      } else if (source === "todo") {
        const lists = await this.graphClient.getTodoLists();
        if (lists.length === 0) { this.showStatus("No To Do lists found.", "error"); return; }
        req = { source, listId: lists[0].id, listName: lists[0].displayName };
      } else if (source === "calendar") {
        req = {
          source,
          fromDate: new Date().toISOString(),
          toDate: new Date(Date.now() + 90 * 86400e3).toISOString(),
        };
      } else if (source === "onedrive") {
        const children = await this.graphClient.getOneDriveChildren("Streamline");
        const compatible = classifyDriveItems(children);
        if (compatible.length === 0) {
          this.showStatus("No Streamline-compatible files in your OneDrive /Streamline folder.", "error");
          return;
        }
        req = { source, driveItemId: compatible[0].id, fileName: compatible[0].name };
      } else if (source === "sharepoint") {
        this.showStatus("SharePoint import requires a site + list picker (future UI).", "error");
        return;
      } else {
        this.showStatus(`Unknown source: ${source}`, "error");
        return;
      }

      const summary = await copilotImportFromM365(req, context);
      this.lastLayout = this.refreshController.lastLayout;
      // Stash recent tasks for describe/update calls
      this.lastTasks = context.lastTasks || this.lastTasks;
      this.showFileInfo(`${source}: ${summary.projectName}`);
      this.el.btnRefresh.disabled = false;
      this.enableExport();
      this.showStats({
        swimLanes: summary.swimLaneCount,
        taskBars: summary.taskCount,
        milestones: summary.milestoneCount,
        dependencies: summary.dependencyCount,
      });
      this.showStatus(
        `Imported ${summary.taskCount} tasks and ${summary.milestoneCount} milestones from ${source}.`,
        "success"
      );
    } catch (err) {
      this.showStatus(`M365 import failed: ${err.message}`, "error");
      console.error("M365 import error:", err);
    }
  }

  /**
   * Generate a Gantt from a Copilot-style natural-language prompt.
   * Calls the same agentActions.createGantt used by the declarative agent
   * so results are identical regardless of invocation source. The actual
   * NL parsing lives in a lightweight server endpoint in production;
   * here we do a best-effort local parse using the same regex parser that
   * the Teams message extension uses.
   */
  async handleCopilotGenerate() {
    const prompt = (this.el.copilotPrompt.value || "").trim();
    if (!prompt) {
      this.showStatus("Enter a project description first.", "error");
      return;
    }
    this.showStatus("Drafting Gantt with Copilot...", "loading");
    try {
      const { extractTasksFromText } = require("../copilot/messageExtension");
      const tasks = extractTasksFromText(prompt);
      if (tasks.length === 0) {
        this.showStatus(
          "Couldn't extract tasks. Use lines like \"Task name - 2026-04-15 to 2026-05-01\" or \"## Swim lane\".",
          "error"
        );
        return;
      }
      const context = {
        refreshController: this.refreshController,
        templateManager: this.templateManager,
        config: this.getConfig(),
      };
      const summary = await copilotCreateGantt(
        { tasks, projectName: "Copilot Draft" },
        context
      );
      this.lastLayout = this.refreshController.lastLayout;
      this.el.btnRefresh.disabled = false;
      this.enableExport();
      this.showStats({
        swimLanes: summary.swimLaneCount,
        taskBars: summary.taskCount,
        milestones: summary.milestoneCount,
        dependencies: summary.dependencyCount,
      });
      this.showFileInfo(summary.projectName);
      this.showStatus(
        `Generated ${summary.taskCount} tasks in ${summary.swimLaneCount} lanes.`,
        "success"
      );
    } catch (err) {
      this.showStatus(`Copilot generation failed: ${err.message}`, "error");
    }
  }

  async handleApplyShapeEdit() {
    if (!shapeInteraction || !this.selectedShapeTag || !this.selectedTask) return;

    try {
      this._suppressNextPoll = true;

      // Apply edits to the task object
      const task = this.selectedTask;
      const newName = this.el.editShapeName.value.trim();
      const newStart = this.el.editShapeStart.value;
      const newEnd = this.el.editShapeEnd.value;
      const newStatus = this.el.editShapeStatus.value;
      const newPct = parseFloat(this.el.editShapePct.value);
      const newColor = this.el.editShapeColor.value;
      const newX = parseFloat(this.el.editShapeX.value);
      const newY = parseFloat(this.el.editShapeY.value);

      if (newName) task.name = newName;
      if (newStart) task.startDate = new Date(newStart);
      if (newEnd && task.type !== "MILESTONE") task.endDate = new Date(newEnd);
      if (newStatus) task.status = labelToStatusKey(newStatus);
      if (!isNaN(newPct)) task.percentComplete = newPct;

      // Move + recolor the shape directly for instant feedback
      if (!isNaN(newX) && !isNaN(newY)) {
        await shapeInteraction.moveShape(this.selectedShapeTag, newX, newY);
      }
      if (newColor) {
        await shapeInteraction.updateShapeColor(this.selectedShapeTag, newColor);
      }

      // Cascade dependency updates
      if (task.startDate && this.lastTasks) {
        const shifted = autoShift(this.lastTasks, task.id, {
          startDate: task.startDate,
          endDate: task.endDate,
        });

        if (shifted.size > 0) {
          this._debouncedRefreshAfterDrag();
          this.showStatus(`Updated ${task.name} (+ ${shifted.size} cascaded)`, "success");
        } else {
          this.showStatus(`Updated ${task.name}`, "success");
        }
      } else {
        this.showStatus(`Updated ${task.name}`, "success");
      }

      // Reset snapshot so tracker doesn't detect its own move
      if (this.positionTracker) {
        setTimeout(() => this.positionTracker.resetSnapshot(), 200);
      }
    } catch (err) {
      this.showStatus("Update failed: " + err.message, "error");
    } finally {
      setTimeout(() => { this._suppressNextPoll = false; }, 300);
    }
  }

  async handleDeleteSelectedShape() {
    if (!this.selectedTask || !shapeInteraction) return;

    const confirmed = await this.showConfirmation(
      `Delete task "${this.selectedTask.name}"? This removes it from the slide and data model.`
    );
    if (!confirmed) return;

    try {
      this._suppressNextPoll = true;

      // Delete all shapes associated with this task
      await shapeInteraction.deleteShapeGroup(this.selectedTask.id);

      // Remove from the task list
      if (this.lastTasks) {
        this.lastTasks = this.lastTasks.filter((t) => t.id !== this.selectedTask.id);
      }

      this.showStatus(`Deleted ${this.selectedTask.name}`, "success");
      this.showEmptyShapeState();

      if (this.positionTracker) {
        await this.positionTracker.resetSnapshot();
      }
    } catch (err) {
      this.showStatus("Delete failed: " + err.message, "error");
    } finally {
      setTimeout(() => { this._suppressNextPoll = false; }, 300);
    }
  }

  async handleTemplateImport(event) {
    const file = event.target.files[0];
    if (!file) return;

    try {
      const text = await file.text();
      const json = JSON.parse(text);
      const key = file.name.replace(/\.json$/, "").replace(/\s+/g, "_").toLowerCase();

      this.templateManager.importTemplate(key, json);

      const option = document.createElement("option");
      option.value = key;
      option.textContent = json.name || key;
      this.el.templateSelect.appendChild(option);
      this.el.templateSelect.value = key;
      this.templateManager.setActiveTemplate(key);

      this.showStatus(`Theme "${json.name}" imported.`, "success");
    } catch (err) {
      this.showStatus("Invalid theme file: " + err.message, "error");
    }

    event.target.value = "";
  }

  handleTemplateExport() {
    try {
      const key = this.el.templateSelect.value;
      const tmpl = this.templateManager.exportTemplate(key);
      const json = JSON.stringify(tmpl, null, 2);
      const blob = new Blob([json], { type: "application/json" });
      const url = URL.createObjectURL(blob);

      const a = document.createElement("a");
      a.href = url;
      a.download = `${key}_theme.json`;
      a.click();

      URL.revokeObjectURL(url);
      this.showStatus(`Theme "${tmpl.name}" exported.`, "success");
    } catch (err) {
      this.showStatus("Export failed: " + err.message, "error");
    }
  }

  enableExport() {
    this.el.btnExportPng.disabled = false;
    this.el.btnExportJpg.disabled = false;
    this.el.btnExportPdf.disabled = false;
    this.el.btnExportMpp.disabled = false;
  }

  // ═══════════════════════════════════════════════════════════
  // Working Days
  // ═══════════════════════════════════════════════════════════

  handleWorkingDaysPreset(preset) {
    if (preset === "custom") {
      this.el.customDaysRow.classList.remove("hidden");
      return;
    }
    const p = WORKING_DAY_PRESETS[preset];
    if (!p) return;
    this.workingDaysConfig.days = [...p.days];
    document.querySelectorAll(".day-toggle").forEach((btn) => {
      const day = parseInt(btn.dataset.day, 10);
      btn.classList.toggle("active", p.days[day]);
    });
    this.el.customDaysRow.classList.add("hidden");
    if (this.refreshController.hasLinkedFile()) {
      this.showStatus("Working days changed. Click Refresh to apply.", "loading");
    }
  }

  handleDayToggle(day) {
    this.workingDaysConfig.days[day] = !this.workingDaysConfig.days[day];
    const btn = document.querySelector(`.day-toggle[data-day="${day}"]`);
    if (btn) btn.classList.toggle("active", this.workingDaysConfig.days[day]);
    this.el.workingDaysPreset.value = "custom";
    this.el.customDaysRow.classList.remove("hidden");
    if (this.refreshController.hasLinkedFile()) {
      this.showStatus("Working days changed. Click Refresh to apply.", "loading");
    }
  }

  // ═══════════════════════════════════════════════════════════
  // Template Category Filter
  // ═══════════════════════════════════════════════════════════

  handleTemplateCategoryChange(category) {
    const optgroups = this.el.templateSelect.querySelectorAll("optgroup");
    optgroups.forEach((og) => {
      const label = og.label.toLowerCase().replace(/\s+/g, "-");
      const shouldShow = category === "all" || label === category;
      og.style.display = shouldShow ? "" : "none";
    });
    const currentOption = this.el.templateSelect.selectedOptions[0];
    if (currentOption && currentOption.parentElement && currentOption.parentElement.style.display === "none") {
      const firstVisible = this.el.templateSelect.querySelector("optgroup:not([style*='none']) option");
      if (firstVisible) {
        this.el.templateSelect.value = firstVisible.value;
        this.templateManager.setActiveTemplate(firstVisible.value);
        this._updateTextStyleButtons();
      }
    }
  }

  // ═══════════════════════════════════════════════════════════
  // Keyboard Shortcuts Modal
  // ═══════════════════════════════════════════════════════════

  showKeyboardShortcuts() {
    if (!this.keyboardShortcuts) return;
    const shortcuts = this.keyboardShortcuts.getShortcutList();
    this.el.shortcutsList.innerHTML = shortcuts.map((s) => `
      <div class="shortcut-entry">
        <span class="shortcut-desc">${escapeHtml(s.description)}</span>
        <span class="shortcut-keys">${escapeHtml(s.keys)}</span>
      </div>
    `).join("");
    this.el.shortcutsModal.classList.remove("hidden");
  }

  hideKeyboardShortcuts() {
    this.el.shortcutsModal.classList.add("hidden");
  }

  /**
   * Dispatch action from keyboard shortcut.
   */
  handleKeyboardShortcut(action) {
    switch (action) {
      case "import": this.el.btnImport.click(); break;
      case "importMpp": this.el.btnImportMpp.click(); break;
      case "paste": this.handleClipboardPaste(); break;
      case "refresh": this.handleRefresh(); break;
      case "exportPng": this.handleExportPng(); break;
      case "exportJpg": this.handleExportJpg(); break;
      case "exportPdf": this.handleExportPdf(); break;
      case "exportMpp": this.handleExportMpp(); break;
      case "tabImport": this.switchToTab("import"); break;
      case "tabEditor": this.switchToTab("editor"); break;
      case "tabStyle": this.switchToTab("style"); break;
      case "tabSettings": this.switchToTab("settings"); break;
      case "deleteShape":
        if (this.selectedTask) this.handleDeleteSelectedShape();
        break;
      case "applyShapeEdit":
        if (this.selectedTask) this.handleApplyShapeEdit();
        break;
      case "showShortcuts": this.showKeyboardShortcuts(); break;
      case "newRow":
        if (this.dataEditor) {
          this.dataEditor.addRow();
          this.switchToTab("editor");
        }
        break;
    }
  }

  // ── UI Helpers ──

  showStatus(message, type) {
    this.el.statusBar.classList.remove("hidden", "success", "error", "loading");
    this.el.statusBar.classList.add(type);

    if (type === "loading") {
      this.el.statusIcon.innerHTML = '<span class="spinner"></span>';
    } else if (type === "success") {
      this.el.statusIcon.textContent = "\u2713";
    } else {
      this.el.statusIcon.textContent = "\u2717";
    }

    this.el.statusMessage.textContent = message;
  }

  showFileInfo(name) {
    this.el.fileInfo.classList.remove("hidden");
    this.el.fileName.textContent = name;
  }

  showStats(stats) {
    this.el.statsPanel.classList.remove("hidden");
    this.el.statLanes.textContent = stats.swimLanes;
    this.el.statTasks.textContent = stats.taskBars;
    this.el.statMilestones.textContent = stats.milestones;
    this.el.statDeps.textContent = stats.dependencies;
  }

  showErrors(errors) {
    if (!errors || errors.length === 0) return;
    this.el.errorPanel.classList.remove("hidden");
    this.el.errorList.innerHTML = errors.map((e) => `<li>${escapeHtml(e)}</li>`).join("");
  }

  showWarnings(warnings) {
    if (!warnings || warnings.length === 0) return;
    this.el.warningPanel.classList.remove("hidden");
    this.el.warningList.innerHTML = warnings.map((w) => `<li>${escapeHtml(w)}</li>`).join("");
  }

  clearPanels() {
    this.el.statsPanel.classList.add("hidden");
    this.el.errorPanel.classList.add("hidden");
    this.el.warningPanel.classList.add("hidden");
    this.el.errorList.innerHTML = "";
    this.el.warningList.innerHTML = "";
  }

  showConfirmation(message) {
    return new Promise((resolve) => {
      const overlay = document.createElement("div");
      overlay.className = "confirm-overlay";
      overlay.innerHTML = `
        <div class="confirm-dialog">
          <p>${escapeHtml(message)}</p>
          <div class="confirm-actions">
            <button class="btn btn-ghost" id="confirm-cancel">Cancel</button>
            <button class="btn btn-primary" id="confirm-ok">Continue</button>
          </div>
        </div>
      `;
      document.body.appendChild(overlay);

      overlay.querySelector("#confirm-ok").addEventListener("click", () => {
        document.body.removeChild(overlay);
        resolve(true);
      });
      overlay.querySelector("#confirm-cancel").addEventListener("click", () => {
        document.body.removeChild(overlay);
        resolve(false);
      });
    });
  }
}

function escapeHtml(text) {
  const div = document.createElement("div");
  div.textContent = text;
  return div.innerHTML;
}

function formatDateInput(date) {
  if (!date || !(date instanceof Date)) return "";
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function statusKeyToLabel(status) {
  const map = {
    ON_TRACK: "On Track",
    AT_RISK: "At Risk",
    DELAYED: "Delayed",
    COMPLETE: "Complete",
  };
  return map[status] || "";
}

function labelToStatusKey(label) {
  const map = {
    "On Track": "ON_TRACK",
    "At Risk": "AT_RISK",
    "Delayed": "DELAYED",
    "Complete": "COMPLETE",
  };
  return map[label] || null;
}

module.exports = { TaskPaneController };
