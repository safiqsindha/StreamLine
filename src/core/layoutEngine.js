/**
 * Streamline Layout Engine
 * 5-phase algorithm that calculates pixel positions for all Gantt chart elements
 * within PowerPoint slide constraints.
 *
 * Slide: 13.33" x 7.5" (landscape 16:9)
 * PowerPoint uses inches internally. Office JS API accepts inches.
 */

const { TaskType, DepType, getTaskDuration, calculateVariance } = require("./dataModel");
const { DEFAULT_WORKING_DAYS, getWeekendRegions } = require("./workingDays");

// Slide constraints (inches)
const SLIDE = {
  width: 13.33,
  height: 7.5,
};

// Default rendering area config
const DEFAULTS = {
  renderWidthPercent: 0.75,
  topMargin: 0.8,
  bottomMargin: 0.2,
  leftMargin: 0.15,
  rightMargin: 0.15,
  swimLaneLabelWidth: 1.5,
  // 3-tier timescale heights
  tier1Height: 0.25, // Top: years or quarters
  tier2Height: 0.25, // Middle: months
  tier3Height: 0.20, // Bottom: weeks or days
  datePaddingPercent: 0.05,
  fiscalYearStartMonth: 1, // 1=Jan, 4=Apr, 7=Jul, 10=Oct
  showTodayMarker: true,
  showElapsedShading: true,
  showDurationLabels: true,
  showBaselines: true,
  showPercentComplete: true,
  sortBy: "default", // "default" | "startDate" | "endDate" | "name" | "status"
  timescaleTiers: 3, // 1, 2, or 3
  timescaleGranularity: "auto", // "auto" | "hours" | "days" | "weeks" | "months"
  styleMode: "flat", // "flat" | "3d"
  // Fiscal year labeling
  fiscalYearLabelFormat: "end", // "end" (FY2026), "start" (FY2025), "both" (FY2025/26)
  fiscalYearPrefix: "FY",
  // Working days & weekend highlighting
  workingDays: DEFAULT_WORKING_DAYS,
  showWeekendHighlighting: false,
  // Label position overrides (also available per-template)
  taskLabelPositionOverride: null, // null = use template, or "inside"/"above"/"below"/"left"/"right"
};

// Dynamic scaling rules
const SCALING_RULES = [
  { maxLanes: 5, fontSize: 11, barHeight: 0.28, laneGap: 0.35, labelMode: "full" },
  { maxLanes: 10, fontSize: 10, barHeight: 0.24, laneGap: 0.22, labelMode: "full" },
  { maxLanes: 15, fontSize: 8, barHeight: 0.18, laneGap: 0.12, labelMode: "abbreviated" },
  { maxLanes: Infinity, fontSize: 7, barHeight: 0.14, laneGap: 0.06, labelMode: "truncated" },
];

function getScalingRule(laneCount) {
  for (const rule of SCALING_RULES) {
    if (laneCount <= rule.maxLanes) return rule;
  }
  return SCALING_RULES[SCALING_RULES.length - 1];
}

/**
 * Main layout function.
 */
function calculateLayout(swimLanes, template, config = {}) {
  const cfg = { ...DEFAULTS, ...config };

  // Sort tasks within lanes if requested
  if (cfg.sortBy !== "default") {
    for (const lane of swimLanes) {
      sortTasks(lane.tasks, cfg.sortBy);
      for (const sub of lane.subLanes || []) {
        sortTasks(sub.tasks, cfg.sortBy);
      }
    }
  }

  // Calculate rendering area
  const totalRenderWidth = SLIDE.width * cfg.renderWidthPercent;
  const ganttLeft = cfg.leftMargin + cfg.swimLaneLabelWidth;
  const ganttWidth = totalRenderWidth - cfg.leftMargin - cfg.rightMargin - cfg.swimLaneLabelWidth;

  const tiersUsed = Math.min(Math.max(cfg.timescaleTiers, 1), 3);
  let timeAxisTotalHeight = cfg.tier1Height;
  if (tiersUsed >= 2) timeAxisTotalHeight += cfg.tier2Height;
  if (tiersUsed >= 3) timeAxisTotalHeight += cfg.tier3Height;

  const ganttTop = cfg.topMargin + timeAxisTotalHeight;
  const availableHeight = SLIDE.height - cfg.topMargin - timeAxisTotalHeight - cfg.bottomMargin;

  const allTasks = swimLanes.flatMap((lane) => lane.tasks);

  // ── Phase 1: Date Range ──
  const dateRange = calculateDateRange(allTasks, cfg.datePaddingPercent);

  // ── Phase 2: Horizontal Mapping ──
  const mapDateToX = createDateMapper(dateRange, ganttLeft, ganttWidth);

  // ── Phase 3: Swim Lane Allocation ──
  const scaling = getScalingRule(swimLanes.length);
  const laneLayouts = allocateSwimLanes(swimLanes, ganttTop, availableHeight, scaling);

  // ── Phase 4: Task Positioning ──
  const taskElements = positionTasks(laneLayouts, mapDateToX, dateRange, scaling, template, cfg);

  // ── Phase 5: Dependency Routing ──
  const taskPositionMap = buildTaskPositionMap(taskElements);
  const dependencyLines = routeDependencies(allTasks, taskPositionMap, template);

  // Critical path
  const criticalPathIds = findCriticalPath(allTasks);
  for (const line of dependencyLines) {
    if (criticalPathIds.has(line.fromTaskId) && criticalPathIds.has(line.toTaskId)) {
      line.isCriticalPath = true;
      line.color = template.colors.criticalPath;
      line.weight = template.shapes.criticalPathLineWeight;
    }
  }

  // 3-tier time axis
  const timeAxisElements = buildTimeAxis(dateRange, mapDateToX, cfg, template, tiersUsed);

  // Today marker
  let todayMarker = null;
  if (cfg.showTodayMarker) {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    if (today >= dateRange.min && today <= dateRange.max) {
      const todayX = mapDateToX(today);
      todayMarker = {
        type: "todayMarker",
        x: todayX,
        top: ganttTop,
        bottom: SLIDE.height - cfg.bottomMargin,
        color: template.colors.todayMarker || "#FF0000",
        label: "Today",
        labelTop: cfg.topMargin + timeAxisTotalHeight - 0.15,
      };
    }
  }

  // Weekend highlighting regions
  let weekendShading = [];
  if (cfg.showWeekendHighlighting) {
    const regions = getWeekendRegions(dateRange.min, dateRange.max, cfg.workingDays);
    for (const region of regions) {
      const left = mapDateToX(region.startDate);
      const right = mapDateToX(region.endDate);
      if (right <= left) continue;
      weekendShading.push({
        type: "weekendShading",
        left: Math.max(left, ganttLeft),
        top: ganttTop,
        width: Math.min(right, ganttLeft + ganttWidth) - Math.max(left, ganttLeft),
        height: availableHeight,
        color: cfg.workingDays.weekendColor || template.colors.weekendShading || "#F5F5F5",
      });
    }
  }

  // Elapsed time shading
  let elapsedShading = null;
  if (cfg.showElapsedShading) {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    if (today > dateRange.min) {
      const shadingRight = Math.min(mapDateToX(today), ganttLeft + ganttWidth);
      elapsedShading = {
        type: "elapsedShading",
        left: ganttLeft,
        top: ganttTop,
        width: shadingRight - ganttLeft,
        height: availableHeight,
        color: template.colors.elapsedShading || "#F0F0F0",
      };
    }
  }

  // Swim lane labels + separators
  const laneLabelElements = [];
  const laneSeparators = [];
  const subLaneLabelElements = [];

  const tStyles = (template.fonts && template.fonts.styles) || {};

  laneLayouts.forEach((lane, i) => {
    laneLabelElements.push({
      type: "swimLaneLabel",
      id: lane.id,
      text: formatLaneLabel(lane.name, scaling.labelMode),
      fullText: lane.name,
      left: cfg.leftMargin,
      top: lane.yOffset,
      width: cfg.swimLaneLabelWidth,
      height: lane.height,
      fontSize: scaling.fontSize,
      bgColor: template.colors.swimLaneHeader,
      textColor: template.colors.swimLaneHeaderText,
      font: template.fonts.primary,
      bold: (tStyles.swimLaneLabel || {}).bold !== undefined ? tStyles.swimLaneLabel.bold : true,
      italic: (tStyles.swimLaneLabel || {}).italic,
      underline: (tStyles.swimLaneLabel || {}).underline,
    });

    // Sub-swimlane labels
    if (lane.subLaneLayouts) {
      for (const sub of lane.subLaneLayouts) {
        subLaneLabelElements.push({
          type: "subSwimLaneLabel",
          id: sub.id,
          text: sub.name,
          left: cfg.leftMargin + 0.15,
          top: sub.yOffset,
          width: cfg.swimLaneLabelWidth - 0.15,
          height: sub.height,
          fontSize: Math.max(scaling.fontSize - 1, 6),
          bgColor: template.colors.subSwimLaneHeader || adjustColor(template.colors.swimLaneHeader, 20),
          textColor: template.colors.subSwimLaneHeaderText || template.colors.swimLaneHeaderText,
          font: template.fonts.primary,
          bold: (tStyles.subSwimLaneLabel || {}).bold,
          italic: (tStyles.subSwimLaneLabel || {}).italic,
          underline: (tStyles.subSwimLaneLabel || {}).underline,
        });
      }
    }

    if (i < laneLayouts.length - 1) {
      const sepY = lane.yOffset + lane.height;
      laneSeparators.push({
        type: "laneSeparator",
        x1: cfg.leftMargin,
        y: sepY + scaling.laneGap / 2,
        x2: totalRenderWidth - cfg.rightMargin,
        color: template.colors.gridLine,
      });
    }
  });

  return {
    slide: SLIDE,
    ganttArea: { left: ganttLeft, top: ganttTop, width: ganttWidth, height: availableHeight },
    timeAxis: timeAxisElements,
    laneLabels: laneLabelElements,
    subLaneLabels: subLaneLabelElements,
    laneSeparators,
    tasks: taskElements,
    dependencies: dependencyLines,
    todayMarker,
    elapsedShading,
    weekendShading,
    criticalPathIds,
    scaling,
    dateRange,
  };
}

// ── Phase 1: Date Range ──

function calculateDateRange(tasks, paddingPercent) {
  let minDate = null;
  let maxDate = null;

  for (const task of tasks) {
    const dates = [task.startDate, task.endDate, task.plannedStartDate, task.plannedEndDate];
    for (const d of dates) {
      if (!d) continue;
      if (!minDate || d < minDate) minDate = d;
      if (!maxDate || d > maxDate) maxDate = d;
    }
  }

  if (!minDate || !maxDate) {
    minDate = new Date();
    maxDate = new Date();
    maxDate.setMonth(maxDate.getMonth() + 1);
  }

  const rangeMs = maxDate.getTime() - minDate.getTime();
  const bufferMs = rangeMs * paddingPercent;
  const paddedMin = new Date(minDate.getTime() - bufferMs);
  const paddedMax = new Date(maxDate.getTime() + bufferMs);

  return { min: paddedMin, max: paddedMax, totalMs: paddedMax - paddedMin };
}

// ── Phase 2: Horizontal Mapping ──

function createDateMapper(dateRange, ganttLeft, ganttWidth) {
  return function mapDateToX(date) {
    if (!date) return ganttLeft;
    const ratio = (date.getTime() - dateRange.min.getTime()) / dateRange.totalMs;
    return ganttLeft + ratio * ganttWidth;
  };
}

// ── Phase 3: Swim Lane Allocation ──

function allocateSwimLanes(swimLanes, ganttTop, availableHeight, scaling) {
  const n = swimLanes.length;
  if (n === 0) return [];

  const totalGapSpace = (n - 1) * scaling.laneGap;
  const laneHeight = (availableHeight - totalGapSpace) / n;

  return swimLanes.map((lane, i) => {
    const yOffset = ganttTop + i * (laneHeight + scaling.laneGap);
    const result = {
      ...lane,
      yOffset,
      height: laneHeight,
    };

    // Allocate sub-swimlanes within this lane
    if (lane.subLanes && lane.subLanes.length > 0) {
      const topTaskCount = lane.topLevelTasks ? lane.topLevelTasks.length : 0;
      const totalItems = topTaskCount + lane.subLanes.length;
      if (totalItems > 0) {
        const subHeight = laneHeight / (totalItems || 1);
        let subY = yOffset;
        if (topTaskCount > 0) {
          subY += subHeight * topTaskCount;
        }
        result.subLaneLayouts = lane.subLanes.map((sub, si) => ({
          ...sub,
          yOffset: subY + si * subHeight,
          height: subHeight,
        }));
      }
    }

    return result;
  });
}

// ── Phase 4: Task Positioning ──

function positionTasks(laneLayouts, mapDateToX, dateRange, scaling, template, cfg) {
  const elements = [];

  for (const lane of laneLayouts) {
    // Position top-level tasks
    positionTasksInArea(
      lane.topLevelTasks || lane.tasks,
      lane.yOffset,
      lane.height,
      mapDateToX,
      scaling,
      template,
      cfg,
      elements
    );

    // Position sub-swimlane tasks
    if (lane.subLaneLayouts) {
      for (const sub of lane.subLaneLayouts) {
        positionTasksInArea(
          sub.tasks,
          sub.yOffset,
          sub.height,
          mapDateToX,
          scaling,
          template,
          cfg,
          elements
        );
      }
    }
  }

  return elements;
}

function positionTasksInArea(tasks, areaTop, areaHeight, mapDateToX, scaling, template, cfg, elements) {
  const headerSpace = scaling.barHeight * 0.2;
  const contentTop = areaTop + headerSpace;

  const sorted = [...tasks].sort((a, b) => (a.startDate || 0) - (b.startDate || 0));
  const rows = [];

  for (const task of sorted) {
    const x1 = mapDateToX(task.startDate);
    const x2 = task.type === TaskType.MILESTONE ? x1 : mapDateToX(task.endDate);

    let placed = false;
    for (let r = 0; r < rows.length; r++) {
      if (x1 >= rows[r] + 0.05) {
        rows[r] = x2;
        placeTask(elements, task, x1, x2, contentTop, r, scaling, template, cfg, mapDateToX);
        placed = true;
        break;
      }
    }

    if (!placed) {
      rows.push(x2);
      placeTask(elements, task, x1, x2, contentTop, rows.length - 1, scaling, template, cfg, mapDateToX);
    }
  }
}

function placeTask(elements, task, x1, x2, contentTop, row, scaling, template, cfg, mapDateToX) {
  const rowHeight = scaling.barHeight + 0.06;
  const top = contentTop + row * rowHeight;

  // Conditional formatting: color by variance if baselines exist
  let barColor = task.status
    ? getStatusColor(template, task.status)
    : template.colors.taskBar;

  const variance = calculateVariance(task);
  if (variance !== null && template.colors.varianceEarly && template.colors.varianceLate) {
    if (variance < -1) barColor = template.colors.varianceEarly;
    else if (variance > 1) barColor = template.colors.varianceLate;
  }

  const duration = getTaskDuration(task);

  // Label configuration (from template, with config override)
  const labelCfg = template.labelConfig || {};
  const taskLabelPos = cfg.taskLabelPositionOverride || labelCfg.taskLabelPosition || "inside";
  const milestoneLabelPos = labelCfg.milestoneLabelPosition || "above";
  const taskLabelAlign = labelCfg.taskLabelAlign || "left";
  const textStyles = (template.fonts && template.fonts.styles) || {};

  if (task.type === TaskType.MILESTONE) {
    const msDate = task.startDate
      ? task.startDate.toLocaleDateString("en-US", { month: "short", day: "numeric" })
      : "";
    const msSize = scaling.barHeight * 0.65;

    // Compose milestone label from config
    let msLabel = task.name;
    if (labelCfg.showOwner && task.owner) msLabel += ` (${task.owner})`;

    // Label position: above or below
    const labelTop = milestoneLabelPos === "below"
      ? top + scaling.barHeight + 0.05
      : top - scaling.barHeight * 0.5;

    elements.push({
      type: "milestone",
      id: task.id,
      taskId: task.id,
      name: msLabel,
      dateLabel: msDate,
      milestoneShape: task.milestoneShape || "diamond",
      left: x1 - msSize / 2,
      top: top + (scaling.barHeight - msSize) / 2,
      size: msSize,
      color: task.status ? getStatusColor(template, task.status) : template.colors.milestone,
      textColor: template.colors.labelText,
      fontSize: template.fonts.sizes.milestoneLabel,
      font: template.fonts.primary,
      bold: (textStyles.milestoneLabel || {}).bold,
      italic: (textStyles.milestoneLabel || {}).italic,
      underline: (textStyles.milestoneLabel || {}).underline,
      centerX: x1,
      centerY: top + scaling.barHeight / 2,
      labelPosition: milestoneLabelPos,
      labelTop: labelTop,
      labelLeft: x1 - 0.5,
      labelWidth: 1.0,
      dateBold: (textStyles.milestoneDateLabel || {}).bold,
      dateItalic: (textStyles.milestoneDateLabel || {}).italic,
      dateUnderline: (textStyles.milestoneDateLabel || {}).underline,
    });
  } else {
    const width = Math.max(x2 - x1, 0.1);

    // Compose label text based on config
    let labelText = task.name;
    if (labelCfg.showOwner && task.owner) labelText += ` (${task.owner})`;
    if (labelCfg.showTaskDates && task.startDate && task.endDate) {
      const s = task.startDate.toLocaleDateString("en-US", { month: "short", day: "numeric" });
      const e = task.endDate.toLocaleDateString("en-US", { month: "short", day: "numeric" });
      labelText += ` [${s} – ${e}]`;
    }

    // Main task bar
    const taskEl = {
      type: "taskBar",
      id: task.id,
      taskId: task.id,
      name: labelText,
      left: x1,
      top: top,
      width: width,
      height: scaling.barHeight,
      color: barColor,
      textColor: template.colors.taskText,
      fontSize: template.fonts.sizes.taskLabel,
      font: template.fonts.primary,
      // Text formatting from template.fonts.styles.taskLabel
      bold: (textStyles.taskLabel || {}).bold,
      italic: (textStyles.taskLabel || {}).italic,
      underline: (textStyles.taskLabel || {}).underline,
      // Label positioning config
      labelPosition: taskLabelPos,      // "inside" | "above" | "below" | "left" | "right"
      labelAlign: taskLabelAlign,       // "left" | "center" | "right"
      labelWrap: labelCfg.labelWrap || false,
      cornerRadius: template.shapes.taskBarCornerRadius,
      centerX: x1 + width / 2,
      centerY: top + scaling.barHeight / 2,
      rightEdge: x1 + width,
      leftEdge: x1,
      // Percent complete
      percentComplete: cfg.showPercentComplete ? task.percentComplete : null,
      percentWidth: task.percentComplete !== null && cfg.showPercentComplete
        ? width * (task.percentComplete / 100)
        : null,
      percentColor: template.colors.percentComplete || darkenColor(barColor, 30),
      // Duration label
      durationLabel: cfg.showDurationLabels && duration > 0 ? `${duration}d` : null,
      durationLabelColor: template.colors.durationLabel || template.colors.labelText,
      durationBold: (textStyles.durationLabel || {}).bold,
      durationItalic: (textStyles.durationLabel || {}).italic,
      durationUnderline: (textStyles.durationLabel || {}).underline,
      // Variance
      variance: variance,
    };

    // 3D/Gel style metadata
    if (cfg.styleMode === "3d") {
      taskEl.style3d = true;
      taskEl.highlightColor = adjustColor(barColor, 40);
      taskEl.shadowColor = darkenColor(barColor, 40);
    }

    elements.push(taskEl);

    // Baseline bar (planned vs actual) — rendered as a thinner bar below the actual
    if (cfg.showBaselines && task.plannedStartDate && task.plannedEndDate) {
      const bx1 = mapDateToX(task.plannedStartDate);
      const bx2 = mapDateToX(task.plannedEndDate);
      const bWidth = Math.max(bx2 - bx1, 0.05);
      const bHeight = scaling.barHeight * 0.35;
      const bTop = top + scaling.barHeight + 0.01;

      elements.push({
        type: "baselineBar",
        id: `baseline_${task.id}`,
        taskId: task.id,
        name: `Planned: ${task.name}`,
        left: bx1,
        top: bTop,
        width: bWidth,
        height: bHeight,
        color: template.colors.baselineBar || "#C0C0C0",
        opacity: 0.6,
        cornerRadius: template.shapes.taskBarCornerRadius,
      });
    }
  }
}

function getStatusColor(template, status) {
  const map = {
    ON_TRACK: template.colors.statusOnTrack,
    AT_RISK: template.colors.statusAtRisk,
    DELAYED: template.colors.statusDelayed,
    COMPLETE: template.colors.statusComplete,
  };
  return map[status] || template.colors.taskBar;
}

// ── Phase 5: Dependency Routing ──

function buildTaskPositionMap(taskElements) {
  const map = new Map();
  for (const el of taskElements) {
    if (el.type === "taskBar" || el.type === "milestone") {
      map.set(el.taskId, el);
    }
  }
  return map;
}

function routeDependencies(allTasks, positionMap, template) {
  const lines = [];

  for (const task of allTasks) {
    if (!task.dependencies || task.dependencies.length === 0) continue;

    const toEl = positionMap.get(task.id);
    if (!toEl) continue;

    for (const depId of task.dependencies) {
      const fromEl = positionMap.get(depId);
      if (!fromEl) continue;

      const depInfo = task.dependencyTypes.get(depId) || { type: DepType.FS, lagDays: 0 };
      let fromX, fromY, toX, toY;

      // Route based on dependency type
      switch (depInfo.type) {
        case DepType.FF:
          fromX = fromEl.type === "milestone" ? fromEl.centerX : fromEl.rightEdge;
          fromY = fromEl.centerY;
          toX = toEl.type === "milestone" ? toEl.centerX : toEl.rightEdge;
          toY = toEl.centerY;
          break;
        case DepType.SS:
          fromX = fromEl.type === "milestone" ? fromEl.centerX : fromEl.leftEdge;
          fromY = fromEl.centerY;
          toX = toEl.type === "milestone" ? toEl.centerX : toEl.leftEdge;
          toY = toEl.centerY;
          break;
        case DepType.SF:
          fromX = fromEl.type === "milestone" ? fromEl.centerX : fromEl.leftEdge;
          fromY = fromEl.centerY;
          toX = toEl.type === "milestone" ? toEl.centerX : toEl.rightEdge;
          toY = toEl.centerY;
          break;
        default: // FS
          fromX = fromEl.type === "milestone" ? fromEl.centerX : fromEl.rightEdge;
          fromY = fromEl.centerY;
          toX = toEl.type === "milestone" ? toEl.centerX : toEl.leftEdge;
          toY = toEl.centerY;
          break;
      }

      const midX = fromX + (toX - fromX) * 0.5;

      lines.push({
        type: "dependencyLine",
        fromTaskId: depId,
        toTaskId: task.id,
        depType: depInfo.type,
        lagDays: depInfo.lagDays,
        points: [
          { x: fromX, y: fromY },
          { x: midX, y: fromY },
          { x: midX, y: toY },
          { x: toX, y: toY },
        ],
        color: template.colors.dependency,
        weight: template.shapes.dependencyLineWeight,
        isCriticalPath: false,
        arrowSize: template.shapes.arrowSize,
      });
    }
  }

  return lines;
}

// ── Critical Path ──

function findCriticalPath(tasks) {
  const taskMap = new Map();
  for (const t of tasks) {
    taskMap.set(t.id, t);
  }

  function getDuration(task) {
    if (task.type === TaskType.MILESTONE) return 0;
    if (!task.startDate || !task.endDate) return 0;
    return (task.endDate - task.startDate) / (1000 * 60 * 60 * 24);
  }

  const memo = new Map();

  function longestPath(taskId) {
    if (memo.has(taskId)) return memo.get(taskId);
    const task = taskMap.get(taskId);
    if (!task) return { length: 0, path: [] };

    let maxPred = { length: 0, path: [] };
    for (const depId of task.dependencies || []) {
      const sub = longestPath(depId);
      if (sub.length > maxPred.length) maxPred = sub;
    }

    const result = {
      length: maxPred.length + getDuration(task),
      path: [...maxPred.path, taskId],
    };
    memo.set(taskId, result);
    return result;
  }

  let criticalPath = { length: 0, path: [] };
  for (const task of tasks) {
    const result = longestPath(task.id);
    if (result.length > criticalPath.length) criticalPath = result;
  }

  return new Set(criticalPath.path);
}

// ── 3-Tier Time Axis ──

function buildTimeAxis(dateRange, mapDateToX, cfg, template, tiersUsed) {
  const elements = [];
  const axStyles = (template.fonts && template.fonts.styles) || {};

  const ganttLeft = cfg.leftMargin + cfg.swimLaneLabelWidth;
  const ganttWidth =
    SLIDE.width * cfg.renderWidthPercent - cfg.leftMargin - cfg.rightMargin - cfg.swimLaneLabelWidth;

  let currentTop = cfg.topMargin;
  const fyStart = cfg.fiscalYearStartMonth;

  // ── Tier 1: Years (or fiscal years / quarters) ──
  const tier1Top = currentTop;
  elements.push({
    type: "timeAxisBg",
    left: ganttLeft,
    top: tier1Top,
    width: ganttWidth,
    height: cfg.tier1Height,
    color: template.colors.yearAxisBg || template.colors.timeAxisBg,
  });

  const years = getYearSpans(dateRange, mapDateToX, ganttLeft, ganttWidth, fyStart, cfg.fiscalYearLabelFormat || "end", cfg.fiscalYearPrefix || "FY");
  for (const yr of years) {
    elements.push({
      type: "yearLabel",
      text: yr.label,
      left: yr.left,
      top: tier1Top,
      width: yr.width,
      height: cfg.tier1Height,
      fontSize: (template.fonts.sizes.timeAxis || 8) + 2,
      textColor: template.colors.yearAxisText || template.colors.timeAxisText,
      font: template.fonts.primary,
      bold: (axStyles.yearLabel || {}).bold !== undefined ? axStyles.yearLabel.bold : true,
      italic: (axStyles.yearLabel || {}).italic,
      underline: (axStyles.yearLabel || {}).underline,
    });
    elements.push({
      type: "yearBoundary",
      x: yr.left,
      top: tier1Top,
      bottom: SLIDE.height - cfg.bottomMargin,
      color: template.colors.yearBoundary || "#B0B0B0",
    });
  }
  currentTop += cfg.tier1Height;

  // ── Tier 2: Months ──
  if (tiersUsed >= 2) {
    const tier2Top = currentTop;
    elements.push({
      type: "timeAxisBg",
      left: ganttLeft,
      top: tier2Top,
      width: ganttWidth,
      height: cfg.tier2Height,
      color: template.colors.monthAxisBg || template.colors.timeAxisBg,
    });

    const months = getMonthTicks(dateRange, mapDateToX, ganttLeft, ganttWidth);
    for (const m of months) {
      elements.push({
        type: "monthLabel",
        text: m.label,
        left: m.left,
        top: tier2Top,
        width: m.width,
        height: cfg.tier2Height,
        fontSize: template.fonts.sizes.timeAxis || 8,
        textColor: template.colors.monthAxisText || template.colors.timeAxisText,
        font: template.fonts.primary,
        bold: (axStyles.monthLabel || {}).bold,
        italic: (axStyles.monthLabel || {}).italic,
        underline: (axStyles.monthLabel || {}).underline,
      });
      elements.push({
        type: "gridLine",
        x: m.left,
        top: tier2Top + cfg.tier2Height + (tiersUsed >= 3 ? cfg.tier3Height : 0),
        bottom: SLIDE.height - cfg.bottomMargin,
        color: template.colors.gridLine,
      });
    }
    currentTop += cfg.tier2Height;
  }

  // ── Tier 3: Weeks or Days ──
  if (tiersUsed >= 3) {
    const tier3Top = currentTop;
    elements.push({
      type: "timeAxisBg",
      left: ganttLeft,
      top: tier3Top,
      width: ganttWidth,
      height: cfg.tier3Height,
      color: template.colors.tier3AxisBg || adjustColor(template.colors.monthAxisBg || template.colors.timeAxisBg, 15),
    });

    const rangeDays = (dateRange.max - dateRange.min) / (1000 * 60 * 60 * 24);
    const granularity = cfg.timescaleGranularity || "auto";
    const weekTicks = getWeekOrDayTicks(dateRange, mapDateToX, ganttLeft, ganttWidth, rangeDays, granularity);
    for (const t of weekTicks) {
      elements.push({
        type: "tier3Label",
        text: t.label,
        left: t.left,
        top: tier3Top,
        width: t.width,
        height: cfg.tier3Height,
        fontSize: Math.max((template.fonts.sizes.timeAxis || 8) - 1, 6),
        textColor: template.colors.tier3AxisText || template.colors.monthAxisText || template.colors.timeAxisText,
        font: template.fonts.primary,
        bold: (axStyles.tier3Label || {}).bold,
        italic: (axStyles.tier3Label || {}).italic,
        underline: (axStyles.tier3Label || {}).underline,
      });
    }
    currentTop += cfg.tier3Height;
  }

  return elements;
}

function getYearSpans(dateRange, mapDateToX, ganttLeft, ganttWidth, fyStart, labelFormat = "end", prefix = "FY") {
  const spans = [];
  const startYear = dateRange.min.getFullYear() - 1;
  const endYear = dateRange.max.getFullYear() + 1;

  if (fyStart === 1) {
    // Calendar years
    for (let y = startYear; y <= endYear; y++) {
      const yearStart = new Date(y, 0, 1);
      const yearEnd = new Date(y + 1, 0, 1);
      const span = clampSpan(yearStart, yearEnd, dateRange, mapDateToX, ganttLeft, ganttWidth);
      if (span) spans.push({ ...span, label: String(y) });
    }
  } else {
    // Fiscal years — label format options
    for (let y = startYear; y <= endYear; y++) {
      const fyStartDate = new Date(y, fyStart - 1, 1);
      const fyEndDate = new Date(y + 1, fyStart - 1, 1);
      const span = clampSpan(fyStartDate, fyEndDate, dateRange, mapDateToX, ganttLeft, ganttWidth);
      if (!span) continue;

      let label;
      switch (labelFormat) {
        case "start":
          // FY starting in year y → labeled with y (e.g. "FY2025" for Jul 2025 – Jun 2026)
          label = `${prefix}${y}`;
          break;
        case "both":
          // FY2025/26
          label = `${prefix}${y}/${String(y + 1).slice(-2)}`;
          break;
        case "end":
        default:
          // FY labeled with end year (e.g. "FY2026" for Jul 2025 – Jun 2026)
          label = `${prefix}${y + 1}`;
          break;
      }
      spans.push({ ...span, label });
    }
  }

  return spans;
}

function getMonthTicks(dateRange, mapDateToX, ganttLeft, ganttWidth) {
  const ticks = [];
  const MONTH_NAMES = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const current = new Date(dateRange.min.getFullYear(), dateRange.min.getMonth(), 1);

  while (current <= dateRange.max) {
    const monthEnd = new Date(current.getFullYear(), current.getMonth() + 1, 1);
    const span = clampSpan(current, monthEnd, dateRange, mapDateToX, ganttLeft, ganttWidth);
    if (span) {
      ticks.push({ ...span, label: MONTH_NAMES[current.getMonth()], date: new Date(current) });
    }
    current.setMonth(current.getMonth() + 1);
  }

  return ticks;
}

function getWeekOrDayTicks(dateRange, mapDateToX, ganttLeft, ganttWidth, rangeDays, granularity) {
  const ticks = [];

  // Hours granularity for very short ranges (< 3 days) or explicit request
  if (rangeDays <= 3 || granularity === "hours") {
    const current = new Date(dateRange.min);
    current.setMinutes(0, 0, 0);
    while (current <= dateRange.max) {
      const hourEnd = new Date(current);
      hourEnd.setHours(hourEnd.getHours() + 1);
      const span = clampSpan(current, hourEnd, dateRange, mapDateToX, ganttLeft, ganttWidth);
      if (span && span.width > 0.04) {
        const h = current.getHours();
        const label = h === 0 ? `${current.getMonth() + 1}/${current.getDate()}`
          : `${h}:00`;
        ticks.push({ ...span, label });
      }
      current.setHours(current.getHours() + 1);
    }
  } else if (rangeDays <= 60 || granularity === "days") {
    // Show individual days
    const current = new Date(dateRange.min);
    current.setHours(0, 0, 0, 0);
    while (current <= dateRange.max) {
      const dayEnd = new Date(current);
      dayEnd.setDate(dayEnd.getDate() + 1);
      const span = clampSpan(current, dayEnd, dateRange, mapDateToX, ganttLeft, ganttWidth);
      if (span && span.width > 0.05) {
        ticks.push({ ...span, label: String(current.getDate()) });
      }
      current.setDate(current.getDate() + 1);
    }
  } else {
    // Show weeks (W1, W2, ...)
    const current = new Date(dateRange.min);
    current.setHours(0, 0, 0, 0);
    // Align to Monday
    const day = current.getDay();
    current.setDate(current.getDate() - (day === 0 ? 6 : day - 1));

    let weekNum = 1;
    while (current <= dateRange.max) {
      const weekEnd = new Date(current);
      weekEnd.setDate(weekEnd.getDate() + 7);
      const span = clampSpan(current, weekEnd, dateRange, mapDateToX, ganttLeft, ganttWidth);
      if (span && span.width > 0.08) {
        ticks.push({ ...span, label: `W${weekNum}` });
      }
      current.setDate(current.getDate() + 7);
      weekNum++;
    }
  }

  return ticks;
}

function clampSpan(start, end, dateRange, mapDateToX, ganttLeft, ganttWidth) {
  const clampedStart = start < dateRange.min ? dateRange.min : start;
  const clampedEnd = end > dateRange.max ? dateRange.max : end;
  let left = mapDateToX(clampedStart);
  let right = mapDateToX(clampedEnd);
  left = Math.max(left, ganttLeft);
  right = Math.min(right, ganttLeft + ganttWidth);
  if (right <= left) return null;
  return { left, width: right - left };
}

// ── Sorting ──

function sortTasks(tasks, sortBy) {
  switch (sortBy) {
    case "startDate":
      tasks.sort((a, b) => (a.startDate || 0) - (b.startDate || 0));
      break;
    case "endDate":
      tasks.sort((a, b) => (a.endDate || 0) - (b.endDate || 0));
      break;
    case "name":
      tasks.sort((a, b) => a.name.localeCompare(b.name));
      break;
    case "status":
      const order = { ON_TRACK: 0, AT_RISK: 1, DELAYED: 2, COMPLETE: 3 };
      tasks.sort((a, b) => (order[a.status] || 99) - (order[b.status] || 99));
      break;
  }
}

// ── Label Formatting ──

function formatLaneLabel(name, mode) {
  if (mode === "full") return name;
  if (mode === "abbreviated") {
    if (name.length <= 15) return name;
    return name.substring(0, 14) + "\u2026";
  }
  if (name.length <= 10) return name;
  return name.substring(0, 9) + "\u2026";
}

// ── Color Utilities ──

function hexToRgb(hex) {
  const h = hex.replace("#", "");
  return {
    r: parseInt(h.substring(0, 2), 16),
    g: parseInt(h.substring(2, 4), 16),
    b: parseInt(h.substring(4, 6), 16),
  };
}

function rgbToHex(r, g, b) {
  return "#" + [r, g, b].map((v) => Math.max(0, Math.min(255, v)).toString(16).padStart(2, "0")).join("");
}

function adjustColor(hex, amount) {
  if (!hex) return "#808080";
  const { r, g, b } = hexToRgb(hex);
  return rgbToHex(r + amount, g + amount, b + amount);
}

function darkenColor(hex, amount) {
  return adjustColor(hex, -amount);
}

module.exports = {
  calculateLayout,
  SLIDE,
  DEFAULTS,
  SCALING_RULES,
};
