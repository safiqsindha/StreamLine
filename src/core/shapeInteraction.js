/**
 * Streamline Shape Interaction
 *
 * Three systems that replicate COM-level interactivity via Office JS:
 *
 * 1. SELECTION CHANGE — onSelectionChanged fires when user clicks any shape.
 *    We parse the tag, find the task, and push it to the task pane for instant editing.
 *
 * 2. POSITION POLLING — a 500ms timer snapshots all Streamline shape positions.
 *    When a shape moves (user drags it natively in PowerPoint), we detect the delta,
 *    reverse-map the new X position to a date, and update the task model.
 *
 * 3. CONTEXTUAL ACTIONS — when a shape is selected, the task pane shows a context
 *    panel with actions: edit name, change color, change dates, delete, duplicate.
 *    This replaces what a right-click context menu would do in a COM add-in.
 */

const { SHAPE_TAG_PREFIX } = require("./powerpointRenderer");

// ═══════════════════════════════════════════════════════════
// Shape Position Tracker (drag detection via polling)
// ═══════════════════════════════════════════════════════════

class ShapePositionTracker {
  constructor() {
    this._snapshot = new Map();   // tag -> { left, top, width, height }
    this._interval = null;
    this._onMoveCallback = null;
    this._onResizeCallback = null;
    this._pollMs = 500;
    this._tolerance = 0.01;       // inches — ignore sub-pixel jitter
  }

  /**
   * Start polling shape positions.
   * @param {Function} onMove  - (tag, shapeType, taskId, oldPos, newPos) => void
   * @param {Function} onResize - (tag, shapeType, taskId, oldSize, newSize) => void
   */
  start(onMove, onResize) {
    this._onMoveCallback = onMove;
    this._onResizeCallback = onResize;

    // Take initial snapshot immediately
    this._poll();

    this._interval = setInterval(() => this._poll(), this._pollMs);
  }

  stop() {
    if (this._interval) {
      clearInterval(this._interval);
      this._interval = null;
    }
    this._snapshot.clear();
  }

  isRunning() {
    return this._interval !== null;
  }

  /**
   * Force a full re-snapshot (call after a fresh render to reset baseline).
   */
  async resetSnapshot() {
    this._snapshot.clear();
    await this._poll();
  }

  async _poll() {
    let currentShapes;
    try {
      currentShapes = await readAllShapePositions();
    } catch (e) {
      // PowerPoint context not available (tab switched, etc.) — skip this tick
      return;
    }

    for (const shape of currentShapes) {
      const prev = this._snapshot.get(shape.tag);

      if (!prev) {
        // First time seeing this shape — just record it
        this._snapshot.set(shape.tag, { left: shape.left, top: shape.top, width: shape.width, height: shape.height });
        continue;
      }

      const dx = Math.abs(shape.left - prev.left);
      const dy = Math.abs(shape.top - prev.top);
      const dw = Math.abs(shape.width - prev.width);
      const dh = Math.abs(shape.height - prev.height);

      // Detect move
      if ((dx > this._tolerance || dy > this._tolerance) && this._onMoveCallback) {
        const parts = parseTag(shape.tag);
        if (parts) {
          this._onMoveCallback(
            shape.tag,
            parts.type,
            parts.id,
            { left: prev.left, top: prev.top },
            { left: shape.left, top: shape.top }
          );
        }
      }

      // Detect resize (width change = duration change for task bars)
      if ((dw > this._tolerance || dh > this._tolerance) && this._onResizeCallback) {
        const parts = parseTag(shape.tag);
        if (parts) {
          this._onResizeCallback(
            shape.tag,
            parts.type,
            parts.id,
            { width: prev.width, height: prev.height },
            { width: shape.width, height: shape.height }
          );
        }
      }

      // Update snapshot
      this._snapshot.set(shape.tag, { left: shape.left, top: shape.top, width: shape.width, height: shape.height });
    }
  }
}

// ═══════════════════════════════════════════════════════════
// Selection Change Handler
// ═══════════════════════════════════════════════════════════

class SelectionWatcher {
  constructor() {
    this._callback = null;
    this._registered = false;
    this._lastSelectedTag = null;
  }

  /**
   * Register for selection change events.
   * @param {Function} callback - (shapeInfo | null) => void
   *   shapeInfo = { tag, type, id, left, top, width, height } or null if deselected
   */
  start(callback) {
    this._callback = callback;

    if (this._registered) return;

    try {
      // Office JS PowerPoint selection change event
      Office.context.document.addHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        () => this._onSelectionChanged(),
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            this._registered = true;
          }
        }
      );
    } catch (e) {
      // Fallback: poll selection every 300ms if event not available
      this._pollInterval = setInterval(() => this._onSelectionChanged(), 300);
      this._registered = true;
    }
  }

  stop() {
    if (this._pollInterval) {
      clearInterval(this._pollInterval);
      this._pollInterval = null;
    }
    this._registered = false;
    this._callback = null;
  }

  async _onSelectionChanged() {
    if (!this._callback) return;

    try {
      const shape = await getSelectedStreamlineShape();

      // Only fire callback if selection actually changed
      const newTag = shape ? shape.tag : null;
      if (newTag === this._lastSelectedTag) return;
      this._lastSelectedTag = newTag;

      this._callback(shape);
    } catch (e) {
      // Ignore — PowerPoint context may not be ready
    }
  }
}

// ═══════════════════════════════════════════════════════════
// Coordinate ↔ Date Mapping (reverse the layout engine)
// ═══════════════════════════════════════════════════════════

/**
 * Create a reverse mapper: given an X position (inches), return the corresponding date.
 * This is the inverse of createDateMapper in layoutEngine.
 *
 * @param {Object} ganttArea - { left, width } from layout
 * @param {Object} dateRange - { min, max, totalMs } from layout
 * @returns {Function} mapXToDate(x) => Date
 */
function createReverseMapper(ganttArea, dateRange) {
  return function mapXToDate(x) {
    const ratio = (x - ganttArea.left) / ganttArea.width;
    const clampedRatio = Math.max(0, Math.min(1, ratio));
    const ms = dateRange.min.getTime() + clampedRatio * dateRange.totalMs;
    const d = new Date(ms);
    // Round to nearest day
    d.setHours(0, 0, 0, 0);
    return d;
  };
}

/**
 * Given a shape's new position and the layout context, compute new task dates.
 *
 * @param {string} shapeType - "taskbar" or "milestone"
 * @param {Object} newPos - { left, width } in inches
 * @param {Function} mapXToDate - reverse mapper
 * @returns {Object} { startDate, endDate } (endDate is null for milestones)
 */
function computeDatesFromPosition(shapeType, newPos, mapXToDate) {
  const startDate = mapXToDate(newPos.left);

  if (shapeType === "milestone") {
    return { startDate, endDate: null };
  }

  const endDate = mapXToDate(newPos.left + newPos.width);
  return { startDate, endDate };
}

// ═══════════════════════════════════════════════════════════
// Shape Read/Write Helpers (Office JS)
// ═══════════════════════════════════════════════════════════

/**
 * Read all Streamline-tagged shape positions from the active slide.
 * Returns only task-level shapes (taskbar, milestone).
 */
async function readAllShapePositions() {
  const results = [];

  await PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const shapes = slide.shapes;
    shapes.load("items/name,items/left,items/top,items/width,items/height");
    await context.sync();

    for (const shape of shapes.items) {
      if (!shape.name || !shape.name.startsWith(SHAPE_TAG_PREFIX + ":")) continue;

      const parts = parseTag(shape.name);
      if (!parts) continue;

      // Only track moveable shapes
      if (["taskbar", "milestone"].includes(parts.type)) {
        results.push({
          tag: shape.name,
          type: parts.type,
          id: parts.id,
          left: shape.left,
          top: shape.top,
          width: shape.width,
          height: shape.height,
        });
      }
    }
  });

  return results;
}

/**
 * Get info about the currently selected shape on the active slide.
 */
async function getSelectedStreamlineShape() {
  let result = null;

  await PowerPoint.run(async (context) => {
    const selection = context.presentation.getSelectedShapes();
    selection.load("items/name,items/left,items/top,items/width,items/height");
    await context.sync();

    if (selection.items.length === 0) return;

    const shape = selection.items[0];
    if (!shape.name || !shape.name.startsWith(SHAPE_TAG_PREFIX + ":")) return;

    const parts = parseTag(shape.name);
    result = {
      tag: shape.name,
      type: parts ? parts.type : "",
      id: parts ? parts.id : "",
      left: shape.left,
      top: shape.top,
      width: shape.width,
      height: shape.height,
    };
  });

  return result;
}

/**
 * Move a shape to new coordinates.
 */
async function moveShape(shapeTag, newLeft, newTop) {
  await PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const shapes = slide.shapes;
    shapes.load("items/name");
    await context.sync();

    for (const shape of shapes.items) {
      if (shape.name === shapeTag) {
        shape.left = newLeft;
        shape.top = newTop;
        break;
      }
    }
    await context.sync();
  });
}

/**
 * Resize a shape.
 */
async function resizeShape(shapeTag, newWidth, newHeight) {
  await PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const shapes = slide.shapes;
    shapes.load("items/name");
    await context.sync();

    for (const shape of shapes.items) {
      if (shape.name === shapeTag) {
        if (newWidth !== undefined) shape.width = newWidth;
        if (newHeight !== undefined) shape.height = newHeight;
        break;
      }
    }
    await context.sync();
  });
}

/**
 * Update shape text content.
 */
async function updateShapeText(shapeTag, newText) {
  await PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const shapes = slide.shapes;
    shapes.load("items/name");
    await context.sync();

    for (const shape of shapes.items) {
      if (shape.name === shapeTag) {
        shape.textFrame.textRange.text = newText;
        break;
      }
    }
    await context.sync();
  });
}

/**
 * Update shape fill color.
 */
async function updateShapeColor(shapeTag, newColor) {
  await PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const shapes = slide.shapes;
    shapes.load("items/name");
    await context.sync();

    for (const shape of shapes.items) {
      if (shape.name === shapeTag) {
        shape.fill.setSolidColor(newColor);
        break;
      }
    }
    await context.sync();
  });
}

/**
 * Delete a Streamline shape and its associated shapes (label, baseline, etc.).
 */
async function deleteShapeGroup(taskId) {
  await PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const shapes = slide.shapes;
    shapes.load("items/name");
    await context.sync();

    const toDelete = [];
    for (const shape of shapes.items) {
      if (shape.name && shape.name.includes(`:${taskId}`)) {
        toDelete.push(shape);
      }
    }

    for (const shape of toDelete) {
      shape.delete();
    }
    await context.sync();
  });
}

// ═══════════════════════════════════════════════════════════
// Tag Parsing
// ═══════════════════════════════════════════════════════════

function parseTag(tag) {
  if (!tag || !tag.startsWith(SHAPE_TAG_PREFIX + ":")) return null;
  const parts = tag.split(":");
  if (parts.length < 3) return null;
  return { type: parts[1], id: parts[2] };
}

module.exports = {
  // Classes
  ShapePositionTracker,
  SelectionWatcher,
  // Reverse mapping
  createReverseMapper,
  computeDatesFromPosition,
  // Shape operations
  getSelectedStreamlineShape,
  readAllShapePositions,
  moveShape,
  resizeShape,
  updateShapeText,
  updateShapeColor,
  deleteShapeGroup,
  // Utilities
  parseTag,
};
