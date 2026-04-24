/**
 * Streamline PowerPoint Renderer
 * Translates layout coordinates into Office JS API calls to create native PowerPoint shapes.
 * All shape dimensions are in inches (Office JS default for PowerPoint).
 */

const SHAPE_TAG_PREFIX = "streamline";

function makeTag(type, id) {
  return `${SHAPE_TAG_PREFIX}:${type}:${id}`;
}

// Map milestone shape names to Office JS geometric shape types
const MILESTONE_SHAPE_MAP = {
  diamond: "Diamond",
  circle: "Ellipse",
  triangle: "IsoscelesTriangle",
  star: "Star5",
  flag: "RightArrow",
  square: "Rectangle",
};

/**
 * Render the full Gantt chart layout onto the active PowerPoint slide.
 */
async function renderGantt(layout) {
  await PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const shapes = slide.shapes;

    // ── Weekend Highlighting (behind everything) ──
    if (layout.weekendShading && layout.weekendShading.length > 0) {
      let wkIdx = 0;
      for (const wk of layout.weekendShading) {
        await addRectangle(shapes, {
          left: wk.left,
          top: wk.top,
          width: wk.width,
          height: wk.height,
          fill: wk.color,
          tag: makeTag("weekend", String(wkIdx++)),
          lineColor: null,
          opacity: 0.5,
        });
      }
    }

    // ── Elapsed Time Shading (behind everything) ──
    if (layout.elapsedShading) {
      const el = layout.elapsedShading;
      await addRectangle(shapes, {
        left: el.left,
        top: el.top,
        width: el.width,
        height: el.height,
        fill: el.color,
        tag: makeTag("elapsed", "bg"),
        lineColor: null,
        opacity: 0.5,
      });
    }

    // ── Time Axis (3-Tier) ──
    let bgIndex = 0;
    for (const el of layout.timeAxis) {
      if (el.type === "timeAxisBg") {
        await addRectangle(shapes, {
          left: el.left,
          top: el.top,
          width: el.width,
          height: el.height,
          fill: el.color,
          tag: makeTag("timeAxisBg", String(bgIndex++)),
          lineColor: null,
        });
      } else if (el.type === "yearLabel") {
        await addTextBox(shapes, {
          left: el.left,
          top: el.top,
          width: el.width,
          height: el.height,
          text: el.text,
          fontSize: el.fontSize,
          fontColor: el.textColor,
          fontFamily: el.font,
          alignment: "center",
          verticalAlignment: "middle",
          bold: el.bold,
          italic: el.italic,
          underline: el.underline,
          tag: makeTag("yearLabel", el.text),
          fill: null,
        });
      } else if (el.type === "monthLabel") {
        await addTextBox(shapes, {
          left: el.left,
          top: el.top,
          width: el.width,
          height: el.height,
          text: el.text,
          fontSize: el.fontSize,
          fontColor: el.textColor,
          fontFamily: el.font,
          alignment: "center",
          verticalAlignment: "middle",
          bold: el.bold,
          italic: el.italic,
          underline: el.underline,
          tag: makeTag("monthLabel", `${el.text}_${el.left.toFixed(1)}`),
          fill: null,
        });
      } else if (el.type === "tier3Label") {
        await addTextBox(shapes, {
          left: el.left,
          top: el.top,
          width: el.width,
          height: el.height,
          text: el.text,
          fontSize: el.fontSize,
          fontColor: el.textColor,
          fontFamily: el.font,
          alignment: "center",
          verticalAlignment: "middle",
          bold: el.bold,
          italic: el.italic,
          underline: el.underline,
          tag: makeTag("tier3", `${el.text}_${el.left.toFixed(1)}`),
          fill: null,
        });
      } else if (el.type === "gridLine") {
        await addLine(shapes, {
          x1: el.x, y1: el.top,
          x2: el.x, y2: el.bottom,
          color: el.color,
          weight: 0.5,
          dashStyle: "dash",
          tag: makeTag("gridLine", el.x.toFixed(2)),
        });
      } else if (el.type === "yearBoundary") {
        await addLine(shapes, {
          x1: el.x, y1: el.top,
          x2: el.x, y2: el.bottom,
          color: el.color,
          weight: 1,
          dashStyle: "solid",
          tag: makeTag("yearBound", el.x.toFixed(2)),
        });
      }
    }

    // ── Lane Separators ──
    if (layout.laneSeparators) {
      for (const sep of layout.laneSeparators) {
        await addLine(shapes, {
          x1: sep.x1, y1: sep.y,
          x2: sep.x2, y2: sep.y,
          color: sep.color,
          weight: 0.5,
          dashStyle: "solid",
          tag: makeTag("laneSep", sep.y.toFixed(2)),
        });
      }
    }

    // ── Swim Lane Labels ──
    for (const el of layout.laneLabels) {
      await addRectangle(shapes, {
        left: el.left,
        top: el.top,
        width: el.width,
        height: el.height,
        fill: el.bgColor,
        tag: makeTag("laneLabel", el.id),
        lineColor: null,
      });
      await addTextBox(shapes, {
        left: el.left + 0.08,
        top: el.top,
        width: el.width - 0.16,
        height: el.height,
        text: el.text,
        fontSize: el.fontSize,
        fontColor: el.textColor,
        fontFamily: el.font,
        alignment: "left",
        verticalAlignment: "middle",
        bold: el.bold !== undefined ? el.bold : true,
        italic: el.italic,
        underline: el.underline,
        tag: makeTag("laneLabelText", el.id),
        fill: null,
      });
    }

    // ── Sub-Swim Lane Labels ──
    if (layout.subLaneLabels) {
      for (const el of layout.subLaneLabels) {
        await addRectangle(shapes, {
          left: el.left,
          top: el.top,
          width: el.width,
          height: el.height,
          fill: el.bgColor,
          tag: makeTag("subLaneLabel", el.id),
          lineColor: null,
        });
        await addTextBox(shapes, {
          left: el.left + 0.08,
          top: el.top,
          width: el.width - 0.16,
          height: el.height,
          text: el.text,
          fontSize: el.fontSize,
          fontColor: el.textColor,
          fontFamily: el.font,
          alignment: "left",
          verticalAlignment: "middle",
          bold: el.bold,
          italic: el.italic,
          underline: el.underline,
          tag: makeTag("subLaneLabelText", el.id),
          fill: null,
        });
      }
    }

    // ── Task Bars, Baselines, Milestones ──
    for (const el of layout.tasks) {
      if (el.type === "baselineBar") {
        // Thin bar below the actual task bar (planned dates)
        await addRectangle(shapes, {
          left: el.left,
          top: el.top,
          width: el.width,
          height: el.height,
          fill: el.color,
          tag: makeTag("baseline", el.id),
          lineColor: null,
          opacity: el.opacity,
        });
      } else if (el.type === "taskBar") {
        // Main task bar
        await addRectangle(shapes, {
          left: el.left,
          top: el.top,
          width: el.width,
          height: el.height,
          fill: el.color,
          tag: makeTag("taskbar", el.id),
          cornerRadius: el.cornerRadius,
          lineColor: null,
        });

        // 3D/Gel highlight strip (top 30% lighter)
        if (el.style3d && el.highlightColor) {
          await addRectangle(shapes, {
            left: el.left,
            top: el.top,
            width: el.width,
            height: el.height * 0.35,
            fill: el.highlightColor,
            tag: makeTag("3dHighlight", el.id),
            cornerRadius: el.cornerRadius,
            lineColor: null,
            opacity: 0.5,
          });
        }

        // Percent complete fill (darker overlay on left portion of bar)
        if (el.percentComplete !== null && el.percentWidth !== null && el.percentWidth > 0) {
          await addRectangle(shapes, {
            left: el.left,
            top: el.top,
            width: el.percentWidth,
            height: el.height,
            fill: el.percentColor,
            tag: makeTag("pctBar", el.id),
            cornerRadius: el.cornerRadius,
            lineColor: null,
          });
        }

        // Task label with configurable position
        const labelPos = el.labelPosition || "inside";
        let labelLeft, labelTop, labelWidth, labelHeight, labelColor;
        const labelAlign = el.labelAlign || "left";

        switch (labelPos) {
          case "above":
            labelLeft = el.left;
            labelTop = Math.max(el.top - 0.2, 0);
            labelWidth = Math.max(el.width, 1.2);
            labelHeight = 0.2;
            labelColor = el.color; // use bar color for outside labels
            break;
          case "below":
            labelLeft = el.left;
            labelTop = el.top + el.height + 0.02;
            labelWidth = Math.max(el.width, 1.2);
            labelHeight = 0.2;
            labelColor = el.color;
            break;
          case "left":
            labelWidth = 1.5;
            labelLeft = el.left - labelWidth - 0.05;
            labelTop = el.top;
            labelHeight = el.height;
            labelColor = el.color;
            break;
          case "right":
            labelLeft = el.left + el.width + 0.05;
            labelTop = el.top;
            labelWidth = 1.5;
            labelHeight = el.height;
            labelColor = el.color;
            break;
          case "inside":
          default:
            // Fallback: if bar is too narrow, auto-place on right
            if (el.width <= 1.2) {
              labelLeft = el.left + el.width + 0.05;
              labelTop = el.top;
              labelWidth = 1.5;
              labelHeight = el.height;
              labelColor = el.color;
            } else {
              labelLeft = el.left + 0.06;
              labelTop = el.top;
              labelWidth = el.width - 0.12;
              labelHeight = el.height;
              labelColor = el.textColor;
            }
            break;
        }

        await addTextBox(shapes, {
          left: labelLeft,
          top: labelTop,
          width: labelWidth,
          height: labelHeight,
          text: el.name,
          fontSize: el.fontSize,
          fontColor: labelColor,
          fontFamily: el.font,
          alignment: labelAlign,
          verticalAlignment: "middle",
          bold: el.bold,
          italic: el.italic,
          underline: el.underline,
          tag: makeTag("taskLabel", el.id),
          fill: null,
        });

        // Duration label (to the right of the bar, e.g., "14d")
        if (el.durationLabel) {
          await addTextBox(shapes, {
            left: el.left + el.width + 0.05,
            top: el.top,
            width: 0.5,
            height: el.height,
            text: el.durationLabel,
            fontSize: 6,
            fontColor: el.durationLabelColor,
            fontFamily: el.font,
            alignment: "left",
            verticalAlignment: "middle",
            bold: el.durationBold,
            italic: el.durationItalic,
            underline: el.durationUnderline,
            tag: makeTag("durLabel", el.id),
            fill: null,
          });
        }
      } else if (el.type === "milestone") {
        // Milestone shape
        const shapeType = MILESTONE_SHAPE_MAP[el.milestoneShape] || "Diamond";

        await addGeometricShape(shapes, shapeType, {
          left: el.left,
          top: el.top,
          width: el.size,
          height: el.size,
          fill: el.color,
          tag: makeTag("milestone", el.id),
        });

        // Milestone label (above or below based on labelPosition)
        const labelText = el.dateLabel ? `${el.name}\n${el.dateLabel}` : el.name;
        const msVertAlign = el.labelPosition === "below" ? "top" : "bottom";
        await addTextBox(shapes, {
          left: el.labelLeft,
          top: el.labelTop,
          width: el.labelWidth,
          height: el.size * 1.5,
          text: labelText,
          fontSize: el.fontSize,
          fontColor: el.textColor,
          fontFamily: el.font,
          alignment: "center",
          verticalAlignment: msVertAlign,
          bold: el.bold,
          italic: el.italic,
          underline: el.underline,
          tag: makeTag("msLabel", el.id),
          fill: null,
        });
      }
    }

    // ── Dependency Lines ──
    for (const dep of layout.dependencies) {
      const points = dep.points;
      for (let i = 0; i < points.length - 1; i++) {
        const isLast = i === points.length - 2;

        await addLine(shapes, {
          x1: points[i].x, y1: points[i].y,
          x2: points[i + 1].x, y2: points[i + 1].y,
          color: dep.color,
          weight: dep.weight,
          dashStyle: "solid",
          tag: makeTag("dep", `${dep.fromTaskId}_${dep.toTaskId}_${i}`),
        });

        if (isLast) {
          await addArrowhead(shapes, {
            x: points[i + 1].x,
            y: points[i + 1].y,
            fromX: points[i].x,
            fromY: points[i].y,
            size: dep.arrowSize / 72,
            color: dep.color,
            tag: makeTag("arrow", `${dep.fromTaskId}_${dep.toTaskId}`),
          });
        }
      }
    }

    // ── Today Marker (on top of everything) ──
    if (layout.todayMarker) {
      const tm = layout.todayMarker;
      await addLine(shapes, {
        x1: tm.x, y1: tm.top,
        x2: tm.x, y2: tm.bottom,
        color: tm.color,
        weight: 2,
        dashStyle: "dashDot",
        tag: makeTag("today", "line"),
      });

      // "Today" label at the top
      await addTextBox(shapes, {
        left: tm.x - 0.3,
        top: tm.labelTop,
        width: 0.6,
        height: 0.15,
        text: tm.label,
        fontSize: 7,
        fontColor: tm.color,
        fontFamily: "Segoe UI",
        alignment: "center",
        verticalAlignment: "middle",
        bold: true,
        tag: makeTag("today", "label"),
        fill: null,
      });
    }

    await context.sync();
  });
}

// ── Shape Helper Functions ──

async function addRectangle(shapes, opts) {
  const shapeType = (opts.cornerRadius && opts.cornerRadius > 0) ? "RoundedRectangle" : "Rectangle";
  const shape = shapes.addGeometricShape(shapeType, {
    left: opts.left,
    top: opts.top,
    width: opts.width,
    height: opts.height,
  });

  shape.name = opts.tag;

  if (opts.fill) {
    shape.fill.setSolidColor(opts.fill);
    if (opts.opacity !== undefined && opts.opacity < 1) {
      shape.fill.transparency = 1 - opts.opacity;
    }
  } else {
    shape.fill.clear();
  }

  if (opts.lineColor === null) {
    shape.lineFormat.visible = false;
  } else if (opts.lineColor) {
    shape.lineFormat.color = opts.lineColor;
  }

  return shape;
}

async function addTextBox(shapes, opts) {
  const shape = shapes.addTextBox(opts.text || "", {
    left: opts.left,
    top: opts.top,
    width: opts.width,
    height: opts.height,
  });

  shape.name = opts.tag;

  if (opts.fill) {
    shape.fill.setSolidColor(opts.fill);
  } else {
    shape.fill.clear();
  }

  shape.lineFormat.visible = false;

  const textRange = shape.textFrame.textRange;
  textRange.font.size = opts.fontSize || 10;
  textRange.font.name = opts.fontFamily || "Segoe UI";
  textRange.font.color = opts.fontColor || "#333333";

  if (opts.bold) {
    textRange.font.bold = true;
  }
  if (opts.italic) {
    textRange.font.italic = true;
  }
  if (opts.underline) {
    textRange.font.underline = "Single";
  }

  // PowerPoint for Mac's current Office.js build returns undefined for
  // textRange.paragraphs on a just-added text box. Alignment is a visual
  // nicety, so fall back to default (left) if the API isn't available.
  try {
    const paragraphs = shape.textFrame.textRange.paragraphs;
    if (paragraphs && typeof paragraphs.getItemAt === "function") {
      const paragraph = paragraphs.getItemAt(0);
      if (opts.alignment === "center") {
        paragraph.horizontalAlignment = "Center";
      } else if (opts.alignment === "right") {
        paragraph.horizontalAlignment = "Right";
      } else {
        paragraph.horizontalAlignment = "Left";
      }
    }
  } catch (_) {
    // Alignment API not available on this host; render with default alignment.
  }

  if (opts.verticalAlignment === "middle") {
    shape.textFrame.verticalAlignment = "Middle";
  } else if (opts.verticalAlignment === "bottom") {
    shape.textFrame.verticalAlignment = "Bottom";
  }

  shape.textFrame.autoSizeSetting = "AutoSizeTextToFitShape";

  return shape;
}

async function addLine(shapes, opts) {
  // ShapeCollection.addLine(connectorType, options) takes a ConnectorType
  // ("Straight" | "Elbow" | "Curve") and a ShapeAddOptions geometry object.
  const shape = shapes.addLine("Straight", {
    left: Math.min(opts.x1, opts.x2),
    top: Math.min(opts.y1, opts.y2),
    width: Math.abs(opts.x2 - opts.x1) || 0.001,
    height: Math.abs(opts.y2 - opts.y1) || 0.001,
  });

  shape.name = opts.tag;
  shape.lineFormat.color = opts.color || "#A5A5A5";
  shape.lineFormat.weight = opts.weight || 1;

  if (opts.dashStyle === "dash") {
    shape.lineFormat.dashStyle = "Dash";
  } else if (opts.dashStyle === "dashDot") {
    shape.lineFormat.dashStyle = "DashDot";
  }

  return shape;
}

async function addGeometricShape(shapes, shapeType, opts) {
  const shape = shapes.addGeometricShape(shapeType, {
    left: opts.left,
    top: opts.top,
    width: opts.width,
    height: opts.height,
  });

  shape.name = opts.tag;
  shape.fill.setSolidColor(opts.fill);
  shape.lineFormat.visible = false;

  return shape;
}

async function addArrowhead(shapes, opts) {
  const shape = shapes.addGeometricShape("IsoscelesTriangle", {
    left: opts.x - opts.size / 2,
    top: opts.y - opts.size / 2,
    width: opts.size,
    height: opts.size,
  });

  shape.name = opts.tag;
  shape.fill.setSolidColor(opts.color);
  shape.lineFormat.visible = false;

  const dx = opts.x - opts.fromX;
  const dy = opts.y - opts.fromY;
  const angle = Math.atan2(dy, dx) * (180 / Math.PI);
  shape.rotation = angle + 90;

  return shape;
}

/**
 * Remove all Streamline-generated shapes from the active slide.
 */
async function clearStreamlineShapes() {
  await PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const shapes = slide.shapes;
    shapes.load("items/name");
    await context.sync();

    const toDelete = [];
    for (const shape of shapes.items) {
      if (shape.name && shape.name.startsWith(SHAPE_TAG_PREFIX + ":")) {
        toDelete.push(shape);
      }
    }

    for (const shape of toDelete) {
      shape.delete();
    }

    await context.sync();
  });
}

/**
 * Check if the current slide has any Streamline shapes.
 */
async function hasStreamlineShapes() {
  let found = false;

  await PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const shapes = slide.shapes;
    shapes.load("items/name");
    await context.sync();

    for (const shape of shapes.items) {
      if (shape.name && shape.name.startsWith(SHAPE_TAG_PREFIX + ":")) {
        found = true;
        break;
      }
    }
  });

  return found;
}

module.exports = {
  renderGantt,
  clearStreamlineShapes,
  hasStreamlineShapes,
  SHAPE_TAG_PREFIX,
};
