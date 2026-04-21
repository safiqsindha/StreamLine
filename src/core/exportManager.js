/**
 * Streamline Export Manager
 * Renders the Gantt layout to an HTML5 Canvas for PNG and PDF export.
 * Works in the task pane web view (browser context).
 */

/**
 * Render layout to a Canvas element.
 * @param {Object} layout - The layout object from calculateLayout()
 * @param {Object} template - Active template
 * @param {Object} opts - { scale: 2 (default), width, height }
 * @returns {HTMLCanvasElement}
 */
function renderToCanvas(layout, template, opts = {}) {
  const scale = opts.scale || 2;
  const slideW = layout.slide.width; // inches
  const slideH = layout.slide.height;
  const dpi = 96; // screen DPI
  const pxW = Math.round(slideW * dpi * scale);
  const pxH = Math.round(slideH * dpi * scale);

  const canvas = document.createElement("canvas");
  canvas.width = pxW;
  canvas.height = pxH;
  const ctx = canvas.getContext("2d");

  // Scale context so we can draw in inches * dpi
  const s = dpi * scale;
  ctx.scale(s, s);

  // Background
  ctx.fillStyle = template.colors.background || "#FFFFFF";
  ctx.fillRect(0, 0, slideW, slideH);

  // Weekend shading
  if (layout.weekendShading && layout.weekendShading.length > 0) {
    ctx.globalAlpha = 0.5;
    for (const wk of layout.weekendShading) {
      ctx.fillStyle = wk.color;
      ctx.fillRect(wk.left, wk.top, wk.width, wk.height);
    }
    ctx.globalAlpha = 1;
  }

  // Elapsed shading
  if (layout.elapsedShading) {
    const el = layout.elapsedShading;
    ctx.fillStyle = el.color;
    ctx.globalAlpha = 0.5;
    ctx.fillRect(el.left, el.top, el.width, el.height);
    ctx.globalAlpha = 1;
  }

  // Time axis backgrounds
  for (const el of layout.timeAxis) {
    if (el.type === "timeAxisBg") {
      ctx.fillStyle = el.color;
      ctx.fillRect(el.left, el.top, el.width, el.height);
    }
  }

  // Time axis labels
  for (const el of layout.timeAxis) {
    if (el.type === "yearLabel" || el.type === "monthLabel" || el.type === "tier3Label") {
      ctx.fillStyle = el.textColor;
      ctx.font = buildFontStr(el);
      ctx.textAlign = "center";
      ctx.textBaseline = "middle";
      ctx.fillText(el.text, el.left + el.width / 2, el.top + el.height / 2, el.width);
    } else if (el.type === "gridLine" || el.type === "yearBoundary") {
      ctx.strokeStyle = el.color;
      ctx.lineWidth = el.type === "yearBoundary" ? 1 / s : 0.5 / s;
      ctx.beginPath();
      ctx.moveTo(el.x, el.top);
      ctx.lineTo(el.x, el.bottom);
      ctx.stroke();
    }
  }

  // Lane separators
  for (const sep of layout.laneSeparators) {
    ctx.strokeStyle = sep.color;
    ctx.lineWidth = 0.5 / s;
    ctx.beginPath();
    ctx.moveTo(sep.x1, sep.y);
    ctx.lineTo(sep.x2, sep.y);
    ctx.stroke();
  }

  // Lane labels
  for (const el of layout.laneLabels) {
    ctx.fillStyle = el.bgColor;
    ctx.fillRect(el.left, el.top, el.width, el.height);
    ctx.fillStyle = el.textColor;
    ctx.font = `bold ${el.fontSize}pt ${el.font}`;
    ctx.textAlign = "left";
    ctx.textBaseline = "middle";
    ctx.fillText(el.text, el.left + 0.08, el.top + el.height / 2, el.width - 0.16);
  }

  // Sub-lane labels
  for (const el of layout.subLaneLabels || []) {
    ctx.fillStyle = el.bgColor;
    ctx.fillRect(el.left, el.top, el.width, el.height);
    ctx.fillStyle = el.textColor;
    ctx.font = `${el.fontSize}pt ${el.font}`;
    ctx.textAlign = "left";
    ctx.textBaseline = "middle";
    ctx.fillText(el.text, el.left + 0.08, el.top + el.height / 2, el.width - 0.16);
  }

  // Task bars, baselines, milestones
  for (const el of layout.tasks) {
    if (el.type === "baselineBar") {
      ctx.globalAlpha = el.opacity || 0.6;
      ctx.fillStyle = el.color;
      drawRoundRect(ctx, el.left, el.top, el.width, el.height, (el.cornerRadius || 0) / 72);
      ctx.fill();
      ctx.globalAlpha = 1;
    } else if (el.type === "taskBar") {
      const r = (el.cornerRadius || 0) / 72;
      // Main bar
      ctx.fillStyle = el.color;
      drawRoundRect(ctx, el.left, el.top, el.width, el.height, r);
      ctx.fill();

      // Percent complete overlay
      if (el.percentWidth && el.percentWidth > 0) {
        ctx.fillStyle = el.percentColor;
        drawRoundRect(ctx, el.left, el.top, el.percentWidth, el.height, r);
        ctx.fill();
      }

      // Task label
      const labelInside = el.width > 1.2;
      ctx.fillStyle = labelInside ? el.textColor : el.color;
      ctx.font = `${el.fontSize}pt ${el.font}`;
      ctx.textAlign = "left";
      ctx.textBaseline = "middle";
      const labelX = labelInside ? el.left + 0.06 : el.left + el.width + 0.05;
      ctx.fillText(el.name, labelX, el.top + el.height / 2, labelInside ? el.width - 0.12 : 1.5);

      // Duration label
      if (el.durationLabel) {
        ctx.fillStyle = el.durationLabelColor || "#888";
        ctx.font = `6pt ${el.font}`;
        ctx.fillText(el.durationLabel, el.left + el.width + 0.05, el.top + el.height / 2);
      }
    } else if (el.type === "milestone") {
      ctx.fillStyle = el.color;
      drawMilestoneShape(ctx, el);

      // Label above
      ctx.fillStyle = el.textColor;
      ctx.font = `${el.fontSize}pt ${el.font}`;
      ctx.textAlign = "center";
      ctx.textBaseline = "bottom";
      ctx.fillText(el.name, el.centerX, el.top - 0.02, el.labelWidth);
      if (el.dateLabel) {
        ctx.font = `${Math.max(el.fontSize - 1, 5)}pt ${el.font}`;
        ctx.fillText(el.dateLabel, el.centerX, el.top + el.size + 0.12, el.labelWidth);
      }
    }
  }

  // Dependency lines
  for (const dep of layout.dependencies) {
    ctx.strokeStyle = dep.color;
    ctx.lineWidth = dep.weight / 72;
    ctx.beginPath();
    for (let i = 0; i < dep.points.length; i++) {
      const p = dep.points[i];
      if (i === 0) ctx.moveTo(p.x, p.y);
      else ctx.lineTo(p.x, p.y);
    }
    ctx.stroke();

    // Arrowhead
    const last = dep.points[dep.points.length - 1];
    const prev = dep.points[dep.points.length - 2];
    const aSize = (dep.arrowSize || 6) / 72;
    const angle = Math.atan2(last.y - prev.y, last.x - prev.x);
    ctx.fillStyle = dep.color;
    ctx.beginPath();
    ctx.moveTo(last.x, last.y);
    ctx.lineTo(last.x - aSize * Math.cos(angle - 0.4), last.y - aSize * Math.sin(angle - 0.4));
    ctx.lineTo(last.x - aSize * Math.cos(angle + 0.4), last.y - aSize * Math.sin(angle + 0.4));
    ctx.closePath();
    ctx.fill();
  }

  // Today marker
  if (layout.todayMarker) {
    const tm = layout.todayMarker;
    ctx.strokeStyle = tm.color;
    ctx.lineWidth = 2 / s;
    ctx.setLineDash([4 / s, 2 / s]);
    ctx.beginPath();
    ctx.moveTo(tm.x, tm.top);
    ctx.lineTo(tm.x, tm.bottom);
    ctx.stroke();
    ctx.setLineDash([]);

    ctx.fillStyle = tm.color;
    ctx.font = `bold 7pt sans-serif`;
    ctx.textAlign = "center";
    ctx.textBaseline = "bottom";
    ctx.fillText(tm.label, tm.x, tm.labelTop);
  }

  return canvas;
}

function buildFontStr(el) {
  const parts = [];
  if (el.italic) parts.push("italic");
  if (el.bold) parts.push("bold");
  parts.push(`${el.fontSize}pt`);
  parts.push(el.font || "sans-serif");
  return parts.join(" ");
}

function drawRoundRect(ctx, x, y, w, h, r) {
  r = Math.min(r, w / 2, h / 2);
  ctx.beginPath();
  ctx.moveTo(x + r, y);
  ctx.lineTo(x + w - r, y);
  ctx.quadraticCurveTo(x + w, y, x + w, y + r);
  ctx.lineTo(x + w, y + h - r);
  ctx.quadraticCurveTo(x + w, y + h, x + w - r, y + h);
  ctx.lineTo(x + r, y + h);
  ctx.quadraticCurveTo(x, y + h, x, y + h - r);
  ctx.lineTo(x, y + r);
  ctx.quadraticCurveTo(x, y, x + r, y);
  ctx.closePath();
}

function drawMilestoneShape(ctx, el) {
  const cx = el.centerX;
  const cy = el.centerY;
  const s = el.size / 2;

  ctx.beginPath();
  switch (el.milestoneShape) {
    case "circle":
      ctx.arc(cx, cy, s, 0, Math.PI * 2);
      break;
    case "triangle":
      ctx.moveTo(cx, cy - s);
      ctx.lineTo(cx + s, cy + s);
      ctx.lineTo(cx - s, cy + s);
      break;
    case "star":
      for (let i = 0; i < 5; i++) {
        const outerAngle = (i * 72 - 90) * Math.PI / 180;
        const innerAngle = ((i * 72 + 36) - 90) * Math.PI / 180;
        const ox = cx + s * Math.cos(outerAngle);
        const oy = cy + s * Math.sin(outerAngle);
        const ix = cx + s * 0.4 * Math.cos(innerAngle);
        const iy = cy + s * 0.4 * Math.sin(innerAngle);
        if (i === 0) ctx.moveTo(ox, oy);
        else ctx.lineTo(ox, oy);
        ctx.lineTo(ix, iy);
      }
      break;
    case "flag":
      ctx.moveTo(cx - s * 0.3, cy - s);
      ctx.lineTo(cx + s, cy - s * 0.3);
      ctx.lineTo(cx - s * 0.3, cy + s * 0.3);
      ctx.lineTo(cx - s * 0.3, cy + s);
      ctx.lineTo(cx - s * 0.5, cy + s);
      ctx.lineTo(cx - s * 0.5, cy - s);
      break;
    case "square":
      ctx.rect(cx - s, cy - s, s * 2, s * 2);
      break;
    default: // diamond
      ctx.moveTo(cx, cy - s);
      ctx.lineTo(cx + s, cy);
      ctx.lineTo(cx, cy + s);
      ctx.lineTo(cx - s, cy);
      break;
  }
  ctx.closePath();
  ctx.fill();
}

/**
 * Export layout as PNG data URL.
 */
function exportPNG(layout, template, opts = {}) {
  const canvas = renderToCanvas(layout, template, opts);
  return canvas.toDataURL("image/png");
}

/**
 * Export layout as JPG data URL (quality 0-1).
 */
function exportJPG(layout, template, opts = {}) {
  const canvas = renderToCanvas(layout, template, opts);
  const quality = opts.quality !== undefined ? opts.quality : 0.92;
  // JPG needs an opaque background
  const ctx = canvas.getContext("2d");
  const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
  // Composite onto white background
  const bgCanvas = document.createElement("canvas");
  bgCanvas.width = canvas.width;
  bgCanvas.height = canvas.height;
  const bgCtx = bgCanvas.getContext("2d");
  bgCtx.fillStyle = template.colors.background || "#FFFFFF";
  bgCtx.fillRect(0, 0, bgCanvas.width, bgCanvas.height);
  bgCtx.drawImage(canvas, 0, 0);
  return bgCanvas.toDataURL("image/jpeg", quality);
}

/**
 * Download layout as PNG file.
 */
function downloadPNG(layout, template, filename = "streamline_gantt.png", opts = {}) {
  const dataUrl = exportPNG(layout, template, opts);
  const a = document.createElement("a");
  a.href = dataUrl;
  a.download = filename;
  a.click();
}

/**
 * Download layout as JPG file.
 */
function downloadJPG(layout, template, filename = "streamline_gantt.jpg", opts = {}) {
  const dataUrl = exportJPG(layout, template, opts);
  const a = document.createElement("a");
  a.href = dataUrl;
  a.download = filename;
  a.click();
}

/**
 * Export layout as PDF by rendering to canvas and wrapping in a minimal PDF.
 * Uses a lightweight approach without external dependencies.
 */
function downloadPDF(layout, template, filename = "streamline_gantt.pdf", opts = {}) {
  const canvas = renderToCanvas(layout, template, { ...opts, scale: opts.scale || 3 });
  const imgData = canvas.toDataURL("image/jpeg", 0.95);

  // Open in new window for printing as PDF
  const printWin = window.open("", "_blank");
  if (!printWin) {
    throw new Error(
      "PDF export was blocked by the browser. Allow pop-ups for this add-in, or use Export PNG/JPG instead."
    );
  }
  printWin.document.write(`
    <!DOCTYPE html>
    <html>
    <head>
      <title>${filename}</title>
      <style>
        @page { size: landscape; margin: 0; }
        body { margin: 0; display: flex; justify-content: center; align-items: center; }
        img { width: 100%; height: auto; }
      </style>
    </head>
    <body>
      <img src="${imgData}" />
      <script>
        window.onload = function() { window.print(); };
      </script>
    </body>
    </html>
  `);
  printWin.document.close();
}

module.exports = {
  renderToCanvas,
  exportPNG,
  exportJPG,
  downloadPNG,
  downloadJPG,
  downloadPDF,
};
