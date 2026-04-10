/**
 * Streamline Template Manager
 * Loads, stores, and applies organizational templates (color schemes, shape styles, font configs).
 */

/**
 * Template categories for grouping templates in the UI.
 */
const TEMPLATE_CATEGORIES = {
  "project-management": { label: "Project Management", order: 1 },
  "professional": { label: "Professional", order: 2 },
  "creative": { label: "Creative", order: 3 },
  "minimal": { label: "Minimal", order: 4 },
  "industry": { label: "Industry", order: 5 },
};

/**
 * Default text styles applied per element type.
 * Per-element bold/italic/underline with sensible defaults.
 */
const DEFAULT_TEXT_STYLES = {
  swimLaneLabel: { bold: true, italic: false, underline: false },
  subSwimLaneLabel: { bold: false, italic: false, underline: false },
  taskLabel: { bold: false, italic: false, underline: false },
  milestoneLabel: { bold: false, italic: false, underline: false },
  milestoneDateLabel: { bold: false, italic: true, underline: false },
  yearLabel: { bold: true, italic: false, underline: false },
  monthLabel: { bold: false, italic: false, underline: false },
  tier3Label: { bold: false, italic: false, underline: false },
  durationLabel: { bold: false, italic: false, underline: false },
  todayLabel: { bold: true, italic: false, underline: false },
  title: { bold: true, italic: false, underline: false },
};

const DEFAULT_TEMPLATES = {
  standard: {
    name: "Standard",
    category: "project-management",
    colors: {
      taskBar: "#5B7FC7",
      milestone: "#7B68EE",
      criticalPath: "#FF4444",
      dependency: "#B0B0B0",
      swimLaneHeader: "#4A4A4A",
      swimLaneHeaderText: "#FFFFFF",
      subSwimLaneHeader: "#5A5A5A",
      subSwimLaneHeaderText: "#FFFFFF",
      yearAxisBg: "#2D3748",
      yearAxisText: "#FFFFFF",
      monthAxisBg: "#4A5568",
      monthAxisText: "#E2E8F0",
      tier3AxisBg: "#5A6A7E",
      tier3AxisText: "#E2E8F0",
      yearBoundary: "#718096",
      timeAxisBg: "#4A5568",
      timeAxisText: "#E2E8F0",
      statusOnTrack: "#48BB78",
      statusAtRisk: "#ECC94B",
      statusDelayed: "#FC8181",
      statusComplete: "#A0AEC0",
      background: "#FFFFFF",
      gridLine: "#E8E8E8",
      taskText: "#FFFFFF",
      labelText: "#2D3748",
      milestoneLabelText: "#4A5568",
      // New feature colors
      todayMarker: "#FF0000",
      todayMarkerLabel: "#FF0000",
      elapsedShading: "#F7F7F7",
      percentComplete: "#3A5A9C",
      baselineBar: "#C0C0C0",
      durationLabel: "#718096",
      varianceEarly: "#38A169",
      varianceLate: "#E53E3E",
    },
    fonts: {
      primary: "Segoe UI",
      sizes: {
        swimLaneLabel: 10,
        taskLabel: 8,
        milestoneLabel: 8,
        timeAxis: 8,
        title: 14,
        durationLabel: 7,
      },
      styles: JSON.parse(JSON.stringify(DEFAULT_TEXT_STYLES)),
    },
    shapes: {
      milestoneShape: "diamond",
      taskBarCornerRadius: 2,
      taskBarHeight: 20,
      milestoneSize: 12,
      dependencyLineWeight: 1.5,
      criticalPathLineWeight: 2.5,
      arrowSize: 6,
    },
    labelConfig: {
      taskLabelPosition: "inside",      // inside, above, below, left, right
      milestoneLabelPosition: "above",  // above, below
      taskLabelAlign: "left",           // left, center, right
      showTaskDates: false,             // append start/end to task label
      showOwner: false,                 // append owner name to task label
      labelWrap: false,                 // wrap long labels
    },
  },

  highContrast: {
    name: "High Contrast",
    category: "professional",
    colors: {
      taskBar: "#2B6CB0",
      milestone: "#D53F8C",
      criticalPath: "#E53E3E",
      dependency: "#718096",
      swimLaneHeader: "#1A202C",
      swimLaneHeaderText: "#FFFFFF",
      subSwimLaneHeader: "#2D3748",
      subSwimLaneHeaderText: "#FFFFFF",
      yearAxisBg: "#1A202C",
      yearAxisText: "#FFFFFF",
      monthAxisBg: "#2D3748",
      monthAxisText: "#FFFFFF",
      tier3AxisBg: "#3D4A5C",
      tier3AxisText: "#FFFFFF",
      yearBoundary: "#4A5568",
      timeAxisBg: "#2D3748",
      timeAxisText: "#FFFFFF",
      statusOnTrack: "#38A169",
      statusAtRisk: "#D69E2E",
      statusDelayed: "#E53E3E",
      statusComplete: "#718096",
      background: "#FFFFFF",
      gridLine: "#CBD5E0",
      taskText: "#FFFFFF",
      labelText: "#000000",
      milestoneLabelText: "#2D3748",
      todayMarker: "#E53E3E",
      todayMarkerLabel: "#E53E3E",
      elapsedShading: "#EDF2F7",
      percentComplete: "#1A365D",
      baselineBar: "#A0AEC0",
      durationLabel: "#4A5568",
      varianceEarly: "#276749",
      varianceLate: "#C53030",
    },
    fonts: {
      primary: "Segoe UI",
      sizes: {
        swimLaneLabel: 11,
        taskLabel: 9,
        milestoneLabel: 9,
        timeAxis: 9,
        title: 16,
        durationLabel: 7,
      },
    },
    shapes: {
      milestoneShape: "diamond",
      taskBarCornerRadius: 0,
      taskBarHeight: 22,
      milestoneSize: 14,
      dependencyLineWeight: 2,
      criticalPathLineWeight: 3,
      arrowSize: 7,
    },
  },

  minimal: {
    name: "Minimal",
    category: "minimal",
    colors: {
      taskBar: "#718096",
      milestone: "#4A5568",
      criticalPath: "#2D3748",
      dependency: "#CBD5E0",
      swimLaneHeader: "#E2E8F0",
      swimLaneHeaderText: "#2D3748",
      subSwimLaneHeader: "#EDF2F7",
      subSwimLaneHeaderText: "#4A5568",
      yearAxisBg: "#E2E8F0",
      yearAxisText: "#2D3748",
      monthAxisBg: "#EDF2F7",
      monthAxisText: "#4A5568",
      tier3AxisBg: "#F7FAFC",
      tier3AxisText: "#718096",
      yearBoundary: "#CBD5E0",
      timeAxisBg: "#EDF2F7",
      timeAxisText: "#4A5568",
      statusOnTrack: "#718096",
      statusAtRisk: "#A0AEC0",
      statusDelayed: "#4A5568",
      statusComplete: "#CBD5E0",
      background: "#FFFFFF",
      gridLine: "#EDF2F7",
      taskText: "#FFFFFF",
      labelText: "#2D3748",
      milestoneLabelText: "#4A5568",
      todayMarker: "#4A5568",
      todayMarkerLabel: "#4A5568",
      elapsedShading: "#FAFAFA",
      percentComplete: "#4A5568",
      baselineBar: "#E2E8F0",
      durationLabel: "#A0AEC0",
      varianceEarly: "#718096",
      varianceLate: "#4A5568",
    },
    fonts: {
      primary: "Segoe UI",
      sizes: {
        swimLaneLabel: 9,
        taskLabel: 7,
        milestoneLabel: 7,
        timeAxis: 7,
        title: 12,
        durationLabel: 6,
      },
    },
    shapes: {
      milestoneShape: "diamond",
      taskBarCornerRadius: 0,
      taskBarHeight: 18,
      milestoneSize: 10,
      dependencyLineWeight: 1,
      criticalPathLineWeight: 2,
      arrowSize: 5,
    },
  },
  corporate: {
    name: "Corporate",
    category: "professional",
    colors: {
      taskBar: "#2C5282", milestone: "#2B6CB0", criticalPath: "#C53030",
      dependency: "#A0AEC0", swimLaneHeader: "#2A4365", swimLaneHeaderText: "#FFFFFF",
      subSwimLaneHeader: "#2C5282", subSwimLaneHeaderText: "#FFFFFF",
      yearAxisBg: "#1A365D", yearAxisText: "#FFFFFF", monthAxisBg: "#2A4365",
      monthAxisText: "#E2E8F0", tier3AxisBg: "#2C5282", tier3AxisText: "#E2E8F0",
      yearBoundary: "#4A5568", timeAxisBg: "#2A4365", timeAxisText: "#E2E8F0",
      statusOnTrack: "#2F855A", statusAtRisk: "#C05621", statusDelayed: "#C53030",
      statusComplete: "#718096", background: "#FFFFFF", gridLine: "#E2E8F0",
      taskText: "#FFFFFF", labelText: "#1A365D", milestoneLabelText: "#2A4365",
      todayMarker: "#C53030", todayMarkerLabel: "#C53030", elapsedShading: "#EBF4FF",
      percentComplete: "#1A365D", baselineBar: "#BEE3F8", durationLabel: "#4A5568",
      varianceEarly: "#22543D", varianceLate: "#9B2C2C",
    },
    fonts: { primary: "Calibri", sizes: { swimLaneLabel: 10, taskLabel: 8, milestoneLabel: 8, timeAxis: 8, title: 14, durationLabel: 7 } },
    shapes: { milestoneShape: "diamond", taskBarCornerRadius: 3, taskBarHeight: 20, milestoneSize: 12, dependencyLineWeight: 1.5, criticalPathLineWeight: 2.5, arrowSize: 6 },
  },

  ocean: {
    name: "Ocean",
    category: "creative",
    colors: {
      taskBar: "#0694A2", milestone: "#047481", criticalPath: "#E53E3E",
      dependency: "#B2DFDB", swimLaneHeader: "#014451", swimLaneHeaderText: "#FFFFFF",
      subSwimLaneHeader: "#065666", subSwimLaneHeaderText: "#FFFFFF",
      yearAxisBg: "#014451", yearAxisText: "#E0F2F1", monthAxisBg: "#065666",
      monthAxisText: "#E0F2F1", tier3AxisBg: "#0E7C86", tier3AxisText: "#E0F2F1",
      yearBoundary: "#4DB6AC", timeAxisBg: "#065666", timeAxisText: "#E0F2F1",
      statusOnTrack: "#38B2AC", statusAtRisk: "#ED8936", statusDelayed: "#FC8181",
      statusComplete: "#A0AEC0", background: "#FFFFFF", gridLine: "#E0F2F1",
      taskText: "#FFFFFF", labelText: "#014451", milestoneLabelText: "#065666",
      todayMarker: "#E53E3E", todayMarkerLabel: "#E53E3E", elapsedShading: "#E6FFFA",
      percentComplete: "#047481", baselineBar: "#B2DFDB", durationLabel: "#4DB6AC",
      varianceEarly: "#234E52", varianceLate: "#C53030",
    },
    fonts: { primary: "Segoe UI", sizes: { swimLaneLabel: 10, taskLabel: 8, milestoneLabel: 8, timeAxis: 8, title: 14, durationLabel: 7 } },
    shapes: { milestoneShape: "diamond", taskBarCornerRadius: 4, taskBarHeight: 20, milestoneSize: 12, dependencyLineWeight: 1.5, criticalPathLineWeight: 2.5, arrowSize: 6 },
  },

  sunset: {
    name: "Sunset",
    category: "creative",
    colors: {
      taskBar: "#DD6B20", milestone: "#C05621", criticalPath: "#E53E3E",
      dependency: "#FBD38D", swimLaneHeader: "#7B341E", swimLaneHeaderText: "#FFFFFF",
      subSwimLaneHeader: "#9C4221", subSwimLaneHeaderText: "#FFFFFF",
      yearAxisBg: "#652B19", yearAxisText: "#FFFAF0", monthAxisBg: "#7B341E",
      monthAxisText: "#FEEBC8", tier3AxisBg: "#9C4221", tier3AxisText: "#FEEBC8",
      yearBoundary: "#C05621", timeAxisBg: "#7B341E", timeAxisText: "#FEEBC8",
      statusOnTrack: "#38A169", statusAtRisk: "#ECC94B", statusDelayed: "#E53E3E",
      statusComplete: "#A0AEC0", background: "#FFFAF0", gridLine: "#FEEBC8",
      taskText: "#FFFFFF", labelText: "#652B19", milestoneLabelText: "#7B341E",
      todayMarker: "#E53E3E", todayMarkerLabel: "#E53E3E", elapsedShading: "#FFF5EB",
      percentComplete: "#9C4221", baselineBar: "#FBD38D", durationLabel: "#C05621",
      varianceEarly: "#276749", varianceLate: "#9B2C2C",
    },
    fonts: { primary: "Segoe UI", sizes: { swimLaneLabel: 10, taskLabel: 8, milestoneLabel: 8, timeAxis: 8, title: 14, durationLabel: 7 } },
    shapes: { milestoneShape: "diamond", taskBarCornerRadius: 2, taskBarHeight: 20, milestoneSize: 12, dependencyLineWeight: 1.5, criticalPathLineWeight: 2.5, arrowSize: 6 },
  },

  forest: {
    name: "Forest",
    category: "creative",
    colors: {
      taskBar: "#276749", milestone: "#22543D", criticalPath: "#E53E3E",
      dependency: "#C6F6D5", swimLaneHeader: "#1C4532", swimLaneHeaderText: "#FFFFFF",
      subSwimLaneHeader: "#22543D", subSwimLaneHeaderText: "#FFFFFF",
      yearAxisBg: "#1C4532", yearAxisText: "#F0FFF4", monthAxisBg: "#22543D",
      monthAxisText: "#C6F6D5", tier3AxisBg: "#276749", tier3AxisText: "#C6F6D5",
      yearBoundary: "#48BB78", timeAxisBg: "#22543D", timeAxisText: "#C6F6D5",
      statusOnTrack: "#48BB78", statusAtRisk: "#ECC94B", statusDelayed: "#FC8181",
      statusComplete: "#9AE6B4", background: "#FFFFFF", gridLine: "#C6F6D5",
      taskText: "#FFFFFF", labelText: "#1C4532", milestoneLabelText: "#22543D",
      todayMarker: "#E53E3E", todayMarkerLabel: "#E53E3E", elapsedShading: "#F0FFF4",
      percentComplete: "#1C4532", baselineBar: "#9AE6B4", durationLabel: "#48BB78",
      varianceEarly: "#22543D", varianceLate: "#C53030",
    },
    fonts: { primary: "Segoe UI", sizes: { swimLaneLabel: 10, taskLabel: 8, milestoneLabel: 8, timeAxis: 8, title: 14, durationLabel: 7 } },
    shapes: { milestoneShape: "diamond", taskBarCornerRadius: 2, taskBarHeight: 20, milestoneSize: 12, dependencyLineWeight: 1.5, criticalPathLineWeight: 2.5, arrowSize: 6 },
  },

  slate: {
    name: "Slate",
    category: "minimal",
    colors: {
      taskBar: "#4A5568", milestone: "#2D3748", criticalPath: "#E53E3E",
      dependency: "#CBD5E0", swimLaneHeader: "#1A202C", swimLaneHeaderText: "#F7FAFC",
      subSwimLaneHeader: "#2D3748", subSwimLaneHeaderText: "#F7FAFC",
      yearAxisBg: "#171923", yearAxisText: "#F7FAFC", monthAxisBg: "#1A202C",
      monthAxisText: "#E2E8F0", tier3AxisBg: "#2D3748", tier3AxisText: "#E2E8F0",
      yearBoundary: "#4A5568", timeAxisBg: "#1A202C", timeAxisText: "#E2E8F0",
      statusOnTrack: "#68D391", statusAtRisk: "#F6E05E", statusDelayed: "#FC8181",
      statusComplete: "#A0AEC0", background: "#F7FAFC", gridLine: "#E2E8F0",
      taskText: "#FFFFFF", labelText: "#1A202C", milestoneLabelText: "#2D3748",
      todayMarker: "#E53E3E", todayMarkerLabel: "#E53E3E", elapsedShading: "#EDF2F7",
      percentComplete: "#2D3748", baselineBar: "#A0AEC0", durationLabel: "#718096",
      varianceEarly: "#276749", varianceLate: "#C53030",
    },
    fonts: { primary: "Segoe UI", sizes: { swimLaneLabel: 10, taskLabel: 8, milestoneLabel: 8, timeAxis: 8, title: 14, durationLabel: 7 } },
    shapes: { milestoneShape: "diamond", taskBarCornerRadius: 0, taskBarHeight: 20, milestoneSize: 12, dependencyLineWeight: 1.5, criticalPathLineWeight: 2.5, arrowSize: 6 },
  },

  royal: {
    name: "Royal Purple",
    category: "creative",
    colors: {
      taskBar: "#6B46C1", milestone: "#553C9A", criticalPath: "#E53E3E",
      dependency: "#D6BCFA", swimLaneHeader: "#44337A", swimLaneHeaderText: "#FFFFFF",
      subSwimLaneHeader: "#553C9A", subSwimLaneHeaderText: "#FFFFFF",
      yearAxisBg: "#322659", yearAxisText: "#FAF5FF", monthAxisBg: "#44337A",
      monthAxisText: "#E9D8FD", tier3AxisBg: "#553C9A", tier3AxisText: "#E9D8FD",
      yearBoundary: "#805AD5", timeAxisBg: "#44337A", timeAxisText: "#E9D8FD",
      statusOnTrack: "#48BB78", statusAtRisk: "#ECC94B", statusDelayed: "#FC8181",
      statusComplete: "#B794F4", background: "#FFFFFF", gridLine: "#E9D8FD",
      taskText: "#FFFFFF", labelText: "#322659", milestoneLabelText: "#44337A",
      todayMarker: "#E53E3E", todayMarkerLabel: "#E53E3E", elapsedShading: "#FAF5FF",
      percentComplete: "#44337A", baselineBar: "#D6BCFA", durationLabel: "#805AD5",
      varianceEarly: "#276749", varianceLate: "#C53030",
    },
    fonts: { primary: "Segoe UI", sizes: { swimLaneLabel: 10, taskLabel: 8, milestoneLabel: 8, timeAxis: 8, title: 14, durationLabel: 7 } },
    shapes: { milestoneShape: "diamond", taskBarCornerRadius: 4, taskBarHeight: 20, milestoneSize: 12, dependencyLineWeight: 1.5, criticalPathLineWeight: 2.5, arrowSize: 6 },
  },

  crimson: {
    name: "Crimson",
    category: "industry",
    colors: {
      taskBar: "#C53030", milestone: "#9B2C2C", criticalPath: "#2D3748",
      dependency: "#FEB2B2", swimLaneHeader: "#742A2A", swimLaneHeaderText: "#FFFFFF",
      subSwimLaneHeader: "#9B2C2C", subSwimLaneHeaderText: "#FFFFFF",
      yearAxisBg: "#63171B", yearAxisText: "#FFF5F5", monthAxisBg: "#742A2A",
      monthAxisText: "#FED7D7", tier3AxisBg: "#9B2C2C", tier3AxisText: "#FED7D7",
      yearBoundary: "#C53030", timeAxisBg: "#742A2A", timeAxisText: "#FED7D7",
      statusOnTrack: "#48BB78", statusAtRisk: "#ECC94B", statusDelayed: "#2D3748",
      statusComplete: "#FC8181", background: "#FFFFFF", gridLine: "#FED7D7",
      taskText: "#FFFFFF", labelText: "#63171B", milestoneLabelText: "#742A2A",
      todayMarker: "#2D3748", todayMarkerLabel: "#2D3748", elapsedShading: "#FFF5F5",
      percentComplete: "#742A2A", baselineBar: "#FEB2B2", durationLabel: "#C53030",
      varianceEarly: "#276749", varianceLate: "#63171B",
    },
    fonts: { primary: "Segoe UI", sizes: { swimLaneLabel: 10, taskLabel: 8, milestoneLabel: 8, timeAxis: 8, title: 14, durationLabel: 7 } },
    shapes: { milestoneShape: "diamond", taskBarCornerRadius: 2, taskBarHeight: 20, milestoneSize: 12, dependencyLineWeight: 1.5, criticalPathLineWeight: 2.5, arrowSize: 6 },
  },

  pastel: {
    name: "Pastel",
    category: "creative",
    colors: {
      taskBar: "#90CDF4", milestone: "#FBB6CE", criticalPath: "#FC8181",
      dependency: "#E2E8F0", swimLaneHeader: "#BEE3F8", swimLaneHeaderText: "#2A4365",
      subSwimLaneHeader: "#C3DAFE", subSwimLaneHeaderText: "#434190",
      yearAxisBg: "#EBF8FF", yearAxisText: "#2A4365", monthAxisBg: "#E6FFFA",
      monthAxisText: "#234E52", tier3AxisBg: "#FEFCBF", tier3AxisText: "#744210",
      yearBoundary: "#BEE3F8", timeAxisBg: "#EBF8FF", timeAxisText: "#2A4365",
      statusOnTrack: "#9AE6B4", statusAtRisk: "#FEFCBF", statusDelayed: "#FEB2B2",
      statusComplete: "#E2E8F0", background: "#FFFFFF", gridLine: "#EDF2F7",
      taskText: "#2D3748", labelText: "#2D3748", milestoneLabelText: "#553C9A",
      todayMarker: "#FC8181", todayMarkerLabel: "#FC8181", elapsedShading: "#F7FAFC",
      percentComplete: "#63B3ED", baselineBar: "#E2E8F0", durationLabel: "#A0AEC0",
      varianceEarly: "#68D391", varianceLate: "#FC8181",
    },
    fonts: { primary: "Segoe UI", sizes: { swimLaneLabel: 10, taskLabel: 8, milestoneLabel: 8, timeAxis: 8, title: 14, durationLabel: 7 } },
    shapes: { milestoneShape: "circle", taskBarCornerRadius: 6, taskBarHeight: 20, milestoneSize: 12, dependencyLineWeight: 1, criticalPathLineWeight: 2, arrowSize: 5 },
  },

  darkMode: {
    name: "Dark Mode",
    category: "minimal",
    colors: {
      taskBar: "#63B3ED", milestone: "#F6AD55", criticalPath: "#FC8181",
      dependency: "#4A5568", swimLaneHeader: "#2D3748", swimLaneHeaderText: "#E2E8F0",
      subSwimLaneHeader: "#4A5568", subSwimLaneHeaderText: "#E2E8F0",
      yearAxisBg: "#171923", yearAxisText: "#E2E8F0", monthAxisBg: "#1A202C",
      monthAxisText: "#CBD5E0", tier3AxisBg: "#2D3748", tier3AxisText: "#CBD5E0",
      yearBoundary: "#4A5568", timeAxisBg: "#1A202C", timeAxisText: "#CBD5E0",
      statusOnTrack: "#68D391", statusAtRisk: "#F6E05E", statusDelayed: "#FC8181",
      statusComplete: "#718096", background: "#1A202C", gridLine: "#2D3748",
      taskText: "#1A202C", labelText: "#E2E8F0", milestoneLabelText: "#CBD5E0",
      todayMarker: "#FC8181", todayMarkerLabel: "#FC8181", elapsedShading: "#2D3748",
      percentComplete: "#3182CE", baselineBar: "#4A5568", durationLabel: "#718096",
      varianceEarly: "#68D391", varianceLate: "#FC8181",
    },
    fonts: { primary: "Segoe UI", sizes: { swimLaneLabel: 10, taskLabel: 8, milestoneLabel: 8, timeAxis: 8, title: 14, durationLabel: 7 } },
    shapes: { milestoneShape: "diamond", taskBarCornerRadius: 3, taskBarHeight: 20, milestoneSize: 12, dependencyLineWeight: 1.5, criticalPathLineWeight: 2.5, arrowSize: 6 },
  },
};

class TemplateManager {
  constructor() {
    this.templates = JSON.parse(JSON.stringify(DEFAULT_TEMPLATES));
    // Ensure every template has textStyles and labelConfig
    for (const key of Object.keys(this.templates)) {
      this.ensureTextStyles(this.templates[key]);
    }
    this.activeTemplateKey = "standard";
  }

  getTemplate(key) {
    return this.templates[key] || null;
  }

  getActiveTemplate() {
    return this.templates[this.activeTemplateKey];
  }

  setActiveTemplate(key) {
    if (!this.templates[key]) {
      throw new Error(`Template "${key}" not found.`);
    }
    this.activeTemplateKey = key;
  }

  listTemplates() {
    return Object.entries(this.templates).map(([key, tmpl]) => ({
      key,
      name: tmpl.name,
      category: tmpl.category || "professional",
    }));
  }

  /**
   * List templates grouped by category.
   * @returns {Array} [{ category, label, templates: [{ key, name }] }]
   */
  listTemplatesByCategory() {
    const grouped = new Map();

    for (const [key, tmpl] of Object.entries(this.templates)) {
      const cat = tmpl.category || "professional";
      if (!grouped.has(cat)) grouped.set(cat, []);
      grouped.get(cat).push({ key, name: tmpl.name });
    }

    const result = [];
    for (const [catKey, catInfo] of Object.entries(TEMPLATE_CATEGORIES)) {
      if (grouped.has(catKey)) {
        result.push({
          category: catKey,
          label: catInfo.label,
          order: catInfo.order,
          templates: grouped.get(catKey),
        });
      }
    }

    // Any templates in unknown categories → add them to "professional"
    for (const [cat, items] of grouped) {
      if (!TEMPLATE_CATEGORIES[cat]) {
        const existing = result.find((r) => r.category === "professional");
        if (existing) existing.templates.push(...items);
      }
    }

    result.sort((a, b) => a.order - b.order);
    return result;
  }

  getCategories() {
    return TEMPLATE_CATEGORIES;
  }

  importTemplate(key, templateJson) {
    const required = ["name", "colors", "fonts", "shapes"];
    for (const field of required) {
      if (!templateJson[field]) {
        throw new Error(`Template is missing required field: "${field}".`);
      }
    }

    const base = JSON.parse(JSON.stringify(DEFAULT_TEMPLATES.standard));
    const merged = {
      name: templateJson.name,
      category: templateJson.category || "professional",
      colors: { ...base.colors, ...templateJson.colors },
      fonts: {
        primary: templateJson.fonts.primary || base.fonts.primary,
        sizes: { ...base.fonts.sizes, ...(templateJson.fonts.sizes || {}) },
        styles: { ...base.fonts.styles, ...(templateJson.fonts.styles || {}) },
      },
      shapes: { ...base.shapes, ...templateJson.shapes },
      labelConfig: { ...base.labelConfig, ...(templateJson.labelConfig || {}) },
    };

    this.templates[key] = merged;
    return merged;
  }

  /**
   * Merge default styles into an existing template (for backward compat).
   */
  ensureTextStyles(template) {
    if (!template.fonts) template.fonts = {};
    if (!template.fonts.styles) {
      template.fonts.styles = JSON.parse(JSON.stringify(DEFAULT_TEXT_STYLES));
    }
    if (!template.labelConfig) {
      template.labelConfig = JSON.parse(JSON.stringify(DEFAULT_TEMPLATES.standard.labelConfig));
    }
    return template;
  }

  /**
   * Export a template as a JSON object for sharing.
   */
  exportTemplate(key) {
    const tmpl = this.templates[key];
    if (!tmpl) throw new Error(`Template "${key}" not found.`);
    return JSON.parse(JSON.stringify(tmpl));
  }

  getStatusColor(template, status) {
    const map = {
      ON_TRACK: template.colors.statusOnTrack,
      AT_RISK: template.colors.statusAtRisk,
      DELAYED: template.colors.statusDelayed,
      COMPLETE: template.colors.statusComplete,
    };
    return map[status] || template.colors.taskBar;
  }
}

module.exports = {
  TemplateManager,
  DEFAULT_TEMPLATES,
  TEMPLATE_CATEGORIES,
  DEFAULT_TEXT_STYLES,
};
