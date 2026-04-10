/**
 * Streamline Excel Parser
 * Parses Excel schedule files using SheetJS and maps columns to the internal task schema.
 */

const XLSX = require("xlsx");

// Expected column names (case-insensitive, trimmed)
const COLUMN_MAP = {
  "swim lane": "swimLane",
  "swimlane": "swimLane",
  "sub swim lane": "subSwimLane",
  "sub swimlane": "subSwimLane",
  "subswim lane": "subSwimLane",
  "sub-swimlane": "subSwimLane",
  "sub lane": "subSwimLane",
  "task name": "taskName",
  "taskname": "taskName",
  "name": "taskName",
  "type": "type",
  "start date": "startDate",
  "startdate": "startDate",
  "start": "startDate",
  "end date": "endDate",
  "enddate": "endDate",
  "end": "endDate",
  "planned start": "plannedStartDate",
  "planned start date": "plannedStartDate",
  "baseline start": "plannedStartDate",
  "planned end": "plannedEndDate",
  "planned end date": "plannedEndDate",
  "baseline end": "plannedEndDate",
  "% complete": "percentComplete",
  "percent complete": "percentComplete",
  "progress": "percentComplete",
  "complete": "percentComplete",
  "completion": "percentComplete",
  "dependency": "dependency",
  "dependencies": "dependency",
  "depends on": "dependency",
  "predecessors": "dependency",
  "notes": "notes",
  "note": "notes",
  "status": "status",
  "owner": "owner",
  "assigned to": "owner",
  "milestone shape": "milestoneShape",
  "shape": "milestoneShape",
};

const DATE_FIELDS = new Set([
  "startDate", "endDate", "plannedStartDate", "plannedEndDate",
]);

function normalizeColumnName(raw) {
  return raw.toLowerCase().trim();
}

function parseExcelDate(value) {
  if (!value) return null;

  if (value instanceof Date) {
    return value;
  }

  if (typeof value === "number") {
    const date = XLSX.SSF.parse_date_code(value);
    if (date) {
      return new Date(date.y, date.m - 1, date.d);
    }
  }

  if (typeof value === "string") {
    const parsed = new Date(value);
    if (!isNaN(parsed.getTime())) {
      return parsed;
    }
  }

  return null;
}

function parseExcelFile(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: "array", cellDates: true });

  const sheetName = workbook.SheetNames[0];
  if (!sheetName) {
    throw new Error("Excel file contains no worksheets.");
  }

  const sheet = workbook.Sheets[sheetName];
  const rawRows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

  if (rawRows.length === 0) {
    throw new Error("Excel worksheet is empty.");
  }

  // Map raw column headers to internal field names
  const rawHeaders = Object.keys(rawRows[0]);
  const headerMap = {};
  const unmappedHeaders = [];

  for (const header of rawHeaders) {
    const normalized = normalizeColumnName(header);
    if (COLUMN_MAP[normalized]) {
      headerMap[header] = COLUMN_MAP[normalized];
    } else {
      unmappedHeaders.push(header);
    }
  }

  // Parse each row into a normalized object
  const rows = rawRows.map((rawRow, index) => {
    const row = {};

    for (const [rawHeader, fieldName] of Object.entries(headerMap)) {
      let value = rawRow[rawHeader];

      if (DATE_FIELDS.has(fieldName)) {
        value = parseExcelDate(value);
      }

      row[fieldName] = value;
    }

    row._rowIndex = index + 2; // 1-based, accounting for header row
    return row;
  });

  return {
    rows,
    sheetName,
    mappedColumns: Object.values(headerMap),
    unmappedColumns: unmappedHeaders,
  };
}

/**
 * Parse tab-separated clipboard text into rows (for paste from Excel/Google Sheets).
 * @param {string} text - Tab-separated text with header row
 * @returns {Object} { rows, mappedColumns }
 */
function parseClipboardData(text) {
  if (!text || !text.trim()) return { rows: [], mappedColumns: [] };

  const lines = text.trim().split("\n");
  if (lines.length < 2) return { rows: [], mappedColumns: [] };

  const rawHeaders = lines[0].split("\t").map((h) => h.trim());
  const headerMap = {};
  const mappedColumns = [];

  for (const header of rawHeaders) {
    const normalized = normalizeColumnName(header);
    if (COLUMN_MAP[normalized]) {
      headerMap[header] = COLUMN_MAP[normalized];
      mappedColumns.push(COLUMN_MAP[normalized]);
    }
  }

  const rows = [];
  for (let i = 1; i < lines.length; i++) {
    const cells = lines[i].split("\t");
    const row = {};

    for (let j = 0; j < rawHeaders.length; j++) {
      const fieldName = headerMap[rawHeaders[j]];
      if (!fieldName) continue;

      let value = j < cells.length ? cells[j].trim() : "";

      if (DATE_FIELDS.has(fieldName) && value) {
        const parsed = new Date(value);
        value = isNaN(parsed.getTime()) ? null : parsed;
      }

      row[fieldName] = value;
    }

    row._rowIndex = i + 1;
    if (row.taskName || row.swimLane) rows.push(row);
  }

  return { rows, mappedColumns };
}

module.exports = {
  parseExcelFile,
  parseExcelDate,
  parseClipboardData,
  COLUMN_MAP,
};
