/**
 * Streamline Data Editor
 * Built-in spreadsheet-like editor for creating/editing tasks without Excel.
 * Renders an editable HTML table in the task pane.
 */

const COLUMNS = [
  { key: "swimLane", label: "Swim Lane", width: 100, type: "text" },
  { key: "taskName", label: "Task Name", width: 120, type: "text" },
  { key: "type", label: "Type", width: 75, type: "select", options: ["Task", "Milestone"] },
  { key: "startDate", label: "Start", width: 95, type: "date" },
  { key: "endDate", label: "End", width: 95, type: "date" },
  { key: "status", label: "Status", width: 85, type: "select", options: ["", "On Track", "At Risk", "Delayed", "Complete"] },
  { key: "percentComplete", label: "%", width: 40, type: "number" },
  { key: "dependency", label: "Dependency", width: 120, type: "text" },
  { key: "subSwimLane", label: "Sub Lane", width: 90, type: "text" },
  { key: "plannedStartDate", label: "Plan Start", width: 95, type: "date" },
  { key: "plannedEndDate", label: "Plan End", width: 95, type: "date" },
  { key: "milestoneShape", label: "Shape", width: 70, type: "select", options: ["", "diamond", "circle", "triangle", "star", "flag", "square"] },
  { key: "owner", label: "Owner", width: 80, type: "text" },
  { key: "notes", label: "Notes", width: 100, type: "text" },
];

const MINIMAL_COLUMNS = ["swimLane", "taskName", "type", "startDate", "endDate", "status", "percentComplete", "dependency"];

class DataEditor {
  constructor(container) {
    this.container = container;
    this.rows = [];
    this.showAllColumns = false;
    this.onChangeCallback = null;
  }

  /**
   * Initialize the editor with optional existing rows.
   */
  init(rows = []) {
    this.rows = rows.length > 0 ? rows.map((r) => ({ ...r })) : [this.createEmptyRow()];
    this.render();
  }

  /**
   * Load parsed Excel/XML data into the editor.
   */
  loadRows(parsedRows) {
    this.rows = parsedRows.map((r) => {
      const row = {};
      for (const col of COLUMNS) {
        let val = r[col.key];
        if (val instanceof Date) {
          val = formatDate(val);
        }
        row[col.key] = val !== undefined && val !== null ? String(val) : "";
      }
      return row;
    });
    if (this.rows.length === 0) this.rows.push(this.createEmptyRow());
    this.render();
  }

  /**
   * Get rows as objects ready for the data model pipeline.
   */
  getRows() {
    return this.rows
      .filter((r) => r.taskName && r.taskName.trim())
      .map((r) => {
        const out = {};
        for (const col of COLUMNS) {
          let val = r[col.key] || "";
          if (col.type === "date" && val) {
            const d = new Date(val);
            out[col.key] = isNaN(d.getTime()) ? null : d;
          } else if (col.type === "number" && val) {
            out[col.key] = parseFloat(val) || null;
          } else {
            out[col.key] = val;
          }
        }
        return out;
      });
  }

  onChange(callback) {
    this.onChangeCallback = callback;
  }

  createEmptyRow() {
    const row = {};
    for (const col of COLUMNS) row[col.key] = "";
    row.type = "Task";
    return row;
  }

  addRow(index = -1) {
    const newRow = this.createEmptyRow();
    if (index < 0 || index >= this.rows.length) {
      this.rows.push(newRow);
    } else {
      this.rows.splice(index + 1, 0, newRow);
    }
    this.render();
  }

  deleteRow(index) {
    if (this.rows.length <= 1) return;
    this.rows.splice(index, 1);
    this.render();
    this.notifyChange();
  }

  duplicateRow(index) {
    const copy = { ...this.rows[index] };
    copy.taskName = copy.taskName + " (copy)";
    this.rows.splice(index + 1, 0, copy);
    this.render();
    this.notifyChange();
  }

  moveRow(from, to) {
    if (to < 0 || to >= this.rows.length) return;
    const [row] = this.rows.splice(from, 1);
    this.rows.splice(to, 0, row);
    this.render();
    this.notifyChange();
  }

  notifyChange() {
    if (this.onChangeCallback) this.onChangeCallback(this.getRows());
  }

  toggleColumns() {
    this.showAllColumns = !this.showAllColumns;
    this.render();
  }

  render() {
    const visibleCols = this.showAllColumns
      ? COLUMNS
      : COLUMNS.filter((c) => MINIMAL_COLUMNS.includes(c.key));

    let html = `
      <div class="de-toolbar">
        <button class="de-btn" data-action="add-row" title="Add row">+ Row</button>
        <button class="de-btn" data-action="toggle-cols">
          ${this.showAllColumns ? "Fewer Columns" : "More Columns"}
        </button>
        <span class="de-count">${this.rows.length} rows</span>
      </div>
      <div class="de-table-wrap">
        <table class="de-table">
          <thead>
            <tr>
              <th class="de-th-num">#</th>
              ${visibleCols.map((c) => `<th style="min-width:${c.width}px">${c.label}</th>`).join("")}
              <th class="de-th-actions"></th>
            </tr>
          </thead>
          <tbody>
    `;

    for (let i = 0; i < this.rows.length; i++) {
      const row = this.rows[i];
      html += `<tr data-row="${i}">`;
      html += `<td class="de-num">${i + 1}</td>`;

      for (const col of visibleCols) {
        const val = row[col.key] || "";
        if (col.type === "select") {
          html += `<td><select class="de-input de-select" data-row="${i}" data-col="${col.key}">`;
          for (const opt of col.options) {
            const sel = val === opt ? "selected" : "";
            html += `<option value="${opt}" ${sel}>${opt || "-"}</option>`;
          }
          html += `</select></td>`;
        } else if (col.type === "date") {
          html += `<td><input type="date" class="de-input" data-row="${i}" data-col="${col.key}" value="${val}" /></td>`;
        } else if (col.type === "number") {
          html += `<td><input type="number" class="de-input de-num-input" data-row="${i}" data-col="${col.key}" value="${val}" min="0" max="100" /></td>`;
        } else {
          html += `<td><input type="text" class="de-input" data-row="${i}" data-col="${col.key}" value="${escapeAttr(val)}" /></td>`;
        }
      }

      html += `<td class="de-actions">
        <button class="de-btn-sm" data-action="dup" data-row="${i}" title="Duplicate">&#x2398;</button>
        <button class="de-btn-sm de-btn-up" data-action="move-up" data-row="${i}" title="Move up">&uarr;</button>
        <button class="de-btn-sm de-btn-down" data-action="move-down" data-row="${i}" title="Move down">&darr;</button>
        <button class="de-btn-sm de-btn-del" data-action="delete" data-row="${i}" title="Delete">&times;</button>
      </td>`;
      html += `</tr>`;
    }

    html += `</tbody></table></div>`;
    this.container.innerHTML = html;
    this.bindTableEvents();
  }

  bindTableEvents() {
    // Input changes
    this.container.querySelectorAll(".de-input").forEach((input) => {
      input.addEventListener("change", (e) => {
        const row = parseInt(e.target.dataset.row, 10);
        const col = e.target.dataset.col;
        this.rows[row][col] = e.target.value;
        this.notifyChange();
      });
    });

    // Toolbar actions
    this.container.querySelectorAll("[data-action]").forEach((btn) => {
      btn.addEventListener("click", (e) => {
        const action = e.target.closest("[data-action]").dataset.action;
        const rowIdx = parseInt(e.target.closest("[data-row]")?.dataset.row, 10);

        switch (action) {
          case "add-row": this.addRow(); break;
          case "toggle-cols": this.toggleColumns(); break;
          case "delete": this.deleteRow(rowIdx); break;
          case "dup": this.duplicateRow(rowIdx); break;
          case "move-up": this.moveRow(rowIdx, rowIdx - 1); break;
          case "move-down": this.moveRow(rowIdx, rowIdx + 1); break;
        }
      });
    });
  }
}

function formatDate(d) {
  if (!d) return "";
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

function escapeAttr(str) {
  return str.replace(/"/g, "&quot;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

/**
 * Parse clipboard text (tab-separated) into row objects.
 * Handles paste from Excel, Google Sheets, etc.
 */
function parseClipboardText(text) {
  if (!text || !text.trim()) return [];

  const lines = text.trim().split("\n");
  if (lines.length < 2) return []; // need header + at least 1 data row

  // First line is header
  const headers = lines[0].split("\t").map((h) => h.trim().toLowerCase());

  // Map headers to column keys
  const HEADER_MAP = {};
  for (const col of COLUMNS) {
    HEADER_MAP[col.label.toLowerCase()] = col.key;
    HEADER_MAP[col.key.toLowerCase()] = col.key;
  }
  // Extra aliases
  Object.assign(HEADER_MAP, {
    "swim lane": "swimLane", "swimlane": "swimLane",
    "task name": "taskName", "taskname": "taskName", "name": "taskName",
    "start date": "startDate", "start": "startDate",
    "end date": "endDate", "end": "endDate",
    "% complete": "percentComplete", "percent complete": "percentComplete",
    "progress": "percentComplete", "completion": "percentComplete",
    "depends on": "dependency", "dependencies": "dependency", "predecessors": "dependency",
    "sub swim lane": "subSwimLane", "sub lane": "subSwimLane",
    "planned start": "plannedStartDate", "plan start": "plannedStartDate", "baseline start": "plannedStartDate",
    "planned end": "plannedEndDate", "plan end": "plannedEndDate", "baseline end": "plannedEndDate",
    "milestone shape": "milestoneShape", "shape": "milestoneShape",
    "assigned to": "owner",
  });

  const colMap = headers.map((h) => HEADER_MAP[h] || null);

  const rows = [];
  for (let i = 1; i < lines.length; i++) {
    const cells = lines[i].split("\t");
    const row = {};
    for (const col of COLUMNS) row[col.key] = "";

    for (let j = 0; j < cells.length; j++) {
      const key = colMap[j];
      if (key) row[key] = cells[j].trim();
    }

    if (row.taskName) rows.push(row);
  }

  return rows;
}

module.exports = { DataEditor, parseClipboardText, COLUMNS };
