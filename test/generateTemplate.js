/**
 * Generate a blank Streamline Excel template for PMs.
 * Run: node test/generateTemplate.js
 */

const XLSX = require("xlsx");
const path = require("path");

const headers = [
  "Swim Lane",
  "Task Name",
  "Type",
  "Start Date",
  "End Date",
  "Dependency",
  "Status",
  "Owner",
  "Notes",
];

// Example rows to show PMs the expected format
const exampleRows = [
  {
    "Swim Lane": "Workstream 1",
    "Task Name": "Kickoff Meeting",
    "Type": "Milestone",
    "Start Date": new Date(2026, 3, 13),
    "End Date": "",
    "Dependency": "",
    "Status": "Complete",
    "Owner": "PM",
    "Notes": "Project kickoff",
  },
  {
    "Swim Lane": "Workstream 1",
    "Task Name": "Requirements Gathering",
    "Type": "Task",
    "Start Date": new Date(2026, 3, 14),
    "End Date": new Date(2026, 3, 24),
    "Dependency": "Kickoff Meeting",
    "Status": "On Track",
    "Owner": "Team A",
    "Notes": "",
  },
  {
    "Swim Lane": "Workstream 1",
    "Task Name": "Design Review",
    "Type": "Milestone",
    "Start Date": new Date(2026, 3, 25),
    "End Date": "",
    "Dependency": "Requirements Gathering",
    "Status": "",
    "Owner": "PM",
    "Notes": "",
  },
  {
    "Swim Lane": "Workstream 2",
    "Task Name": "Development Phase 1",
    "Type": "Task",
    "Start Date": new Date(2026, 3, 25),
    "End Date": new Date(2026, 4, 15),
    "Dependency": "Design Review",
    "Status": "",
    "Owner": "Team B",
    "Notes": "",
  },
  {
    "Swim Lane": "Workstream 2",
    "Task Name": "Testing",
    "Type": "Task",
    "Start Date": new Date(2026, 4, 10),
    "End Date": new Date(2026, 4, 22),
    "Dependency": "",
    "Status": "",
    "Owner": "QA",
    "Notes": "",
  },
  {
    "Swim Lane": "Program",
    "Task Name": "Release",
    "Type": "Milestone",
    "Start Date": new Date(2026, 4, 25),
    "End Date": "",
    "Dependency": "Development Phase 1, Testing",
    "Status": "",
    "Owner": "PM",
    "Notes": "GA target",
  },
];

const ws = XLSX.utils.json_to_sheet(exampleRows, { header: headers });

// Set column widths for readability
ws["!cols"] = [
  { wch: 18 }, // Swim Lane
  { wch: 28 }, // Task Name
  { wch: 10 }, // Type
  { wch: 14 }, // Start Date
  { wch: 14 }, // End Date
  { wch: 30 }, // Dependency
  { wch: 12 }, // Status
  { wch: 14 }, // Owner
  { wch: 30 }, // Notes
];

const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, "Schedule");

const outPath = path.join(__dirname, "..", "assets", "Streamline_Template.xlsx");
XLSX.writeFile(wb, outPath);
console.log(`Created: ${outPath}`);
