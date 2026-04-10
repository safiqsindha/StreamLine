/**
 * Generate sample Excel files for testing Streamline.
 * Run: node test/generateTestData.js
 */

const XLSX = require("xlsx");
const path = require("path");

function createWorkbook(rows) {
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Schedule");
  return wb;
}

function date(y, m, d) {
  return new Date(y, m - 1, d);
}

// ── Test File 1: Realistic NPI Schedule (10 swim lanes, all features) ──

const npiSchedule = [
  // BIOS Qualification
  { "Swim Lane": "BIOS Qualification", "Task Name": "BIOS Code Freeze", "Type": "Milestone", "Start Date": date(2026, 4, 14), "End Date": "", "Dependency": "", "Status": "On Track", "Owner": "BIOS Team", "Notes": "All BIOS changes locked", "% Complete": "", "Planned Start": "", "Planned End": "", "Sub Swim Lane": "" },
  { "Swim Lane": "BIOS Qualification", "Task Name": "BIOS Functional Testing", "Type": "Task", "Start Date": date(2026, 4, 15), "End Date": date(2026, 5, 9), "Dependency": "BIOS Code Freeze", "Status": "On Track", "Owner": "BIOS QA", "Notes": "", "% Complete": 45, "Planned Start": date(2026, 4, 14), "Planned End": date(2026, 5, 7), "Sub Swim Lane": "" },
  { "Swim Lane": "BIOS Qualification", "Task Name": "BIOS Stress Testing", "Type": "Task", "Start Date": date(2026, 5, 4), "End Date": date(2026, 5, 22), "Dependency": "", "Status": "On Track", "Owner": "BIOS QA", "Notes": "72-hour burn-in cycles", "% Complete": 20, "Planned Start": date(2026, 5, 1), "Planned End": date(2026, 5, 20), "Sub Swim Lane": "" },
  { "Swim Lane": "BIOS Qualification", "Task Name": "BIOS Sign-Off", "Type": "Milestone", "Start Date": date(2026, 5, 23), "End Date": "", "Dependency": "BIOS Stress Testing", "Status": "", "Owner": "BIOS Lead", "Notes": "", "% Complete": "", "Planned Start": date(2026, 5, 21), "Planned End": "", "Sub Swim Lane": "" },

  // SKU Validation
  { "Swim Lane": "SKU Validation", "Task Name": "SKU Matrix Finalization", "Type": "Milestone", "Start Date": date(2026, 4, 10), "End Date": "", "Dependency": "", "Status": "Complete", "Owner": "PM", "Notes": "12 SKUs confirmed", "% Complete": 100, "Planned Start": "", "Planned End": "", "Sub Swim Lane": "" },
  { "Swim Lane": "SKU Validation", "Task Name": "SKU Config Testing", "Type": "Task", "Start Date": date(2026, 4, 14), "End Date": date(2026, 5, 15), "Dependency": "SKU Matrix Finalization", "Status": "On Track", "Owner": "Validation Team", "Notes": "", "% Complete": 60, "Planned Start": date(2026, 4, 14), "Planned End": date(2026, 5, 12), "Sub Swim Lane": "" },
  { "Swim Lane": "SKU Validation", "Task Name": "SKU Power Characterization", "Type": "Task", "Start Date": date(2026, 4, 21), "End Date": date(2026, 5, 9), "Dependency": "", "Status": "At Risk", "Owner": "Power Team", "Notes": "Lab capacity constrained", "% Complete": 30, "Planned Start": date(2026, 4, 18), "Planned End": date(2026, 5, 5), "Sub Swim Lane": "" },
  { "Swim Lane": "SKU Validation", "Task Name": "SKU Validation Complete", "Type": "Milestone", "Start Date": date(2026, 5, 16), "End Date": "", "Dependency": "SKU Config Testing", "Status": "", "Owner": "PM", "Notes": "", "% Complete": "", "Planned Start": "", "Planned End": "", "Sub Swim Lane": "" },

  // Thermal Validation
  { "Swim Lane": "Thermal Validation", "Task Name": "Thermal Design Review", "Type": "Milestone", "Start Date": date(2026, 4, 11), "End Date": "", "Dependency": "", "Status": "Complete", "Owner": "Thermal Lead", "Notes": "", "% Complete": 100, "Planned Start": "", "Planned End": "", "Sub Swim Lane": "" },
  { "Swim Lane": "Thermal Validation", "Task Name": "Thermal Simulation", "Type": "Task", "Start Date": date(2026, 4, 14), "End Date": date(2026, 4, 30), "Dependency": "Thermal Design Review", "Status": "On Track", "Owner": "Thermal Team", "Notes": "CFD modeling", "% Complete": 80, "Planned Start": date(2026, 4, 14), "Planned End": date(2026, 4, 28), "Sub Swim Lane": "" },
  { "Swim Lane": "Thermal Validation", "Task Name": "Thermal Chamber Testing", "Type": "Task", "Start Date": date(2026, 5, 1), "End Date": date(2026, 5, 20), "Dependency": "Thermal Simulation", "Status": "", "Owner": "Thermal Team", "Notes": "", "% Complete": 0, "Planned Start": date(2026, 4, 29), "Planned End": date(2026, 5, 18), "Sub Swim Lane": "" },

  // Silicon Bring-Up (with sub-swimlanes)
  { "Swim Lane": "Silicon Bring-Up", "Task Name": "A0 Silicon Arrival", "Type": "Milestone", "Start Date": date(2026, 4, 9), "End Date": "", "Dependency": "", "Status": "Complete", "Owner": "Fab Ops", "Notes": "Wafers received", "% Complete": 100, "Planned Start": "", "Planned End": "", "Sub Swim Lane": "" },
  { "Swim Lane": "Silicon Bring-Up", "Task Name": "Initial Power-On", "Type": "Task", "Start Date": date(2026, 4, 10), "End Date": date(2026, 4, 18), "Dependency": "A0 Silicon Arrival", "Status": "Complete", "Owner": "Silicon Team", "Notes": "", "% Complete": 100, "Planned Start": date(2026, 4, 10), "Planned End": date(2026, 4, 17), "Sub Swim Lane": "Core Validation" },
  { "Swim Lane": "Silicon Bring-Up", "Task Name": "Core Functional Validation", "Type": "Task", "Start Date": date(2026, 4, 19), "End Date": date(2026, 5, 16), "Dependency": "Initial Power-On", "Status": "On Track", "Owner": "Silicon Team", "Notes": "", "% Complete": 35, "Planned Start": date(2026, 4, 18), "Planned End": date(2026, 5, 14), "Sub Swim Lane": "Core Validation" },
  { "Swim Lane": "Silicon Bring-Up", "Task Name": "Speed Path Characterization", "Type": "Task", "Start Date": date(2026, 5, 5), "End Date": date(2026, 5, 30), "Dependency": "", "Status": "", "Owner": "Speed Team", "Notes": "Fmax binning", "% Complete": 10, "Planned Start": date(2026, 5, 3), "Planned End": date(2026, 5, 28), "Sub Swim Lane": "Speed Team" },

  // Platform Integration
  { "Swim Lane": "Platform Integration", "Task Name": "Board Schematic Review", "Type": "Milestone", "Start Date": date(2026, 4, 12), "End Date": "", "Dependency": "", "Status": "Complete", "Owner": "HW Lead", "Notes": "", "% Complete": 100, "Planned Start": "", "Planned End": "", "Sub Swim Lane": "" },
  { "Swim Lane": "Platform Integration", "Task Name": "CRB Assembly", "Type": "Task", "Start Date": date(2026, 4, 14), "End Date": date(2026, 4, 28), "Dependency": "Board Schematic Review", "Status": "On Track", "Owner": "Board Team", "Notes": "Customer Reference Board", "% Complete": 90, "Planned Start": date(2026, 4, 14), "Planned End": date(2026, 4, 25), "Sub Swim Lane": "" },
  { "Swim Lane": "Platform Integration", "Task Name": "Platform Debug", "Type": "Task", "Start Date": date(2026, 4, 29), "End Date": date(2026, 5, 16), "Dependency": "CRB Assembly", "Status": "", "Owner": "Debug Team", "Notes": "", "% Complete": 0, "Planned Start": date(2026, 4, 26), "Planned End": date(2026, 5, 14), "Sub Swim Lane": "" },
  { "Swim Lane": "Platform Integration", "Task Name": "Platform Integration Complete", "Type": "Milestone", "Start Date": date(2026, 5, 17), "End Date": "", "Dependency": "Platform Debug", "Status": "", "Owner": "PM", "Notes": "", "% Complete": "", "Planned Start": "", "Planned End": "", "Sub Swim Lane": "" },

  // Memory Qualification
  { "Swim Lane": "Memory Qualification", "Task Name": "DDR5 Vendor Sampling", "Type": "Task", "Start Date": date(2026, 4, 10), "End Date": date(2026, 4, 24), "Dependency": "", "Status": "On Track", "Owner": "Memory Team", "Notes": "3 vendors", "% Complete": 75, "Planned Start": date(2026, 4, 10), "Planned End": date(2026, 4, 22), "Sub Swim Lane": "" },
  { "Swim Lane": "Memory Qualification", "Task Name": "Memory Compatibility Testing", "Type": "Task", "Start Date": date(2026, 4, 25), "End Date": date(2026, 5, 16), "Dependency": "DDR5 Vendor Sampling", "Status": "", "Owner": "Memory Team", "Notes": "", "% Complete": 0, "Planned Start": date(2026, 4, 23), "Planned End": date(2026, 5, 14), "Sub Swim Lane": "" },
  { "Swim Lane": "Memory Qualification", "Task Name": "Memory Qual Sign-Off", "Type": "Milestone", "Start Date": date(2026, 5, 17), "End Date": "", "Dependency": "Memory Compatibility Testing", "Status": "", "Owner": "Memory Lead", "Notes": "", "% Complete": "", "Planned Start": "", "Planned End": "", "Sub Swim Lane": "" },

  // Firmware (with dependency types)
  { "Swim Lane": "Firmware", "Task Name": "FW Feature Complete", "Type": "Milestone", "Start Date": date(2026, 4, 18), "End Date": "", "Dependency": "", "Status": "Delayed", "Owner": "FW Lead", "Notes": "2 days behind", "% Complete": "", "Planned Start": date(2026, 4, 16), "Planned End": "", "Sub Swim Lane": "" },
  { "Swim Lane": "Firmware", "Task Name": "FW Integration Testing", "Type": "Task", "Start Date": date(2026, 4, 21), "End Date": date(2026, 5, 9), "Dependency": "FW Feature Complete [FS+2d]", "Status": "At Risk", "Owner": "FW QA", "Notes": "", "% Complete": 15, "Planned Start": date(2026, 4, 18), "Planned End": date(2026, 5, 6), "Sub Swim Lane": "" },
  { "Swim Lane": "Firmware", "Task Name": "FW Regression Suite", "Type": "Task", "Start Date": date(2026, 5, 10), "End Date": date(2026, 5, 23), "Dependency": "FW Integration Testing [FF]", "Status": "", "Owner": "FW QA", "Notes": "", "% Complete": 0, "Planned Start": date(2026, 5, 7), "Planned End": date(2026, 5, 20), "Sub Swim Lane": "" },

  // Power Delivery
  { "Swim Lane": "Power Delivery", "Task Name": "VR Design Validation", "Type": "Task", "Start Date": date(2026, 4, 14), "End Date": date(2026, 5, 2), "Dependency": "", "Status": "On Track", "Owner": "Power Team", "Notes": "Voltage regulator testing", "% Complete": 55, "Planned Start": date(2026, 4, 14), "Planned End": date(2026, 5, 1), "Sub Swim Lane": "" },
  { "Swim Lane": "Power Delivery", "Task Name": "Load Line Optimization", "Type": "Task", "Start Date": date(2026, 5, 3), "End Date": date(2026, 5, 16), "Dependency": "VR Design Validation [SS+3d]", "Status": "", "Owner": "Power Team", "Notes": "", "% Complete": 0, "Planned Start": date(2026, 5, 1), "Planned End": date(2026, 5, 14), "Sub Swim Lane": "" },
  { "Swim Lane": "Power Delivery", "Task Name": "Power Delivery Qual", "Type": "Milestone", "Start Date": date(2026, 5, 17), "End Date": "", "Dependency": "Load Line Optimization", "Status": "", "Owner": "Power Lead", "Notes": "", "% Complete": "", "Planned Start": "", "Planned End": "", "Sub Swim Lane": "" },

  // Compliance
  { "Swim Lane": "Compliance", "Task Name": "EMI Pre-Scan", "Type": "Task", "Start Date": date(2026, 4, 21), "End Date": date(2026, 5, 2), "Dependency": "", "Status": "", "Owner": "Compliance Team", "Notes": "", "% Complete": 0, "Planned Start": date(2026, 4, 20), "Planned End": date(2026, 5, 1), "Sub Swim Lane": "" },
  { "Swim Lane": "Compliance", "Task Name": "FCC/CE Certification Submission", "Type": "Task", "Start Date": date(2026, 5, 5), "End Date": date(2026, 5, 30), "Dependency": "EMI Pre-Scan", "Status": "", "Owner": "Compliance Team", "Notes": "", "% Complete": 0, "Planned Start": date(2026, 5, 2), "Planned End": date(2026, 5, 28), "Sub Swim Lane": "" },

  // Program Milestones
  { "Swim Lane": "Program Milestones", "Task Name": "PRQ (Product Release Qualification)", "Type": "Milestone", "Start Date": date(2026, 5, 25), "End Date": "", "Dependency": "BIOS Sign-Off, SKU Validation Complete, Platform Integration Complete", "Status": "", "Owner": "Program PM", "Notes": "Gate review required", "% Complete": "", "Planned Start": "", "Planned End": "", "Sub Swim Lane": "", "Milestone Shape": "star" },
  { "Swim Lane": "Program Milestones", "Task Name": "GA (General Availability)", "Type": "Milestone", "Start Date": date(2026, 6, 8), "End Date": "", "Dependency": "PRQ (Product Release Qualification)", "Status": "", "Owner": "Program PM", "Notes": "", "% Complete": "", "Planned Start": "", "Planned End": "", "Sub Swim Lane": "", "Milestone Shape": "flag" },
];

// ── Test File 2: Minimal (3 swim lanes - smoke test) ──

const minimalSchedule = [
  { "Swim Lane": "Design", "Task Name": "Requirements", "Type": "Task", "Start Date": date(2026, 4, 10), "End Date": date(2026, 4, 17), "Dependency": "", "Status": "Complete", "Owner": "Alice", "Notes": "", "% Complete": 100 },
  { "Swim Lane": "Design", "Task Name": "Design Review", "Type": "Milestone", "Start Date": date(2026, 4, 18), "End Date": "", "Dependency": "Requirements", "Status": "", "Owner": "Alice", "Notes": "" },
  { "Swim Lane": "Development", "Task Name": "Sprint 1", "Type": "Task", "Start Date": date(2026, 4, 18), "End Date": date(2026, 5, 1), "Dependency": "Design Review", "Status": "On Track", "Owner": "Bob", "Notes": "", "% Complete": 50 },
  { "Swim Lane": "Development", "Task Name": "Sprint 2", "Type": "Task", "Start Date": date(2026, 5, 2), "End Date": date(2026, 5, 15), "Dependency": "Sprint 1", "Status": "", "Owner": "Bob", "Notes": "" },
  { "Swim Lane": "Testing", "Task Name": "QA Testing", "Type": "Task", "Start Date": date(2026, 5, 5), "End Date": date(2026, 5, 20), "Dependency": "", "Status": "", "Owner": "Carol", "Notes": "", "% Complete": 25 },
  { "Swim Lane": "Testing", "Task Name": "Release", "Type": "Milestone", "Start Date": date(2026, 5, 22), "End Date": "", "Dependency": "Sprint 2, QA Testing", "Status": "", "Owner": "PM", "Notes": "" },
];

// ── Test File 3: Stress test (15 swim lanes) ──

const stressSchedule = [];
const laneNames = [
  "BIOS Qualification", "SKU Validation", "Thermal Validation",
  "Silicon Bring-Up", "Platform Integration", "Memory Qualification",
  "Firmware", "Power Delivery", "Compliance", "Signal Integrity",
  "Security Audit", "Manufacturing Test", "OS Enablement",
  "Debug Tools", "Program Milestones"
];

for (let i = 0; i < laneNames.length; i++) {
  const lane = laneNames[i];
  const baseDate = date(2026, 4, 10 + i);

  stressSchedule.push({
    "Swim Lane": lane,
    "Task Name": `${lane} Kickoff`,
    "Type": "Milestone",
    "Start Date": baseDate,
    "End Date": "",
    "Dependency": "",
    "Status": i < 3 ? "Complete" : i < 7 ? "On Track" : "",
    "Owner": `Team ${i + 1}`,
    "Notes": "",
    "% Complete": i < 3 ? 100 : "",
  });

  for (let j = 1; j <= 5; j++) {
    const start = new Date(baseDate);
    start.setDate(start.getDate() + j * 5);
    const end = new Date(start);
    end.setDate(end.getDate() + 8 + Math.floor(Math.random() * 7));

    stressSchedule.push({
      "Swim Lane": lane,
      "Task Name": `${lane} Phase ${j}`,
      "Type": "Task",
      "Start Date": start,
      "End Date": end,
      "Dependency": j === 1 ? `${lane} Kickoff` : `${lane} Phase ${j - 1}`,
      "Status": j <= 2 && i < 5 ? "On Track" : "",
      "Owner": `Team ${i + 1}`,
      "Notes": "",
      "% Complete": j <= 2 && i < 5 ? Math.floor(Math.random() * 80) + 10 : 0,
    });
  }

  stressSchedule.push({
    "Swim Lane": lane,
    "Task Name": `${lane} Complete`,
    "Type": "Milestone",
    "Start Date": new Date(baseDate.getTime() + 45 * 24 * 60 * 60 * 1000),
    "End Date": "",
    "Dependency": `${lane} Phase 5`,
    "Status": "",
    "Owner": `Lead ${i + 1}`,
    "Notes": "",
  });
}

// ── Test File 4: Invalid data (for validator testing) ──

const invalidSchedule = [
  { "Swim Lane": "", "Task Name": "Missing Lane", "Type": "Task", "Start Date": date(2026, 4, 10), "End Date": date(2026, 4, 17) },
  { "Swim Lane": "Design", "Task Name": "", "Type": "Task", "Start Date": date(2026, 4, 10), "End Date": date(2026, 4, 17) },
  { "Swim Lane": "Design", "Task Name": "Bad Type", "Type": "Widget", "Start Date": date(2026, 4, 10), "End Date": date(2026, 4, 17) },
  { "Swim Lane": "Design", "Task Name": "No End Date", "Type": "Task", "Start Date": date(2026, 4, 10), "End Date": "" },
  { "Swim Lane": "Design", "Task Name": "End Before Start", "Type": "Task", "Start Date": date(2026, 5, 10), "End Date": date(2026, 4, 10) },
  { "Swim Lane": "Design", "Task Name": "Ghost Dep", "Type": "Task", "Start Date": date(2026, 4, 10), "End Date": date(2026, 4, 17), "Dependency": "Non-Existent Task" },
  { "Swim Lane": "Design", "Task Name": "Cycle A", "Type": "Task", "Start Date": date(2026, 4, 10), "End Date": date(2026, 4, 17), "Dependency": "Cycle B" },
  { "Swim Lane": "Design", "Task Name": "Cycle B", "Type": "Task", "Start Date": date(2026, 4, 10), "End Date": date(2026, 4, 17), "Dependency": "Cycle A" },
];

// ── Write all files ──

const outDir = path.join(__dirname, "fixtures");
const fs = require("fs");
if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

const files = [
  { name: "npi_schedule_10lanes.xlsx", data: npiSchedule },
  { name: "minimal_3lanes.xlsx", data: minimalSchedule },
  { name: "stress_15lanes.xlsx", data: stressSchedule },
  { name: "invalid_data.xlsx", data: invalidSchedule },
];

for (const file of files) {
  const wb = createWorkbook(file.data);
  const outPath = path.join(outDir, file.name);
  XLSX.writeFile(wb, outPath);
  console.log(`Created: ${outPath} (${file.data.length} rows)`);
}

console.log("\nDone. Test fixtures generated.");
