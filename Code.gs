// ============================================================
// Analytics & Performance Dashboard — Google Apps Script
// Version 5.10 — Custom Cycle Start/End Date for Section 4
//
// Built by: Mushfiqur Rahman
// AI Assistance: Claude (Anthropic) — architecture, logic design,
//                and iterative development support
//
// Description:
//   A full-featured analytics dashboard for tracking team
//   performance in large-scale data processing operations.
//   Supports daily, weekly, monthly, and historical reporting
//   with target tracking, holiday exclusions, and trend charts.
// ============================================================

// ── Quality Controller Master List ──────────────────────────
// Replace QC IDs and names with your team's actual data.
// Keywords are used for fuzzy name matching from raw input.
const QC_LIST = [
  { qcId: "QC-001", fullName: "Quality Controller 01",  keywords: ["qc01", "controller01"] },
  { qcId: "QC-002", fullName: "Quality Controller 02",  keywords: ["qc02", "controller02"] },
  { qcId: "QC-003", fullName: "Quality Controller 03",  keywords: ["qc03", "controller03"] },
  { qcId: "QC-004", fullName: "Quality Controller 04",  keywords: ["qc04", "controller04"] },
  { qcId: "QC-005", fullName: "Quality Controller 05",  keywords: ["qc05", "controller05"] },
  { qcId: "QC-006", fullName: "Quality Controller 06",  keywords: ["qc06", "controller06"] },
  { qcId: "QC-007", fullName: "Quality Controller 07",  keywords: ["qc07", "controller07"] },
  { qcId: "QC-008", fullName: "Quality Controller 08",  keywords: ["qc08", "controller08"] },
  { qcId: "QC-009", fullName: "Quality Controller 09",  keywords: ["qc09", "controller09"] },
  { qcId: "QC-010", fullName: "Quality Controller 10",  keywords: ["qc10", "controller10"] },
  { qcId: "QC-011", fullName: "Quality Controller 11",  keywords: ["qc11", "controller11"] },
  { qcId: "QC-012", fullName: "Quality Controller 12",  keywords: ["qc12", "controller12"] },
];

// ── Constants ────────────────────────────────────────────────
const EXPORT_SHEET   = "Export Data";       // Sheet containing raw export data
const STATS_SHEET    = "Statistics";        // Sheet where dashboard is rendered
const EXPORT_HEADER  = 3;                   // Header row number in export sheet
const COL_SERIES     = 7;                   // Column G: Series number
const COL_ENTRY      = 8;                   // Column H: Entry count
const COL_PROCESSED  = 9;                   // Column I: Processed by (QC name)
const COL_DIST_DATE  = 10;                  // Column J: Distribution date
const DATE_CELL      = "B2";               // Date selector cell
const OUTPUT_START   = 5;                   // Row where Section 1 data starts
const FIXED_TEAM     = 6;                   // Fixed team number for output
const FIXED_COMMENTS = "Process";          // Fixed comment label
const FIXED_PROJECT  = "PROJECT";          // Project label (replace with your project name)

// Target table — Col M(13), N(14), O(15)
const TGT_COL_START  = 13;
const TGT_HEADER_ROW = 4;
const TGT_DATA_ROW   = 5;

// Holiday input area — Col Q(17), R(18)
const HOL_COL_START  = 17;
const HOL_HEADER_ROW = 1;
const HOL_DATA_ROW   = 2;
const HOL_MAX_ROWS   = 60;

// Chart helper data — Col T(20), V(22)
const CHART_DATA_COL = 20;

// Cycle input — Col S(19), T(20)
// S2 = Cycle Start, T2 = Cycle End (optional — leave blank for default month)

// ════════════════════════════════════════════════════════════
// MENU
// ════════════════════════════════════════════════════════════
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Analytics Dashboard")
    .addItem("Setup Statistics Sheet",               "setupStatisticsSheet")
    .addSeparator()
    .addItem("Refresh All Sections",                 "refreshAll")
    .addSeparator()
    .addItem("Section 1 — Statistics for Date",      "runStatistics")
    .addItem("Section 2 — Weekly Report",            "runWeeklyReport")
    .addItem("Section 3 — Monthly Report",           "runMonthlyReport")
    .addItem("Section 4 — Target & Progress",        "runTargetProgress")
    .addItem("Section 5 — Series Frequency",         "runSeriesFrequency")
    .addSeparator()
    .addItem("Refresh Date Dropdown",                "refreshDateDropdown")
    .addSeparator()
    .addItem("Refresh Charts",                       "runCharts")
    .addToUi();
}

// ════════════════════════════════════════════════════════════
// CYCLE DATE HELPER — Section 4 custom date range
// S2 = Cycle Start (optional), T2 = Cycle End (optional)
// If blank → default: current month 1st to last day
// ════════════════════════════════════════════════════════════
function getCycleDates() {
  const today = new Date();
  let cycleStart, cycleEnd;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName("Statistics");
    const startVal = sh ? sh.getRange("S2").getValue() : null;
    const endVal   = sh ? sh.getRange("T2").getValue() : null;

    if (startVal instanceof Date && !isNaN(startVal)) {
      cycleStart = new Date(startVal);
      cycleStart.setHours(0, 0, 0, 0);
    } else {
      cycleStart = new Date(today.getFullYear(), today.getMonth(), 1);
    }

    if (endVal instanceof Date && !isNaN(endVal)) {
      cycleEnd = new Date(endVal);
      cycleEnd.setHours(23, 59, 59, 999);
    } else {
      cycleEnd = new Date(today.getFullYear(), today.getMonth() + 1, 0, 23, 59, 59);
    }
  } catch(e) {
    cycleStart = new Date(today.getFullYear(), today.getMonth(), 1);
    cycleEnd   = new Date(today.getFullYear(), today.getMonth() + 1, 0, 23, 59, 59);
  }

  return { cycleStart, cycleEnd };
}

// ════════════════════════════════════════════════════════════
// SHARED HELPERS
// ════════════════════════════════════════════════════════════

// Fuzzy name matching — matches raw input names to QC_LIST entries
// using keyword patterns with word-boundary regex
function lookupQC(processedBy) {
  if (!processedBy) return null;
  const lc = processedBy.toString().toLowerCase().trim();
  const sorted = QC_LIST.slice().sort((a, b) => {
    const aMax = Math.max(...a.keywords.map(k => k.length));
    const bMax = Math.max(...b.keywords.map(k => k.length));
    return bMax - aMax;
  });
  for (let i = 0; i < sorted.length; i++) {
    const entry = sorted[i];
    for (let k = 0; k < entry.keywords.length; k++) {
      const kw  = entry.keywords[k];
      const esc = kw.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      const re  = new RegExp("(?<![a-z])" + esc + "(?![a-z])");
      if (re.test(lc)) return entry;
    }
  }
  return null;
}

function getEntryCount(val) {
  if (!val && val !== 0) return 0;
  if (val instanceof Date) return 0;
  return parseInt(Number(val), 10) || 0;
}

function getQCId(name)     { const e = lookupQC(name); return e ? e.qcId    : null; }
function getFullName(name) { const e = lookupQC(name); return e ? e.fullName : null; }

function loadExportData() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const export_ = ss.getSheetByName(EXPORT_SHEET);
  if (!export_) return [];
  const lastRow = export_.getLastRow();
  if (lastRow <= EXPORT_HEADER) return [];
  return export_.getRange(EXPORT_HEADER + 1, 1, lastRow - EXPORT_HEADER, COL_DIST_DATE).getValues();
}

// Writes a colored section header spanning 8 columns
function writeSectionHeader(sheet, row, title, bgColor) {
  const r = sheet.getRange(row, 1, 1, 8);
  try { r.merge(); } catch(e) {}
  r.setValue(title)
   .setBackground(bgColor || "#1a73e8")
   .setFontColor("#ffffff")
   .setFontWeight("bold")
   .setFontSize(12)
   .setHorizontalAlignment("center");
}

// Writes a colored table header row
function writeTableHeader(sheet, row, headers, bgColor) {
  sheet.getRange(row, 1, 1, headers.length)
       .setValues([headers])
       .setBackground(bgColor || "#34a853")
       .setFontColor("#ffffff")
       .setFontWeight("bold")
       .setHorizontalAlignment("center");
}

// Writes alternating-color data rows with borders
function writeDataRows(sheet, startRow, data, colCount) {
  if (!data || data.length === 0) return;
  const range = sheet.getRange(startRow, 1, data.length, colCount);
  range.setValues(data);
  try {
    if (colCount >= 3) {
      sheet.getRange(startRow, 3, data.length, Math.min(colCount - 2, 6)).setNumberFormat("0");
    }
  } catch(e) {}
  try { sheet.getRange(startRow, 1, data.length, 1).setNumberFormat("@"); } catch(e) {}
  try { sheet.getRange(startRow, 2, data.length, 1).setNumberFormat("@"); } catch(e) {}
  for (let i = 0; i < data.length; i++) {
    try {
      sheet.getRange(startRow + i, 1, 1, colCount)
           .setBackground(i % 2 === 0 ? "#f8f9fa" : "#ffffff")
           .setHorizontalAlignment("center")
           .setBorder(false, false, true, false, false, false,
                      "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID);
      sheet.getRange(startRow + i, 2).setHorizontalAlignment("left");
    } catch(e) {}
  }
}

function writeTotalRow(sheet, row, values, colCount) {
  const r = sheet.getRange(row, 1, 1, colCount);
  r.setValues([values]);
  try { r.setNumberFormat("0"); } catch(e) {}
  r.setBackground("#e8f0fe").setFontWeight("bold").setHorizontalAlignment("center");
}

function writeRefreshNote(sheet, row) {
  sheet.getRange(row, 1, 1, 8).merge()
       .setValue("Last refreshed: " + new Date().toLocaleString())
       .setFontColor("#aaaaaa").setFontStyle("italic")
       .setFontSize(9).setHorizontalAlignment("right")
       .setBackground("#ffffff");
}

function clearFromRow(sheet, fromRow) {
  const lastRow = sheet.getLastRow();
  if (lastRow >= fromRow) {
    const range = sheet.getRange(fromRow, 1, lastRow - fromRow + 1, 11);
    range.clearContent();
    range.clearFormat();
  }
}

// Finds the row where a section header label exists
function findSectionRow(sheet, label) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return 10;
  const colA = sheet.getRange(1, 1, lastRow, 1).getValues();
  for (let i = 0; i < colA.length; i++) {
    if ((colA[i][0] || "").toString().indexOf(label) !== -1) return i + 1;
  }
  return lastRow + 3;
}

// Reads target values from the target input area (Col M/N/O)
function readTargetMap(stats) {
  const map = {};
  try {
    const lastRow = stats.getLastRow();
    if (lastRow < TGT_DATA_ROW) return map;
    const rows = stats.getRange(TGT_DATA_ROW, TGT_COL_START, QC_LIST.length, 3).getValues();
    for (let i = 0; i < rows.length; i++) {
      const qcId   = (rows[i][0] || "").toString().trim();
      const target = Number(rows[i][2]) || 0;
      if (qcId) map[qcId] = target;
    }
  } catch(e) {}
  return map;
}

// ════════════════════════════════════════════════════════════
// REFRESH ALL
// ════════════════════════════════════════════════════════════
function refreshAll() {
  const errors = [];
  try { runStatistics();            } catch(e) { errors.push("S1: " + e.message); }
  try { runWeeklyReport();          } catch(e) { errors.push("S2: " + e.message); }
  try { runMonthlyReport();         } catch(e) { errors.push("S3: " + e.message); }
  try { runTargetProgress();        } catch(e) { errors.push("S4: " + e.message); }
  try { runSeriesFrequency();       } catch(e) { errors.push("S5: " + e.message); }
  try { runCharts();                } catch(e) { errors.push("Charts: " + e.message); }
  if (errors.length > 0) {
    SpreadsheetApp.getActive().toast("Done with issues: " + errors.join(" | "), "Refresh Complete", 8);
  } else {
    SpreadsheetApp.getActive().toast("All sections refreshed!", "Done", 4);
  }
}

// ════════════════════════════════════════════════════════════
// SETUP — Statistics Sheet
// Run this once to initialize the dashboard layout.
// ════════════════════════════════════════════════════════════
function setupStatisticsSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   stats = ss.getSheetByName(STATS_SHEET);
  if (!stats) stats = ss.insertSheet(STATS_SHEET);

  // Preserve existing target values during re-setup
  const savedTargets = {};
  try {
    const lastRow = stats.getLastRow();
    if (lastRow >= TGT_DATA_ROW) {
      const oldRows = stats.getRange(TGT_DATA_ROW, TGT_COL_START, QC_LIST.length, 3).getValues();
      for (let i = 0; i < oldRows.length; i++) {
        const id  = (oldRows[i][0] || "").toString().trim();
        const tgt = Number(oldRows[i][2]) || 0;
        if (id && tgt > 0) savedTargets[id] = tgt;
      }
    }
  } catch(e) {}

  stats.clearContents();
  stats.clearFormats();

  // Dashboard title
  stats.getRange("A1:H1").merge()
       .setValue("Analytics & Performance Dashboard")
       .setFontSize(14).setFontWeight("bold")
       .setBackground("#0d47a1").setFontColor("#ffffff")
       .setHorizontalAlignment("center");

  // Date selector
  stats.getRange("A2").setValue("Select Date:").setFontWeight("bold");
  stats.getRange(DATE_CELL).setBackground("#fff2cc").setFontWeight("bold")
       .setHorizontalAlignment("center");
  stats.getRange("C2")
       .setValue("Type date (e.g. 3/1/2025) or select from dropdown")
       .setFontColor("#888888").setFontStyle("italic");

  // Summary row
  stats.getRange("A3").setValue("Total QC Staff:").setFontWeight("bold");
  stats.getRange("B3").setValue(0).setFontWeight("bold").setFontColor("#1a73e8");
  stats.getRange("D3").setValue("Total Records:").setFontWeight("bold");
  stats.getRange("E3").setValue(0).setFontWeight("bold").setFontColor("#1a73e8");

  // Section 1 header
  writeSectionHeader(stats, 4, "SECTION 1 — Date-wise Entry Statistics", "#1565c0");
  const sec1Headers = ["QC ID", "Date", "Team", "Comments", "Project", "Records"];
  stats.getRange(OUTPUT_START, 1, 1, 6)
       .setValues([sec1Headers])
       .setBackground("#34a853").setFontColor("#ffffff")
       .setFontWeight("bold").setHorizontalAlignment("center");

  // Column widths
  stats.setColumnWidth(1, 150);
  stats.setColumnWidth(2, 130);
  stats.setColumnWidth(3, 60);
  stats.setColumnWidth(4, 100);
  stats.setColumnWidth(5, 90);
  stats.setColumnWidth(6, 110);
  stats.setColumnWidth(7, 120);
  stats.setColumnWidth(8, 120);
  stats.setColumnWidth(9,  25);
  stats.setColumnWidth(10, 25);
  stats.setColumnWidth(11, 25);

  // Column L: visual divider between analytics and target areas
  stats.setColumnWidth(12, 15);
  stats.getRange(1, 12, stats.getMaxRows(), 1)
       .setBackground("#b0bec5")
       .setBorder(false, true, false, true, false, false, "#90a4ae", SpreadsheetApp.BorderStyle.SOLID);

  setupTargetInputArea(stats, savedTargets);
  setupHolidayInputArea(stats);

  // Cycle date input (Section 4 custom range)
  stats.getRange("S1").setValue("Cycle Start (opt)").setFontWeight("bold").setBackground("#CFE2FF").setHorizontalAlignment("center");
  stats.getRange("T1").setValue("Cycle End (opt)").setFontWeight("bold").setBackground("#CFE2FF").setHorizontalAlignment("center");
  stats.getRange("S2").setBackground("#FFF2CC").setNumberFormat("D/M/yyyy").setHorizontalAlignment("center");
  stats.getRange("T2").setBackground("#FFF2CC").setNumberFormat("D/M/yyyy").setHorizontalAlignment("center");
  stats.getRange("S3").setValue("← Section 4 only").setFontColor("#888888").setFontStyle("italic").setFontSize(8);
  stats.getRange("T3").setValue("Leave blank = default month").setFontColor("#888888").setFontStyle("italic").setFontSize(8);
  stats.setColumnWidth(19, 130);
  stats.setColumnWidth(20, 130);

  stats.setFrozenRows(OUTPUT_START);
  refreshDateDropdown();

  SpreadsheetApp.getUi().alert(
    "Setup complete!\n\n" +
    "LEFT   : Analytics — Col A to H\n" +
    "DIVIDER: Col L (gray)\n" +
    "RIGHT  : Monthly Targets — Col M | N | O\n" +
    "         Public Holidays — Col Q | R\n" +
    "         Cycle Dates (Section 4) — Col S | T\n\n" +
    "Set Target on Column O then press Refresh All."
  );
}

// Sets up the monthly target input area (right side of sheet)
function setupTargetInputArea(stats, savedTargets) {
  savedTargets = savedTargets || {};
  stats.getRange(1, TGT_COL_START, 1, 3).merge()
       .setValue("MONTHLY TARGETS")
       .setBackground("#37474f").setFontColor("#ffffff")
       .setFontWeight("bold").setFontSize(12)
       .setHorizontalAlignment("center");
  stats.getRange(2, TGT_COL_START, 1, 3).merge()
       .setValue("Edit column O (Monthly Target) only")
       .setFontColor("#888888").setFontStyle("italic")
       .setFontSize(9).setHorizontalAlignment("center");
  stats.getRange(3, TGT_COL_START, 1, 3).merge()
       .setValue("Columns M & N are auto-updated on refresh")
       .setFontColor("#aaaaaa").setFontStyle("italic")
       .setFontSize(9).setHorizontalAlignment("center");
  stats.getRange(TGT_HEADER_ROW, TGT_COL_START, 1, 3)
       .setValues([["QC ID", "Full Name", "Monthly Target"]])
       .setBackground("#546e7a").setFontColor("#ffffff")
       .setFontWeight("bold").setHorizontalAlignment("center");
  const rows = QC_LIST.map(e => [e.qcId, e.fullName, savedTargets[e.qcId] || 0]);
  stats.getRange(TGT_DATA_ROW, TGT_COL_START, rows.length, 3).setValues(rows);
  stats.getRange(TGT_DATA_ROW, TGT_COL_START + 2, rows.length, 1)
       .setBackground("#fffde7");
  stats.setColumnWidth(TGT_COL_START,     120);
  stats.setColumnWidth(TGT_COL_START + 1, 200);
  stats.setColumnWidth(TGT_COL_START + 2, 150);
}

// ════════════════════════════════════════════════════════════
// HOLIDAY INPUT AREA SETUP
// Fridays are auto-excluded. Add public holidays here.
// ════════════════════════════════════════════════════════════
function setupHolidayInputArea(stats) {
  const col = HOL_COL_START;

  // Spacer column
  stats.setColumnWidth(16, 20);
  stats.getRange(1, 16, stats.getMaxRows(), 1).setBackground("#e0e0e0");

  stats.getRange(HOL_HEADER_ROW, col, 1, 2).merge()
       .setValue("PUBLIC HOLIDAYS")
       .setBackground("#263238").setFontColor("#ffffff")
       .setFontWeight("bold").setFontSize(11)
       .setHorizontalAlignment("center");

  stats.getRange(HOL_HEADER_ROW + 1, col, 1, 2).merge()
       .setValue("Friday auto-excluded. Add public holidays below.")
       .setFontColor("#888888").setFontStyle("italic")
       .setFontSize(9).setHorizontalAlignment("center");

  stats.getRange(HOL_HEADER_ROW + 2, col, 1, 2)
       .setValues([["Date (DD/MM/YYYY)", "Description"]])
       .setBackground("#37474f").setFontColor("#ffffff")
       .setFontWeight("bold").setHorizontalAlignment("center");

  // Sample public holidays — replace with your country's holidays
  const sampleHolidays = [
    ["01/01/2026", "New Year's Day"],
    ["01/05/2026", "Labour Day"],
    ["25/12/2026", "Christmas Day"],
  ];

  stats.getRange(HOL_DATA_ROW, col, sampleHolidays.length, 2).setValues(sampleHolidays);
  stats.getRange(HOL_DATA_ROW, col, HOL_MAX_ROWS, 1).setBackground("#fce4ec");
  stats.getRange(HOL_DATA_ROW, col, HOL_MAX_ROWS, 1).setNumberFormat("dd/mm/yyyy");
  stats.setColumnWidth(col,     140);
  stats.setColumnWidth(col + 1, 220);

  stats.getRange(HOL_DATA_ROW + HOL_MAX_ROWS, col, 1, 2).merge()
       .setValue("↑ Add more holidays above as needed")
       .setFontColor("#1565c0").setFontStyle("italic").setFontSize(9);
}

// Reads holiday dates into a Set for working-day calculations
function readHolidaySet(stats, tz) {
  const set = {};
  try {
    const rows = stats.getRange(HOL_DATA_ROW, HOL_COL_START, HOL_MAX_ROWS, 1).getValues();
    for (let i = 0; i < rows.length; i++) {
      const v = rows[i][0];
      if (!v) continue;
      let d;
      if (v instanceof Date) {
        d = v;
      } else {
        const parts = v.toString().trim().split("/");
        if (parts.length === 3) {
          d = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
        }
      }
      if (d && !isNaN(d)) {
        set[Utilities.formatDate(d, tz, "yyyy-MM-dd")] = true;
      }
    }
  } catch(e) {}
  return set;
}

// Counts working days between two dates, excluding Fridays and holidays
function workingDaysBetween(dateStr1, dateStr2, holidaySet) {
  const d1 = new Date(dateStr1);
  const d2 = new Date(dateStr2);
  if (d2 <= d1) return 0;
  let count = 0;
  const cur = new Date(d1);
  cur.setDate(cur.getDate() + 1);
  while (cur <= d2) {
    const dow = cur.getDay();
    const key = cur.getFullYear() + "-" +
                String(cur.getMonth() + 1).padStart(2, "0") + "-" +
                String(cur.getDate()).padStart(2, "0");
    if (dow !== 5 && !holidaySet[key]) count++;
    cur.setDate(cur.getDate() + 1);
  }
  return count;
}

// ════════════════════════════════════════════════════════════
// DATE DROPDOWN — Builds dropdown from available dates in export
// ════════════════════════════════════════════════════════════
function refreshDateDropdown() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const export_ = ss.getSheetByName(EXPORT_SHEET);
  const stats   = ss.getSheetByName(STATS_SHEET);
  if (!export_ || !stats) return;
  const lastRow = export_.getLastRow();
  if (lastRow <= EXPORT_HEADER) return;
  const dateVals = export_.getRange(EXPORT_HEADER + 1, COL_DIST_DATE, lastRow - EXPORT_HEADER, 1).getValues();
  const tz = Session.getScriptTimeZone();
  const seen = {}, list = [];
  for (let i = 0; i < dateVals.length; i++) {
    const v = dateVals[i][0];
    if (v instanceof Date && !isNaN(v)) {
      const key = Utilities.formatDate(v, tz, "D/M/yyyy");
      if (!seen[key]) { seen[key] = true; list.push(key); }
    }
  }
  list.sort((a, b) => new Date(a) - new Date(b));
  if (!list.length) return;
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(list, true).setAllowInvalid(true)
    .setHelpText("Select a date or type manually (M/D/YYYY)").build();
  stats.getRange(DATE_CELL).setDataValidation(rule);
}

// ════════════════════════════════════════════════════════════
// SECTION 1 — Date-wise Entry Statistics
// Shows records processed by each QC on a selected date.
// ════════════════════════════════════════════════════════════
function runStatistics() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const export_ = ss.getSheetByName(EXPORT_SHEET);
  const stats   = ss.getSheetByName(STATS_SHEET);
  if (!export_) { SpreadsheetApp.getUi().alert("'" + EXPORT_SHEET + "' sheet not found!"); return; }
  if (!stats)   { SpreadsheetApp.getUi().alert("'" + STATS_SHEET + "' sheet not found! Run Setup first."); return; }

  const dateInput = stats.getRange(DATE_CELL).getValue();
  if (!dateInput) { SpreadsheetApp.getUi().alert("Please enter a date in cell B2."); return; }
  const targetDate = (dateInput instanceof Date) ? dateInput : new Date(dateInput);
  if (isNaN(targetDate)) { SpreadsheetApp.getUi().alert("Invalid date format. Example: 3/1/2025"); return; }

  const tz          = Session.getScriptTimeZone();
  const targetStr   = Utilities.formatDate(targetDate, tz, "yyyy-MM-dd");
  const displayDate = Utilities.formatDate(targetDate, tz, "d-MMM-yyyy");
  const lastRow     = export_.getLastRow();
  if (lastRow <= EXPORT_HEADER) { SpreadsheetApp.getUi().alert("No data found in '" + EXPORT_SHEET + "'."); return; }

  const data  = export_.getRange(EXPORT_HEADER + 1, 1, lastRow - EXPORT_HEADER, COL_DIST_DATE).getValues();
  const opMap = {};

  for (let i = 0; i < data.length; i++) {
    const row      = data[i];
    const distDate = row[COL_DIST_DATE - 1];
    if (!(distDate instanceof Date) || isNaN(distDate)) continue;
    if (Utilities.formatDate(distDate, tz, "yyyy-MM-dd") !== targetStr) continue;
    const rawName    = (row[COL_PROCESSED - 1] || "").toString().trim();
    const entryCount = getEntryCount(row[COL_ENTRY - 1]);
    if (!rawName) continue;
    const qcEntry = lookupQC(rawName);
    const key     = qcEntry ? qcEntry.qcId : rawName.toLowerCase();
    if (opMap[key]) {
      opMap[key].totalEntry += entryCount;
    } else {
      opMap[key] = {
        qcId:       qcEntry ? qcEntry.qcId    : "UNKNOWN",
        fullName:   qcEntry ? qcEntry.fullName : rawName,
        totalEntry: entryCount
      };
    }
  }

  stats.getRange(OUTPUT_START + 1, 1, 200, 6).clearContent().clearFormat();
  const processors = Object.values(opMap);

  if (processors.length === 0) {
    stats.getRange(OUTPUT_START + 1, 1).setValue("No data found for this date.")
         .setFontColor("#cc0000").setFontStyle("italic");
    stats.getRange("B3").setValue(0);
    stats.getRange("E3").setValue(0);
    return;
  }

  processors.sort((a, b) => a.qcId.localeCompare(b.qcId));
  const outputData = [];
  let   totalRecord = 0;

  for (let i = 0; i < processors.length; i++) {
    const op = processors[i];
    totalRecord += op.totalEntry;
    outputData.push([op.qcId, displayDate, FIXED_TEAM, FIXED_COMMENTS, FIXED_PROJECT, op.totalEntry]);
  }

  const outRange = stats.getRange(OUTPUT_START + 1, 1, outputData.length, 6);
  outRange.setValues(outputData);
  try { stats.getRange(OUTPUT_START + 1, 6, outputData.length, 1).setNumberFormat("0"); } catch(e) {}

  for (let i = 0; i < outputData.length; i++) {
    stats.getRange(OUTPUT_START + 1 + i, 1, 1, 6)
         .setBackground(i % 2 === 0 ? "#f8f9fa" : "#ffffff")
         .setHorizontalAlignment("center")
         .setBorder(false, false, true, false, false, false,
                    "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID);
  }

  stats.getRange(OUTPUT_START + 1, 6, outputData.length, 1).setFontWeight("bold").setFontColor("#1a73e8");
  stats.getRange(OUTPUT_START + 1, 1, outputData.length, 1).setFontWeight("bold");

  const totalRow = OUTPUT_START + 1 + outputData.length;
  const totR = stats.getRange(totalRow, 1, 1, 6);
  totR.setValues([["TOTAL", displayDate, "", "", "", totalRecord]])
      .setBackground("#e8f0fe").setFontWeight("bold")
      .setHorizontalAlignment("center");
  try { stats.getRange(totalRow, 6, 1, 1).setNumberFormat("0"); } catch(e) {}

  stats.getRange("B3").setValue(processors.length);
  stats.getRange("E3").setValue(totalRecord);
  SpreadsheetApp.getActive().toast(
    processors.length + " QC staff found | Total Records: " + totalRecord,
    "Section 1 Updated", 5
  );
}

// ════════════════════════════════════════════════════════════
// SECTION 2 — Weekly Performance
// Current week: Saturday to Friday (adjust dow logic if needed)
// ════════════════════════════════════════════════════════════
function runWeeklyReport() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const stats = ss.getSheetByName(STATS_SHEET);
  if (!stats) { SpreadsheetApp.getUi().alert("Run Setup first."); return; }

  const tz    = Session.getScriptTimeZone();
  const today = new Date();
  const data  = loadExportData();

  const dow     = today.getDay();
  const diffSat = -((dow + 1) % 7);
  const wStart  = new Date(today); wStart.setDate(today.getDate() + diffSat); wStart.setHours(0,0,0,0);
  const wEnd    = new Date(wStart); wEnd.setDate(wStart.getDate() + 6); wEnd.setHours(23,59,59,999);
  const wStartStr = Utilities.formatDate(wStart, tz, "yyyy-MM-dd");
  const wEndStr   = Utilities.formatDate(wEnd,   tz, "yyyy-MM-dd");

  const opMap = {};
  for (let i = 0; i < data.length; i++) {
    const row      = data[i];
    const distDate = row[COL_DIST_DATE - 1];
    if (!(distDate instanceof Date) || isNaN(distDate)) continue;
    const rowDate = Utilities.formatDate(distDate, tz, "yyyy-MM-dd");
    if (rowDate < wStartStr || rowDate > wEndStr) continue;
    const rawName    = (row[COL_PROCESSED - 1] || "").toString().trim();
    const entryCount = getEntryCount(row[COL_ENTRY - 1]);
    const seriesNum  = row[COL_SERIES - 1];
    if (!rawName) continue;
    const qcId = getQCId(rawName);
    if (!qcId) continue;
    const fullName = getFullName(rawName);
    if (!opMap[qcId]) {
      opMap[qcId] = { qcId, fullName, totalEntry: 0, seriesSet: new Set(), activeDays: new Set() };
    }
    opMap[qcId].totalEntry += entryCount;
    opMap[qcId].seriesSet.add(String(seriesNum));
    opMap[qcId].activeDays.add(rowDate);
  }

  const sec2Row = findSectionRow(stats, "SECTION 2");
  clearFromRow(stats, sec2Row);

  const weekLabel = Utilities.formatDate(wStart, tz, "d-MMM") + " to " +
                    Utilities.formatDate(wEnd, tz, "d-MMM-yyyy");
  writeSectionHeader(stats, sec2Row,
    "SECTION 2 — Weekly Performance  (" + weekLabel + ")", "#4a148c");
  writeTableHeader(stats, sec2Row + 1,
    ["QC ID", "Full Name", "Active Days", "Total Series", "Total Records", "Daily Avg Records", "Daily Avg Series", ""],
    "#6a1b9a");

  const ops = Object.values(opMap).sort((a, b) => b.totalEntry - a.totalEntry);
  const outData = [];
  let grandEntry = 0, grandSeries = 0;

  for (let i = 0; i < ops.length; i++) {
    const op   = ops[i];
    const days = op.activeDays.size || 1;
    const ser  = op.seriesSet.size;
    grandEntry  += op.totalEntry;
    grandSeries += ser;
    outData.push([op.qcId, op.fullName, days, ser, op.totalEntry,
                  Math.round(op.totalEntry / days), (ser / days).toFixed(1), ""]);
  }

  writeDataRows(stats, sec2Row + 2, outData, 8);
  if (ops.length > 0) stats.getRange(sec2Row + 2, 1, 1, 8).setBackground("#e8f5e9").setFontWeight("bold");
  stats.getRange(sec2Row + 2, 5, Math.max(outData.length,1), 1).setFontWeight("bold").setFontColor("#1a73e8");

  const totalRow = sec2Row + 2 + outData.length;
  writeTotalRow(stats, totalRow, ["TOTAL", "", "", grandSeries, grandEntry, "", "", ""], 8);
  writeRefreshNote(stats, totalRow + 1);
  SpreadsheetApp.getActive().toast("Weekly report updated!", "Section 2 Updated", 3);
}

// ════════════════════════════════════════════════════════════
// SECTION 3 — Monthly Performance
// Aggregates all records for the current calendar month.
// ════════════════════════════════════════════════════════════
function runMonthlyReport() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const stats = ss.getSheetByName(STATS_SHEET);
  if (!stats) return;

  const tz    = Session.getScriptTimeZone();
  const today = new Date();
  const data  = loadExportData();

  const mStart    = new Date(today.getFullYear(), today.getMonth(), 1);
  const mEnd      = new Date(today.getFullYear(), today.getMonth() + 1, 0, 23, 59, 59);
  const mStartStr = Utilities.formatDate(mStart, tz, "yyyy-MM-dd");
  const mEndStr   = Utilities.formatDate(mEnd,   tz, "yyyy-MM-dd");

  const opMap = {};
  for (let i = 0; i < data.length; i++) {
    const row      = data[i];
    const distDate = row[COL_DIST_DATE - 1];
    if (!(distDate instanceof Date) || isNaN(distDate)) continue;
    const rowDate = Utilities.formatDate(distDate, tz, "yyyy-MM-dd");
    if (rowDate < mStartStr || rowDate > mEndStr) continue;
    const rawName    = (row[COL_PROCESSED - 1] || "").toString().trim();
    const entryCount = getEntryCount(row[COL_ENTRY - 1]);
    const seriesNum  = row[COL_SERIES - 1];
    if (!rawName) continue;
    const qcId = getQCId(rawName);
    if (!qcId) continue;
    const fullName = getFullName(rawName);
    if (!opMap[qcId]) {
      opMap[qcId] = { qcId, fullName, totalEntry: 0, seriesSet: new Set(), activeDays: new Set() };
    }
    opMap[qcId].totalEntry += entryCount;
    opMap[qcId].seriesSet.add(String(seriesNum));
    opMap[qcId].activeDays.add(rowDate);
  }

  const sec3Row = findSectionRow(stats, "SECTION 3");
  clearFromRow(stats, sec3Row);

  const monthLabel = Utilities.formatDate(mStart, tz, "MMMM yyyy");
  writeSectionHeader(stats, sec3Row,
    "SECTION 3 — Monthly Performance  (" + monthLabel + ")", "#bf360c");
  writeTableHeader(stats, sec3Row + 1,
    ["QC ID", "Full Name", "Active Days", "Total Series", "Total Records", "Daily Avg Records", "Daily Avg Series", ""],
    "#e64a19");

  const ops = Object.values(opMap).sort((a, b) => b.totalEntry - a.totalEntry);
  const outData = [];
  let grandEntry = 0, grandSeries = 0;

  for (let i = 0; i < ops.length; i++) {
    const op   = ops[i];
    const days = op.activeDays.size || 1;
    const ser  = op.seriesSet.size;
    grandEntry  += op.totalEntry;
    grandSeries += ser;
    outData.push([op.qcId, op.fullName, days, ser, op.totalEntry,
                  Math.round(op.totalEntry / days), (ser / days).toFixed(1), ""]);
  }

  writeDataRows(stats, sec3Row + 2, outData, 8);
  if (ops.length > 0) stats.getRange(sec3Row + 2, 1, 1, 8).setBackground("#fbe9e7").setFontWeight("bold");
  stats.getRange(sec3Row + 2, 5, Math.max(outData.length,1), 1).setFontWeight("bold").setFontColor("#1a73e8");

  const totalRow = sec3Row + 2 + outData.length;
  writeTotalRow(stats, totalRow, ["TOTAL", "", "", grandSeries, grandEntry, "", "", ""], 8);
  writeRefreshNote(stats, totalRow + 1);
  SpreadsheetApp.getActive().toast("Monthly report updated!", "Section 3 Updated", 3);
}

// ════════════════════════════════════════════════════════════
// SECTION 4 — Target & Progress
// Compares each QC's achievement against their monthly target.
// Supports custom cycle dates (S2 = start, T2 = end).
// Status logic: Achieved / On Track / At Risk / Not Started
// ════════════════════════════════════════════════════════════
function runTargetProgress() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const stats = ss.getSheetByName(STATS_SHEET);
  if (!stats) return;

  const tz    = Session.getScriptTimeZone();
  const today = new Date();
  const data  = loadExportData();

  const { cycleStart: mStart, cycleEnd: mEnd } = getCycleDates();
  const mStartStr = Utilities.formatDate(mStart, tz, "yyyy-MM-dd");
  const mEndStr   = Utilities.formatDate(mEnd,   tz, "yyyy-MM-dd");

  const daysInMonth = Math.round((mEnd - mStart) / (1000 * 60 * 60 * 24)) + 1;
  const daysElapsed = Math.min(
    Math.round((today - mStart) / (1000 * 60 * 60 * 24)) + 1,
    daysInMonth
  );
  const daysLeft   = Math.max(0, daysInMonth - daysElapsed);
  const daysPassed = daysElapsed;

  const dow       = today.getDay();
  const diffSat   = -((dow + 1) % 7);
  const wStart    = new Date(today); wStart.setDate(today.getDate() + diffSat); wStart.setHours(0,0,0,0);
  const wEnd      = new Date(wStart); wEnd.setDate(wStart.getDate() + 6); wEnd.setHours(23,59,59,999);
  const wStartStr = Utilities.formatDate(wStart, tz, "yyyy-MM-dd");
  const wEndStr   = Utilities.formatDate(wEnd,   tz, "yyyy-MM-dd");

  const targetMap = readTargetMap(stats);
  const opMap = {};

  for (let i = 0; i < data.length; i++) {
    const row      = data[i];
    const distDate = row[COL_DIST_DATE - 1];
    if (!(distDate instanceof Date) || isNaN(distDate)) continue;
    const rowDate    = Utilities.formatDate(distDate, tz, "yyyy-MM-dd");
    const rawName    = (row[COL_PROCESSED - 1] || "").toString().trim();
    const entryCount = getEntryCount(row[COL_ENTRY - 1]);
    const seriesNum  = row[COL_SERIES - 1];
    if (!rawName) continue;
    const qcId = getQCId(rawName);
    if (!qcId) continue;

    if (!opMap[qcId]) {
      opMap[qcId] = {
        qcId,
        fullName:    getFullName(rawName),
        cycleEntry:  0,
        weeklyEntry: 0,
        seriesSet:   new Set()
      };
    }
    if (rowDate >= mStartStr && rowDate <= mEndStr) {
      opMap[qcId].cycleEntry += entryCount;
      opMap[qcId].seriesSet.add(String(seriesNum));
    }
    if (rowDate >= wStartStr && rowDate <= wEndStr) {
      opMap[qcId].weeklyEntry += entryCount;
    }
  }

  const sec4Row = findSectionRow(stats, "SECTION 4");
  clearFromRow(stats, sec4Row);

  const cycleLabel = Utilities.formatDate(mStart, tz, "d-MMM-yyyy") + " → " +
                     Utilities.formatDate(mEnd,   tz, "d-MMM-yyyy");

  stats.getRange(sec4Row, 1, 1, 9).merge()
       .setValue("SECTION 4 — Target & Progress  (" + cycleLabel + "  |  Day " + daysElapsed + "/" + daysInMonth + ")")
       .setBackground("#880e4f").setFontColor("#ffffff")
       .setFontWeight("bold").setFontSize(12)
       .setHorizontalAlignment("center");

  stats.getRange(sec4Row + 1, 1, 1, 9)
       .setValues([["QC ID", "Full Name", "Cycle Target", "This Cycle", "This Week", "% Achieved", "Remaining", "Avg Rec/Series", "Status"]])
       .setBackground("#ad1457").setFontColor("#ffffff")
       .setFontWeight("bold").setHorizontalAlignment("center");

  const outData = [];
  for (let i = 0; i < QC_LIST.length; i++) {
    const qc     = QC_LIST[i];
    const op     = opMap[qc.qcId] || { cycleEntry: 0, weeklyEntry: 0, seriesSet: new Set() };
    const target = targetMap[qc.qcId] || 0;
    if (target === 0 && op.cycleEntry === 0) continue;

    const achieved     = op.cycleEntry;
    const weekEntry    = op.weeklyEntry;
    const totalSeries  = op.seriesSet.size;
    const avgRecSeries = totalSeries > 0 ? Math.round(achieved / totalSeries) : 0;

    const pct       = target > 0 ? ((achieved / target) * 100).toFixed(1) + "%" : "N/A";
    const remaining = target > 0 ? Math.max(0, target - achieved) : 0;
    const reqPerDay = daysLeft > 0 ? Math.round(remaining / daysLeft) : remaining;
    const status    = target === 0                        ? "No Target"    :
                      achieved >= target                  ? "Achieved ✓"   :
                      (achieved === 0 && daysPassed > 3) ? "Not Started !" :
                      reqPerDay > Math.round(target / daysInMonth) * 1.09 ? "At Risk !" :
                      "On Track";

    outData.push([qc.qcId, qc.fullName, target, achieved, weekEntry, pct, remaining, avgRecSeries, status]);
  }

  // Sort: Achieved first, then On Track, At Risk, Not Started, No Target
  outData.sort((a, b) => {
    const order = { "Achieved ✓": 0, "On Track": 1, "At Risk !": 2, "Not Started !": 3, "No Target": 4 };
    return (order[a[8]] || 9) - (order[b[8]] || 9) || b[3] - a[3];
  });

  if (outData.length > 0) {
    stats.getRange(sec4Row + 2, 1, outData.length, 9).setValues(outData);
    try { stats.getRange(sec4Row + 2, 3, outData.length, 2).setNumberFormat("0"); } catch(e) {}
    try { stats.getRange(sec4Row + 2, 5, outData.length, 1).setNumberFormat("0"); } catch(e) {}
    try { stats.getRange(sec4Row + 2, 6, outData.length, 1).setNumberFormat("@"); } catch(e) {}
    try { stats.getRange(sec4Row + 2, 7, outData.length, 2).setNumberFormat("0"); } catch(e) {}
  }

  for (let i = 0; i < outData.length; i++) {
    stats.getRange(sec4Row + 2 + i, 1, 1, 9)
         .setBackground(i % 2 === 0 ? "#f8f9fa" : "#ffffff")
         .setHorizontalAlignment("center")
         .setBorder(false, false, true, false, false, false,
                    "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID);
    stats.getRange(sec4Row + 2 + i, 2).setHorizontalAlignment("left");

    const statusCell = stats.getRange(sec4Row + 2 + i, 9);
    const pctCell    = stats.getRange(sec4Row + 2 + i, 6);
    const s = outData[i][8];

    if (s === "Achieved ✓")        { statusCell.setBackground("#c8e6c9").setFontColor("#1b5e20").setFontWeight("bold"); }
    else if (s === "At Risk !")     { statusCell.setBackground("#ffcdd2").setFontColor("#b71c1c").setFontWeight("bold"); }
    else if (s === "Not Started !") { statusCell.setBackground("#f8bbd0").setFontColor("#880e4f").setFontWeight("bold"); }
    else if (s === "On Track")      { statusCell.setBackground("#fff9c4").setFontColor("#f57f17"); }

    const pctVal = parseFloat(outData[i][5]);
    if (!isNaN(pctVal)) {
      if (pctVal >= 100)    pctCell.setFontColor("#1b5e20").setFontWeight("bold");
      else if (pctVal < 50) pctCell.setFontColor("#b71c1c");
    }
  }

  const totalRow  = sec4Row + 2 + outData.length;
  const gTarget   = outData.reduce((s, r) => s + (Number(r[2]) || 0), 0);
  const gAchieved = outData.reduce((s, r) => s + (Number(r[3]) || 0), 0);
  const gWeekly   = outData.reduce((s, r) => s + (Number(r[4]) || 0), 0);
  const gRemain   = outData.reduce((s, r) => s + (Number(r[6]) || 0), 0);

  const totalRange = stats.getRange(totalRow, 1, 1, 9);
  totalRange.setValues([["TOTAL", "", gTarget, gAchieved, gWeekly,
                          gTarget > 0 ? ((gAchieved / gTarget) * 100).toFixed(1) + "%" : "N/A",
                          gRemain, "", ""]]);
  try { stats.getRange(totalRow, 3, 1, 2).setNumberFormat("0"); } catch(e) {}
  try { stats.getRange(totalRow, 5, 1, 1).setNumberFormat("0"); } catch(e) {}
  try { stats.getRange(totalRow, 6, 1, 1).setNumberFormat("@"); } catch(e) {}
  try { stats.getRange(totalRow, 7, 1, 1).setNumberFormat("0"); } catch(e) {}
  totalRange.setBackground("#e8f0fe").setFontWeight("bold").setHorizontalAlignment("center");

  writeRefreshNote(stats, totalRow + 1);
  SpreadsheetApp.getActive().toast("Target & Progress updated!", "Section 4 Updated", 3);
}

// ════════════════════════════════════════════════════════════
// SECTION 5 — Series Frequency & Verdict
// Measures how consistently each QC submits work.
// Verdict: Active (≤3 days gap) / Moderate (≤5) / Inconsistent (>5)
// ════════════════════════════════════════════════════════════
function getVerdict(avgDays) {
  if (avgDays === null) return "N/A";
  if (avgDays <= 3)     return "Active";
  if (avgDays <= 5)     return "Moderate";
  return "Inconsistent";
}

// Calculates average working-day gap between submission dates
function calcAvgGap(uniqueDates, holidaySet) {
  holidaySet = holidaySet || {};
  if (uniqueDates.length < 2) return null;
  let total = 0;
  for (let d = 1; d < uniqueDates.length; d++) {
    total += workingDaysBetween(uniqueDates[d - 1], uniqueDates[d], holidaySet);
  }
  return total / (uniqueDates.length - 1);
}

function runSeriesFrequency() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const stats = ss.getSheetByName(STATS_SHEET);
  if (!stats) return;

  const tz    = Session.getScriptTimeZone();
  const today = new Date();
  const data  = loadExportData();

  const mStart    = new Date(today.getFullYear(), today.getMonth(), 1);
  const mEnd      = new Date(today.getFullYear(), today.getMonth() + 1, 0, 23, 59, 59);
  const mStartStr = Utilities.formatDate(mStart, tz, "yyyy-MM-dd");
  const mEndStr   = Utilities.formatDate(mEnd,   tz, "yyyy-MM-dd");
  const monthLabel = Utilities.formatDate(mStart, tz, "MMMM yyyy");

  const yStart    = new Date(today.getFullYear(), 0, 1);
  const yStartStr = Utilities.formatDate(yStart, tz, "yyyy-MM-dd");
  const todayStr  = Utilities.formatDate(today,  tz, "yyyy-MM-dd");
  const yearLabel = today.getFullYear().toString();

  const opMap = {};
  for (let i = 0; i < data.length; i++) {
    const row      = data[i];
    const distDate = row[COL_DIST_DATE - 1];
    if (!(distDate instanceof Date) || isNaN(distDate)) continue;
    const rowDate = Utilities.formatDate(distDate, tz, "yyyy-MM-dd");
    const rawName = (row[COL_PROCESSED - 1] || "").toString().trim();
    if (!rawName) continue;
    const qcId = getQCId(rawName);
    if (!qcId) continue;
    if (!opMap[qcId]) {
      opMap[qcId] = {
        qcId, fullName: getFullName(rawName),
        mDates: [], mSeries: 0,
        yDates: [], ySeries: 0, yTotalRecords: 0
      };
    }
    if (rowDate >= mStartStr && rowDate <= mEndStr) {
      opMap[qcId].mDates.push(rowDate);
      opMap[qcId].mSeries++;
    }
    if (rowDate >= yStartStr && rowDate <= todayStr) {
      opMap[qcId].yDates.push(rowDate);
      opMap[qcId].ySeries++;
      opMap[qcId].yTotalRecords += getEntryCount(row[COL_ENTRY - 1]);
    }
  }

  const holidaySet = readHolidaySet(stats, tz);
  const sec5Row = findSectionRow(stats, "SECTION 5");
  clearFromRow(stats, sec5Row);

  writeSectionHeader(stats, sec5Row,
    "SECTION 5 — Series Frequency & Verdict  (" + monthLabel + "  |  Yearly: " + yearLabel + ")  [Working Days — Weekly off & Holidays excluded]",
    "#004d40");
  writeTableHeader(stats, sec5Row + 1,
    ["QC ID", "Full Name", "Month Series", "Month Avg Days", "Monthly Verdict",
     "Year Series", "Year Avg Days", "Yearly Verdict", "Year Total Records"],
    "#00695c");

  const ops = Object.values(opMap).sort((a, b) => a.qcId.localeCompare(b.qcId));
  const outData = [];

  for (let i = 0; i < ops.length; i++) {
    const op = ops[i];
    const mUnique  = [...new Set(op.mDates)].sort();
    const mAvgRaw  = calcAvgGap(mUnique, holidaySet);
    const mAvgStr  = mAvgRaw !== null ? mAvgRaw.toFixed(1) + " days" : (mUnique.length === 1 ? "1 date" : "N/A");
    const mVerdict = mAvgRaw !== null ? getVerdict(mAvgRaw) : (mUnique.length === 1 ? getVerdict(0) : "N/A");
    const yUnique  = [...new Set(op.yDates)].sort();
    const yAvgRaw  = calcAvgGap(yUnique, holidaySet);
    const yAvgStr  = yAvgRaw !== null ? yAvgRaw.toFixed(1) + " days" : (yUnique.length === 1 ? "1 date" : "N/A");
    const yVerdict = yAvgRaw !== null ? getVerdict(yAvgRaw) : (yUnique.length === 1 ? getVerdict(0) : "N/A");
    if (op.mSeries === 0 && op.ySeries === 0) continue;
    outData.push([
      op.qcId, op.fullName,
      op.mSeries, mAvgStr, mVerdict,
      op.ySeries, yAvgStr, yVerdict,
      op.yTotalRecords
    ]);
  }

  writeDataRows(stats, sec5Row + 2, outData, 9);

  if (outData.length > 0) {
    try { stats.getRange(sec5Row + 2, 3, outData.length, 1).setNumberFormat("0"); } catch(e) {}
    try { stats.getRange(sec5Row + 2, 6, outData.length, 1).setNumberFormat("0"); } catch(e) {}
    try { stats.getRange(sec5Row + 2, 4, outData.length, 2).setNumberFormat("@"); } catch(e) {}
    try { stats.getRange(sec5Row + 2, 7, outData.length, 2).setNumberFormat("@"); } catch(e) {}
    try { stats.getRange(sec5Row + 2, 9, outData.length, 1).setNumberFormat("#,##0"); } catch(e) {}
    try { stats.getRange(sec5Row + 2, 9, outData.length, 1).setFontWeight("bold").setFontColor("#1565c0"); } catch(e) {}

    for (let i = 0; i < outData.length; i++) {
      const mV = outData[i][4];
      const yV = outData[i][7];
      const mCell = stats.getRange(sec5Row + 2 + i, 5);
      const yCell = stats.getRange(sec5Row + 2 + i, 8);
      try {
        if (mV === "Active")            { mCell.setBackground("#c8e6c9").setFontColor("#1b5e20").setFontWeight("bold"); }
        else if (mV === "Moderate")     { mCell.setBackground("#fff9c4").setFontColor("#f57f17").setFontWeight("bold"); }
        else if (mV === "Inconsistent") { mCell.setBackground("#ffcdd2").setFontColor("#b71c1c").setFontWeight("bold"); }
        if (yV === "Active")            { yCell.setBackground("#c8e6c9").setFontColor("#1b5e20").setFontWeight("bold"); }
        else if (yV === "Moderate")     { yCell.setBackground("#fff9c4").setFontColor("#f57f17").setFontWeight("bold"); }
        else if (yV === "Inconsistent") { yCell.setBackground("#ffcdd2").setFontColor("#b71c1c").setFontWeight("bold"); }
      } catch(e) {}
    }
  }

  const totalRow    = sec5Row + 2 + outData.length;
  const grandMSer   = outData.reduce((s, r) => s + (Number(r[2]) || 0), 0);
  const grandYSer   = outData.reduce((s, r) => s + (Number(r[5]) || 0), 0);
  const grandYRecords = outData.reduce((s, r) => s + (Number(r[8]) || 0), 0);
  writeTotalRow(stats, totalRow,
    ["TOTAL", "", grandMSer, "", "", grandYSer, "", "", grandYRecords], 9);
  writeRefreshNote(stats, totalRow + 1);
  SpreadsheetApp.getActive().toast(
    "Series Frequency & Verdict updated! (" + monthLabel + " | " + yearLabel + ")",
    "Section 5 Updated", 4);
}

// ════════════════════════════════════════════════════════════
// CHARTS — Status Distribution (Pie) + Monthly Trend (Line)
// ════════════════════════════════════════════════════════════
function runCharts() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const stats = ss.getSheetByName(STATS_SHEET);
  if (!stats) return;
  const existing = stats.getCharts();
  for (let i = 0; i < existing.length; i++) {
    try {
      const title = existing[i].getOptions().get("title") || "";
      if (title.indexOf("DASHBOARD_CHART") !== -1) stats.removeChart(existing[i]);
    } catch(e) {}
  }
  writeStatusChartData(stats);
  writeTrendChartData(stats);
  createStatusChart(stats);
  createTrendChart(stats);
  SpreadsheetApp.getActive().toast("Charts updated!", "Charts Done", 3);
}

function writeStatusChartData(stats) {
  const sec4Row = findSectionRow(stats, "SECTION 4");
  const lastRow = stats.getLastRow();
  const counts = { "Achieved ✓": 0, "On Track": 0, "At Risk !": 0, "Not Started !": 0, "No Target": 0 };
  if (lastRow > sec4Row + 2) {
    const vals = stats.getRange(sec4Row + 2, 8, lastRow - sec4Row - 1, 1).getValues();
    for (let i = 0; i < vals.length; i++) {
      const s = (vals[i][0] || "").toString().trim();
      if (!s || s === "TOTAL" || s === "Status") break;
      if (counts.hasOwnProperty(s)) counts[s]++;
    }
  }
  const rows = [["Status", "Count"]].concat(Object.entries(counts).map(([k, v]) => [k, v]));
  stats.getRange(1, CHART_DATA_COL, rows.length, 2).setValues(rows);
}

function writeTrendChartData(stats) {
  const tz    = Session.getScriptTimeZone();
  const today = new Date();
  const data  = loadExportData();
  const mStart      = new Date(today.getFullYear(), today.getMonth(), 1);
  const mStartStr   = Utilities.formatDate(mStart, tz, "yyyy-MM-dd");
  const todayStr    = Utilities.formatDate(today,  tz, "yyyy-MM-dd");
  const daysInMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate();
  const daysElapsed = today.getDate();
  const targetMap = readTargetMap(stats);
  let totalTarget = 0;
  for (const k in targetMap) totalTarget += (Number(targetMap[k]) || 0);
  const holidaySet = readHolidaySet(stats, tz);
  let totalWorkingDays = 0;
  for (let d = 1; d <= daysInMonth; d++) {
    const dt  = new Date(today.getFullYear(), today.getMonth(), d);
    const key = Utilities.formatDate(dt, tz, "yyyy-MM-dd");
    if (dt.getDay() !== 5 && !holidaySet[key]) totalWorkingDays++;
  }
  const dailyMap = {};
  for (let i = 0; i < data.length; i++) {
    const row      = data[i];
    const distDate = row[COL_DIST_DATE - 1];
    if (!(distDate instanceof Date) || isNaN(distDate)) continue;
    const rowDate = Utilities.formatDate(distDate, tz, "yyyy-MM-dd");
    if (rowDate < mStartStr || rowDate > todayStr) continue;
    const rawName = (row[COL_PROCESSED - 1] || "").toString().trim();
    if (!rawName || !getQCId(rawName)) continue;
    dailyMap[rowDate] = (dailyMap[rowDate] || 0) + getEntryCount(row[COL_ENTRY - 1]);
  }
  const trendRows = [["Day", "Expected", "Actual"]];
  let cumActual = 0, workingDaysPassed = 0;
  for (let d = 1; d <= daysElapsed; d++) {
    const dt  = new Date(today.getFullYear(), today.getMonth(), d);
    const key = Utilities.formatDate(dt, tz, "yyyy-MM-dd");
    const isWorking = (dt.getDay() !== 5 && !holidaySet[key]);
    if (isWorking) workingDaysPassed++;
    const expected = totalWorkingDays > 0
      ? Math.round((totalTarget / totalWorkingDays) * workingDaysPassed)
      : 0;
    cumActual += (dailyMap[key] || 0);
    trendRows.push([d, expected, cumActual]);
  }
  const col = CHART_DATA_COL + 2;
  stats.getRange(1, col, trendRows.length, 3).setValues(trendRows);
  if (trendRows.length < 35) {
    try { stats.getRange(trendRows.length + 1, col, 35 - trendRows.length, 3).clearContent(); } catch(e) {}
  }
}

// Pie chart: QC status distribution
function createStatusChart(stats) {
  const dataRange = stats.getRange(1, CHART_DATA_COL, 6, 2);
  try {
    const chart = stats.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(dataRange)
      .setPosition(16, 1, 5, 5)
      .setOption("title", "DASHBOARD_CHART | Status Distribution")
      .setOption("pieHole", 0.45)
      .setOption("width", 430)
      .setOption("height", 400)
      .setOption("colors", ["#4caf50", "#ffeb3b", "#f44336", "#e91e63", "#9e9e9e"])
      .setOption("legend", { position: "right", textStyle: { fontSize: 11 } })
      .setOption("titleTextStyle", { fontSize: 12, bold: true, color: "#1a237e" })
      .setOption("pieSliceText", "value")
      .build();
    stats.insertChart(chart);
  } catch(e) {}
}

// Line chart: actual vs expected cumulative progress
function createTrendChart(stats) {
  const col = CHART_DATA_COL + 2;
  let dataRows = 1;
  try {
    const check = stats.getRange(1, col, 35, 1).getValues();
    for (let i = 0; i < check.length; i++) {
      if (check[i][0] !== "") dataRows = i + 1;
    }
  } catch(e) {}
  const dataRange = stats.getRange(1, col, dataRows, 3);
  try {
    const chart = stats.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(dataRange)
      .setPosition(16, 5, 5, 5)
      .setOption("title", "DASHBOARD_CHART | Monthly Progress Trend")
      .setOption("width", 480)
      .setOption("height", 400)
      .setOption("hAxis", { title: "Day of Month", gridlines: { count: 10 }, textStyle: { fontSize: 10 } })
      .setOption("vAxis", { title: "Cumulative Records", format: "#,###", textStyle: { fontSize: 10 } })
      .setOption("series", {
        0: { color: "#1565c0", lineWidth: 2, lineDashStyle: [6, 3], pointSize: 3 },
        1: { color: "#2e7d32", lineWidth: 2, pointSize: 4 }
      })
      .setOption("legend", { position: "bottom", textStyle: { fontSize: 11 } })
      .setOption("titleTextStyle", { fontSize: 12, bold: true, color: "#1a237e" })
      .setOption("curveType", "function")
      .build();
    stats.insertChart(chart);
  } catch(e) {}
}

