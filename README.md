# Analytics & Performance Dashboard
### Google Apps Script — Google Sheets

A full-featured team performance analytics dashboard built with Google Apps Script, designed for managing and monitoring large-scale data processing operations with multi-section reporting, target tracking, and consistency analysis.

---

> **Built by:** Mushfiqur Rahman  
> **AI Assistance:** Claude (Anthropic) — architecture, logic design, and iterative development support  
> *Concept, requirements, and domain knowledge by Mushfiqur Rahman. Code structure and implementation developed collaboratively with Claude.*

---

## Features

### 6-Section Dashboard

| Section | Description |
|---|---|
| **Section 1** | Date-wise entry statistics — select any date to see records by QC staff |
| **Section 2** | Weekly performance — active days, total series, daily averages |
| **Section 3** | Monthly performance — full month aggregation per QC |
| **Section 4** | Target & Progress — cycle-based achievement tracking with status flags |
| **Section 5** | Series Frequency & Verdict — submission consistency analysis |


### Key Capabilities

- **Fuzzy Name Matching** — maps raw input names to QC IDs using keyword patterns and word-boundary regex
- **Custom Cycle Dates** — Section 4 supports non-calendar cycles (e.g. Mar 29 – Apr 28)
- **Working Day Calculation** — excludes weekly off-days and public holidays automatically
- **Status Color Coding** — Achieved ✓ / On Track / At Risk ! / Not Started ! with color highlights
- **Auto Charts** — pie chart (status distribution) + line chart (actual vs expected trend)
- **Date Dropdown** — auto-populated from available data in the export sheet

---

## How It Works

```
Export Data Sheet  →  Apps Script reads raw data
                   →  Fuzzy matches QC names
                   →  Aggregates by date / week / month / cycle
                   →  Writes formatted output to Statistics Sheet
                   →  Applies conditional color formatting
                   →  Renders charts
```

---

## Setup Instructions

### 1. Prepare Your Google Sheet
- Create a sheet named **`Export Data`** with your raw data starting at row 4
- Required columns:
  - Column G: Series number
  - Column H: Entry count (records)
  - Column I: Processed by (QC staff name)
  - Column J: Distribution date

### 2. Configure QC Staff List
Open `Code.gs` and edit the `QC_LIST` array at the top:

```javascript
const QC_LIST = [
  { qcId: "QC-001", fullName: "Quality Controller 01", keywords: ["qc01"] },
  { qcId: "QC-002", fullName: "Quality Controller 02", keywords: ["qc02"] },
  // Add your team members here
];
```

`keywords` are the name fragments that appear in your raw data. Multiple keywords per person are supported.

### 3. Add the Script
- Open your Google Sheet → **Extensions → Apps Script**
- Paste the contents of `Code.gs`
- Save and run `setupStatisticsSheet` once

### 4. Set Monthly Targets
- After setup, enter targets in **Column O** (highlighted yellow)
- Press **Refresh All** from the menu

### 5. Add Public Holidays
- Enter holiday dates in **Column Q** (DD/MM/YYYY format)
- These are excluded from working day calculations

### 6. Optional — Custom Cycle Dates (Section 4)
- Enter a custom cycle start in **Cell S2**
- Enter a custom cycle end in **Cell T2**
- Leave blank to use the default calendar month

---

## Configuration Reference

```javascript
// Column positions in the Export Data sheet
const COL_SERIES    = 7;   // Series number
const COL_ENTRY     = 8;   // Entry count
const COL_PROCESSED = 9;   // QC staff name
const COL_DIST_DATE = 10;  // Distribution date
const EXPORT_HEADER = 3;   // Header row number

// Sheet names
const EXPORT_SHEET = "Export Data";
const STATS_SHEET  = "Statistics";

// Weekly off-day
// Default: Friday (day index 5). Change to 0 (Sunday) if needed.
// Find: if (dow !== 5 && !holidaySet[key])
```

---

## Status Logic (Section 4)

| Status | Condition |
|---|---|
| **Achieved ✓** | Records ≥ Target |
| **On Track** | On pace to finish within cycle |
| **At Risk !** | Required daily rate exceeds normal pace by >9% |
| **Not Started !** | Zero records after 3+ days have passed |
| **No Target** | Target not set for this QC |

---

## Verdict Logic (Section 5)

Average working-day gap between submission dates:

| Verdict | Gap |
|---|---|
| **Active** | ≤ 3 working days |
| **Moderate** | 4–5 working days |
| **Inconsistent** | > 5 working days |

---

## Tech Stack

- **Platform:** Google Sheets + Google Apps Script (V8 runtime)
- **Language:** JavaScript (ES6+)
- **Libraries:** None — pure Apps Script
- **Charts:** Google Visualization (built-in Apps Script Charts API)

---

## File Structure

```
/
├── Code.gs        ← Main script (all sections, helpers, charts)
└── README.md      ← This file
```

---

## Notes

- This is a **sanitized public version** — staff names and project identifiers have been replaced with generic placeholders
- The original system runs in production managing a team of 36 QC staff across 254+ data entry operators
- The dashboard is designed for **Bangladesh working week** (Friday off) — adjust the `dow !== 5` condition for other regions

---

## License

MIT — free to use, adapt, and distribute with attribution.
