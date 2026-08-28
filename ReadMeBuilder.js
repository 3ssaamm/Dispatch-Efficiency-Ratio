/**
 * @fileoverview Read Me & Metric Dictionary Guide Sheet Builder.
 * Creates a clean '📖 Read Me & Guide' tab in Google Sheets explaining
 * how to read the file, metrics definitions, 20h operating window logic,
 * and real-time concurrency ratios.
 */

function buildReadMeSheet(targetSpreadsheet) {
  const ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  const tabName = '📖 Read Me & Guide';

  let sheet = ss.getSheetByName(tabName);
  if (!sheet) sheet = ss.getSheetByName('Read Me & Guide');
  if (!sheet) sheet = ss.getSheetByName('Read Me');

  if (sheet) {
    sheet.clear();
    sheet.clearConditionalFormatRules();
  } else {
    sheet = ss.insertSheet(tabName, 0);
  }

  const rows = [
    // Banner
    ['📖 FLEET DISPATCH EFFICIENCY ENGINE — USER GUIDE & METRIC DICTIONARY', '', '', ''],
    ['Reference manual for understanding Fleet Labor Overhead, 20h Operating Window, and Real-Time Concurrency Metrics.', '', '', ''],
    ['', '', '', ''],

    // Section 1: Overview
    ['1. EXECUTIVE OVERVIEW: LABOR OVERHEAD VS. REAL-TIME CONCURRENCY', '', '', ''],
    ['Concept', 'Operational Definition & Calculation', 'Target / Healthy Range', 'Business Interpretation'],
    ['Daily Operating Window', 'The fleet operates 20 active hours per day (4:00 AM – 12:00 AM Midnight) across 2 overlapping shifts.', '20.0 Hours / Day', 'Defines the daily timeframe over which active road cars and dispatchers operate.'],
    ['Concurrent Dispatchers On-Duty', 'Total daily scheduled weekday dispatch hours (39h) divided by the 20h window = ~1.95 dispatchers on-duty at any hour.', '~1.95 Dispatchers', 'Shows that individual shifts run lean with approx. 2 dispatchers active concurrently.'],
    ['Total Labor Overhead Ratio', 'Total Monthly Dispatch Hours / Total Driver Road Hours.', '0.70 – 1.10', 'Measures full payroll labor balance. Sits near ~1:1 because 39 dispatch hours cover a 20h window.'],
    ['Real-Time Driver Concurrency', '(Total Driver Hours in Month / Operating Days) / 20 Hours.', 'Varies with fleet size', 'Estimates the average number of active cars physically on the road at any given hour.'],
    ['Live Cars per On-Duty Dispatcher', 'Real-Time Driver Concurrency / 1.95 Concurrent Dispatchers.', '3.0 – 5.0 Cars / Disp', 'Real-time operational load: How many live cars 1 on-duty dispatcher is actively managing simultaneously.'],
    ['Desk Capacity Utilization (%)', '(Live Cars per Dispatcher / 4.0 Optimal Target) × 100.', '75% – 100%', 'Measures how close the dispatch desk is running relative to optimal peak capacity (1:4 ratio).'],
    ['', '', '', ''],

    // Section 2: Core Metrics
    ['2. DETAILED METRICS & KPI DICTIONARY', '', '', ''],
    ['Metric Name', 'Formula / Data Source', 'Unit of Measure', 'How to Interpret'],
    ['Driver Active Hours', 'Sum of active driving/shift hours logged by all drivers in the monthly balance file.', 'Hours (hrs)', 'Direct measure of fleet road output. Months with 0 driver hours are excluded from summary.'],
    ['Standard Dispatch Hours', 'Fetched from Working Time timesheet (Muhammad, Mariam, Nourween, Mohanad) + Fixed Nour (9h).', 'Hours (hrs)', 'Total baseline contractual labor provided by dispatchers in the month.'],
    ['Overtime / Adjustments (OT)', 'Logged overtime from Working Time timesheet. Also editable on cell E4 of monthly tabs.', 'Hours (hrs)', 'Additional hours beyond regular shift schedules for peak demand/coverage.'],
    ['Total Dispatch Labor Hours', 'Standard Dispatch Hours + Overtime / Adjustments.', 'Hours (hrs)', 'Total human labor overhead dedicated to managing fleet operations.'],
    ['Support Minutes / Road Hour', 'Fleet Dispatch Ratio × 60.', 'Minutes / Road Hr', 'Minutes of dispatcher attention required to support 1 driver on the road for 1 hour.'],
    ['Driver Leverage Ratio', 'Total Driver Road Hours / Total Dispatch Hours.', 'Ratio (x)', 'Driver road hours generated for every 1 hour of dispatch labor (e.g. 1.1x).'],
    ['Trips per Driver Hour', 'Completed Trips / Driver Active Hours.', 'Trips / Hr', 'Driver productivity: Average trips fulfilled per road driving hour.'],
    ['Trips per Dispatch Hour', 'Completed Trips / Total Dispatch Hours.', 'Trips / Hr', 'Dispatch efficiency: Average completed rides processed per dispatch labor hour.'],
    ['', '', '', ''],

    // Section 3: Dispatcher Roster
    ['3. DISPATCH TEAM ROSTER & SCHEDULE SUMMARY', '', '', ''],
    ['Staff Member', 'Data Source', 'Shift Schedule & Working Days', 'Status in Fleet Overhead'],
    ['Muhammad', 'Working Time Timesheet', '10.0 hrs/day (Mon, Tue, Wed, Thu, Fri | Sat & Sun OFF)', '✅ Counted (Timesheet WT + OT)'],
    ['Mariam', 'Working Time Timesheet', '10.0 hrs/day (Mon, Tue, Wed, Fri, Sat | Thu & Sun OFF)', '✅ Counted (Works Sat 10h, Thu OFF)'],
    ['Nourween', 'Working Time Timesheet', '10.0 hrs/day (Mon, Tue, Wed, Thu, Fri | Sat & Sun OFF)', '✅ Counted (Active in 2026 tabs)'],
    ['Mohanad', 'Working Time Timesheet', 'Logged Hours in Timesheet (WT + OT)', '✅ Counted across all months'],
    ['Nour', 'Fixed Schedule', '9.0 hrs/day (Mon, Tue, Wed, Thu, Fri | Sat & Sun OFF)', '✅ Counted (Fixed 9h Mon-Fri)'],
    ['Abdulrahman', 'Working Time Timesheet', 'Support Staff', '❌ Excluded from dispatch labor'],
    ['Fares', 'Working Time Timesheet', 'Support Staff', '❌ Disregarded / Excluded']
  ];

  const numRows = rows.length;
  const numCols = 4;
  const range = sheet.getRange(1, 1, numRows, numCols);

  // Set all cell formats as plain text to eliminate formula errors
  range.setNumberFormat('@');
  range.setValues(rows);

  // Apply Styling
  sheet.getRange('A1:D1').merge()
    .setBackground('#1e293b')
    .setFontColor('#ffffff')
    .setFontSize(13)
    .setFontWeight('bold')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 38);

  sheet.getRange('A2:D2').merge()
    .setBackground('#334155')
    .setFontColor('#cbd5e1')
    .setFontSize(9)
    .setFontStyle('italic')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(2, 22);

  const secHeaders = [4, 14, 26];
  for (const sRow of secHeaders) {
    sheet.getRange(sRow, 1, 1, 4).merge()
      .setBackground('#0f172a')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setFontSize(10)
      .setVerticalAlignment('middle');
    sheet.setRowHeight(sRow, 26);
  }

  const tableHeaders = [5, 15, 27];
  for (const tRow of tableHeaders) {
    sheet.getRange(tRow, 1, 1, 4)
      .setBackground('#475569')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setFontSize(9)
      .setVerticalAlignment('middle');
    sheet.setRowHeight(tRow, 26);
  }

  // Data rows styling
  const dataSections = [
    { start: 6, end: 12 },
    { start: 16, end: 24 },
    { start: 28, end: 34 }
  ];

  for (const sec of dataSections) {
    for (let r = sec.start; r <= sec.end; r++) {
      sheet.getRange(r, 1, 1, 4)
        .setBackground(r % 2 === 0 ? '#f8fafc' : '#ffffff')
        .setFontSize(9)
        .setVerticalAlignment('middle');
      sheet.getRange(r, 1).setFontWeight('bold');
      sheet.setRowHeight(r, 24);
    }
    sheet.getRange(sec.start, 1, sec.end - sec.start + 1, 4)
      .setBorder(true, true, true, true, true, true, '#e2e8f0', SpreadsheetApp.BorderStyle.SOLID);
  }

  sheet.setColumnWidth(1, 230);
  sheet.setColumnWidth(2, 420);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 380);

  SpreadsheetApp.getActiveSpreadsheet().toast('Read Me & Guide updated!', 'Fleet Tools', 3);
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    buildReadMeSheet
  };
}
