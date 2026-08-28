/**
 * @fileoverview Sheet builder for monthly Fleet Dispatch Efficiency Analysis.
 * Creates/refreshes compact monthly tabs (e.g. 'August 2026', 'September 2025')
 * with top KPI cards, 20h operating window & real-time concurrency metrics,
 * dynamic formulas, driver allocation table, and premium styling.
 * Purely neutral data without subjective targets.
 */

// Node.js fallback import for testing
let _CONFIG_AB = (typeof CONFIG !== 'undefined') ? CONFIG : null;
let _calculateMonthlyDispatchHours = (typeof calculateMonthlyDispatchHours === 'function') ? calculateMonthlyDispatchHours : null;
let _fetchDispatcherHoursFromTimesheet = (typeof fetchDispatcherHoursFromTimesheet === 'function') ? fetchDispatcherHoursFromTimesheet : null;

if (typeof require !== 'undefined') {
  try {
    if (!_CONFIG_AB) _CONFIG_AB = require('./Config').CONFIG;
    const ds = require('./DispatcherSchedule');
    if (!_calculateMonthlyDispatchHours) _calculateMonthlyDispatchHours = ds.calculateMonthlyDispatchHours;
    if (!_fetchDispatcherHoursFromTimesheet) _fetchDispatcherHoursFromTimesheet = ds.fetchDispatcherHoursFromTimesheet;
  } catch (e) {}
}

/**
 * Builds or refreshes the compact monthly tab for a given month dataset.
 * 
 * @param {Object} monthData - Output bundle from extractDriverDataFromSummary
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function buildMonthlyAnalysisSheet(monthData, targetSpreadsheet) {
  const ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  const monthName = monthData.monthName;
  const year = monthData.year;
  const tabName = `${monthName} ${year}`;

  let sheet = ss.getSheetByName(tabName);
  
  if (!sheet) sheet = ss.getSheetByName(`${monthName} ${year} - Analysis`);
  if (!sheet) sheet = ss.getSheetByName(`${monthName} - Analysis`);
  if (!sheet) sheet = ss.getSheetByName(monthName);

  if (sheet) {
    if (sheet.getName() !== tabName) {
      const existing = ss.getSheetByName(tabName);
      if (existing && existing !== sheet) {
        ss.deleteSheet(existing);
      }
      sheet.setName(tabName);
    }
    sheet.clear();
    sheet.clearConditionalFormatRules();
  } else {
    sheet = ss.insertSheet(tabName);
  }

  const fetcherFn = (typeof fetchDispatcherHoursFromTimesheet === 'function') 
    ? fetchDispatcherHoursFromTimesheet 
    : (_fetchDispatcherHoursFromTimesheet || _calculateMonthlyDispatchHours);
  
  const dispatchLabor = fetcherFn(monthData.monthIndex, year);
  const standardDispatchHours = dispatchLabor.standardHours || dispatchLabor.totalHours || 0;
  const overtimeHours = dispatchLabor.overtimeHours || 0;
  const daysInMonth = new Date(year, monthData.monthIndex + 1, 0).getDate();

  const laborSourceText = (dispatchLabor.source === 'timesheet') 
    ? 'Timesheet (Muhammad, Mariam, Nourween, Mohanad) + Fixed Nour (9h)' 
    : (dispatchLabor.source === 'no_timesheet')
    ? 'No Timesheet Found (0.0 hrs)'
    : `Calendar Schedule (Muhammad 10h, Mariam 10h [Thu OFF, Sat 10h], Nourween 10h, Nour 9h)`;

  const drivers = monthData.drivers || [];
  const driverCount = drivers.length;
  const startDataRow = 14;
  const endDataRow = startDataRow + Math.max(driverCount - 1, 0);
  const totalRow = endDataRow + 1;

  const totalRowsNeeded = Math.max(totalRow + 3, 28);
  const totalCols = 7;
  const matrix = Array.from({ length: totalRowsNeeded }, () => Array(totalCols).fill(''));

  // Row 1: Title Banner
  matrix[0][0] = `🚀 ${monthName.toUpperCase()} ${year} — FLEET DISPATCH EFFICIENCY & LABOR RATIO ANALYSIS`;

  // Row 2: Subtitle / Notice
  if (monthData.totalDriverHours === 0) {
    matrix[1][0] = `⚠️ NOTICE: No active driver road hours recorded in source file. Excluded from Master Summary. | Dispatch Labor: ${laborSourceText}`;
  } else {
    matrix[1][0] = `Generated: ${new Date().toLocaleString()} | Source: ${monthData.fileName} | Dispatch Labor: ${laborSourceText}`;
  }

  // Row 3-5: Top Cumulative Labor KPI Cards
  // Card 1: TOTAL FLEET DRIVER HOURS (Cols A:B)
  matrix[2][0] = 'TOTAL FLEET DRIVER HOURS';
  matrix[3][0] = driverCount > 0 ? `=SUM(B${startDataRow}:B${endDataRow})` : 0;
  matrix[4][0] = 'Active Road Hours';

  // Card 2: STANDARD DISPATCH HOURS (Cols C:D)
  matrix[2][2] = 'STANDARD DISPATCH HOURS';
  matrix[3][2] = standardDispatchHours;
  matrix[4][2] = laborSourceText;

  // Card 3: OVERTIME / ADJUSTMENTS (Col E)
  matrix[2][4] = 'OVERTIME / ADJ (HRS)';
  matrix[3][4] = overtimeHours;
  matrix[4][4] = '✏️ Editable Adjustment';

  // Card 4: TOTAL DISPATCH HOURS (Cols F:G)
  matrix[2][5] = 'TOTAL DISPATCH HOURS';
  matrix[3][5] = `=$C$4+$E$4`;
  matrix[4][5] = 'Standard + Overtime';

  // Row 6-8: Total Labor Overhead Ratio Cards
  // Card 5: FLEET DISPATCH RATIO (Cols A:B)
  matrix[5][0] = 'FLEET DISPATCH RATIO (OVERHEAD)';
  matrix[6][0] = `=IF($A$4>0, $F$4/$A$4, 0)`;
  matrix[7][0] = 'Total Dispatch Hrs / Driver Road Hr';

  // Card 6: SUPPORT MINS / ROAD HOUR (Cols C:D)
  matrix[5][2] = 'SUPPORT MINS / ROAD HOUR';
  matrix[6][2] = `=$A$7*60`;
  matrix[7][2] = 'Minutes of Dispatch per 1 Road Hr';

  // Card 7: DRIVER LEVERAGE RATIO (Cols E:G)
  matrix[5][4] = 'DRIVER LEVERAGE RATIO';
  matrix[6][4] = `=IF($F$4>0, $A$4/$F$4, 0)`;
  matrix[7][4] = 'Driver Road Hours per 1 Dispatch Hr';

  // Row 9-11: 20-Hour Operating Window & Real-Time Concurrency KPI Cards (Neutral)
  // Card 8: EST. CONCURRENT CARS ON ROAD (Cols A:B)
  matrix[8][0] = 'EST. CONCURRENT CARS ON ROAD';
  matrix[9][0] = `=IF($A$4>0, ($A$4/${daysInMonth})/20, 0)`;
  matrix[10][0] = `Active Cars in 20h Window (${daysInMonth} days)`;

  // Card 9: LIVE CARS / ON-DUTY DISPATCHER (Cols C:D)
  matrix[8][2] = 'LIVE CARS / ON-DUTY DISPATCHER';
  matrix[9][2] = `=IF($A$10>0, $A$10/1.95, 0)`;
  matrix[10][2] = 'Real-Time Load (~1.95 Concurrent Disp)';

  // Card 10: DAILY OPERATING WINDOW (Cols E:G)
  matrix[8][4] = 'DAILY OPERATING COVERAGE';
  matrix[9][4] = '20.0 hrs/day';
  matrix[10][4] = '4:00 AM – 12:00 AM (1.95 Avg Disp)';

  // Row 12: Section Divider
  matrix[11][0] = 'DRIVER PERFORMANCE & DISPATCH LABOR ALLOCATION';

  // Row 13: Table Headers
  matrix[12][0] = 'Driver Name';
  matrix[12][1] = 'Driver Active Hours';
  matrix[12][2] = '% Share of Fleet Hours';
  matrix[12][3] = 'Allocated Dispatch Hours';
  matrix[12][4] = 'Completed Trips';
  matrix[12][5] = 'Trips / Driver Hour';
  matrix[12][6] = 'Trips / Dispatch Hour';

  // Rows 14+ : Data Rows
  if (driverCount > 0) {
    for (let i = 0; i < driverCount; i++) {
      const r = startDataRow + i;
      const d = drivers[i];
      const rowIdx = r - 1;

      matrix[rowIdx][0] = d.driverName;
      matrix[rowIdx][1] = d.hours;
      matrix[rowIdx][2] = `=IF($A$4>0, B${r}/$A$4, 0)`;
      matrix[rowIdx][3] = `=B${r}*$A$7`;
      matrix[rowIdx][4] = d.trips;
      matrix[rowIdx][5] = `=IF(B${r}>0, E${r}/B${r}, 0)`;
      matrix[rowIdx][6] = `=IF(D${r}>0, E${r}/D${r}, 0)`;
    }

    // Total Row
    const totRowIdx = totalRow - 1;
    matrix[totRowIdx][0] = 'TOTAL / FLEET AVERAGE';
    matrix[totRowIdx][1] = `=SUM(B${startDataRow}:B${endDataRow})`;
    matrix[totRowIdx][2] = `=SUM(C${startDataRow}:C${endDataRow})`;
    matrix[totRowIdx][3] = `=SUM(D${startDataRow}:D${endDataRow})`;
    matrix[totRowIdx][4] = `=SUM(E${startDataRow}:E${endDataRow})`;
    matrix[totRowIdx][5] = `=IF(B${totalRow}>0, E${totalRow}/B${totalRow}, 0)`;
    matrix[totRowIdx][6] = `=IF(D${totalRow}>0, E${totalRow}/D${totalRow}, 0)`;
  }

  sheet.getRange(1, 1, totalRowsNeeded, totalCols).setValues(matrix);
  applyAnalysisSheetFormatting(sheet, startDataRow, endDataRow, totalRow, driverCount, monthData.totalDriverHours === 0);

  return sheet;
}

function applyAnalysisSheetFormatting(sheet, startDataRow, endDataRow, totalRow, driverCount, isZeroHourMonth) {
  const activeCfg = (typeof CONFIG !== 'undefined') ? CONFIG : _CONFIG_AB;
  const theme = (activeCfg && activeCfg.THEME) ? activeCfg.THEME : {
    headerBg: '#1e293b',
    headerColor: '#ffffff',
    zebraBg: '#f1f5f9'
  };

  // Title Banner
  sheet.getRange('A1:G1').merge()
    .setBackground(theme.headerBg)
    .setFontColor(theme.headerColor)
    .setFontSize(13)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 38);

  // Subtitle
  sheet.getRange('A2:G2').merge()
    .setBackground(isZeroHourMonth ? '#7f1d1d' : '#334155')
    .setFontColor('#ffffff')
    .setFontSize(9)
    .setFontStyle(isZeroHourMonth ? 'normal' : 'italic')
    .setFontWeight(isZeroHourMonth ? 'bold' : 'normal')
    .setHorizontalAlignment('left');
  sheet.setRowHeight(2, 22);

  // Row 1 KPI Cards (Driver Hours & Dispatch Hours)
  sheet.getRange('A3:B3').merge().setValue('TOTAL FLEET DRIVER HOURS');
  sheet.getRange('A4:B4').merge();
  sheet.getRange('A5:B5').merge().setValue('Active Road Hours');
  formatKpiCard(sheet, 'A3:B5', 'A4:B4', '#,##0.0 "hrs"', '#1e3a8a');

  sheet.getRange('C3:D3').merge().setValue('STANDARD DISPATCH HOURS');
  sheet.getRange('C4:D4').merge();
  sheet.getRange('C5:D5').merge();
  formatKpiCard(sheet, 'C3:D5', 'C4:D4', '#,##0.0 "hrs"', '#1e293b');

  formatKpiCard(sheet, 'E3:E5', 'E4:E4', '#,##0.0 "hrs"', '#b45309');
  sheet.getRange('E3:E5').setBackground('#fef3c7');
  sheet.getRange('E4').setNote('👉 EDIT THIS CELL to adjust overtime hours.');

  sheet.getRange('F3:G3').merge().setValue('TOTAL DISPATCH HOURS');
  sheet.getRange('F4:G4').merge();
  sheet.getRange('F5:G5').merge().setValue('Standard + Overtime');
  formatKpiCard(sheet, 'F3:G5', 'F4:G4', '#,##0.0 "hrs"', '#0f172a');

  // Row 2 KPI Cards (Total Labor Ratios)
  sheet.getRange('A6:B6').merge().setValue('FLEET DISPATCH RATIO');
  sheet.getRange('A7:B7').merge();
  sheet.getRange('A8:B8').merge().setValue('Total Dispatch Hrs / Road Hr');
  formatKpiCard(sheet, 'A6:B8', 'A7:B7', '0.000', '#0369a1');

  sheet.getRange('C6:D6').merge().setValue('SUPPORT MINS / ROAD HR');
  sheet.getRange('C7:D7').merge();
  sheet.getRange('C8:D8').merge().setValue('Dispatch Mins per 1 Road Hr');
  formatKpiCard(sheet, 'C6:D8', 'C7:D7', '0.0 "mins"', '#4338ca');

  sheet.getRange('E6:G6').merge().setValue('DRIVER LEVERAGE RATIO');
  sheet.getRange('E7:G7').merge();
  sheet.getRange('E8:G8').merge().setValue('Road Hrs driven per 1 Dispatch Hr');
  formatKpiCard(sheet, 'E6:G8', 'E7:G7', '0.00 "x"', '#15803d');

  // Row 3 KPI Cards (20h Operating Window & Real-Time Concurrency - Neutral)
  sheet.getRange('A9:B9').merge().setValue('EST. CONCURRENT CARS ON ROAD');
  sheet.getRange('A10:B10').merge();
  sheet.getRange('A11:B11').merge();
  formatKpiCard(sheet, 'A9:B11', 'A10:B10', '0.0 "cars"', '#0d9488');

  sheet.getRange('C9:D9').merge().setValue('LIVE CARS / ON-DUTY DISPATCHER');
  sheet.getRange('C10:D10').merge();
  sheet.getRange('C11:D11').merge();
  formatKpiCard(sheet, 'C9:D11', 'C10:D10', '0.00 "cars/disp"', '#0284c7');

  sheet.getRange('E9:G9').merge().setValue('DAILY OPERATING COVERAGE');
  sheet.getRange('E10:G10').merge();
  sheet.getRange('E11:G11').merge();
  formatKpiCard(sheet, 'E9:G11', 'E10:G10', '@', '#475569');

  sheet.setRowHeight(3, 20);
  sheet.setRowHeight(4, 30);
  sheet.setRowHeight(5, 18);
  sheet.setRowHeight(6, 20);
  sheet.setRowHeight(7, 30);
  sheet.setRowHeight(8, 18);
  sheet.setRowHeight(9, 20);
  sheet.setRowHeight(10, 30);
  sheet.setRowHeight(11, 18);

  // Table Section Header
  sheet.getRange('A12:G12').merge()
    .setBackground('#0f172a')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('left');
  sheet.setRowHeight(12, 26);

  const tableHeaderRange = sheet.getRange('A13:G13');
  tableHeaderRange
    .setBackground('#334155')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
  sheet.setRowHeight(13, 30);

  if (driverCount > 0) {
    const dataRange = sheet.getRange(startDataRow, 1, driverCount, 7);
    dataRange.setFontSize(10).setVerticalAlignment('middle');

    sheet.getRange(startDataRow, 1, driverCount, 1).setHorizontalAlignment('left');
    sheet.getRange(startDataRow, 2, driverCount, 6).setHorizontalAlignment('right');

    sheet.getRange(startDataRow, 2, driverCount, 1).setNumberFormat('#,##0.0');
    sheet.getRange(startDataRow, 3, driverCount, 1).setNumberFormat('0.0%');
    sheet.getRange(startDataRow, 4, driverCount, 1).setNumberFormat('#,##0.0');
    sheet.getRange(startDataRow, 5, driverCount, 1).setNumberFormat('#,##0');
    sheet.getRange(startDataRow, 6, driverCount, 1).setNumberFormat('0.00');
    sheet.getRange(startDataRow, 7, driverCount, 1).setNumberFormat('0.00');

    for (let r = startDataRow; r <= endDataRow; r++) {
      sheet.getRange(r, 1, 1, 7).setBackground(r % 2 === 0 ? theme.zebraBg : '#ffffff');
      sheet.setRowHeight(r, 22);
    }

    dataRange.setBorder(true, true, true, true, true, true, '#e2e8f0', SpreadsheetApp.BorderStyle.SOLID);

    const totalRange = sheet.getRange(totalRow, 1, 1, 7);
    totalRange
      .setBackground('#e2e8f0')
      .setFontWeight('bold')
      .setFontSize(10)
      .setVerticalAlignment('middle');
    sheet.getRange(totalRow, 1).setHorizontalAlignment('left');
    sheet.getRange(totalRow, 2, 1, 6).setHorizontalAlignment('right');

    sheet.getRange(totalRow, 2).setNumberFormat('#,##0.0');
    sheet.getRange(totalRow, 3).setNumberFormat('0.0%');
    sheet.getRange(totalRow, 4).setNumberFormat('#,##0.0');
    sheet.getRange(totalRow, 5).setNumberFormat('#,##0');
    sheet.getRange(totalRow, 6).setNumberFormat('0.00');
    sheet.getRange(totalRow, 7).setNumberFormat('0.00');

    totalRange.setBorder(true, true, true, true, false, false, '#475569', SpreadsheetApp.BorderStyle.SOLID);
    sheet.setRowHeight(totalRow, 26);
  }

  sheet.setColumnWidth(1, 190);
  sheet.setColumnWidth(2, 140);
  sheet.setColumnWidth(3, 140);
  sheet.setColumnWidth(4, 160);
  sheet.setColumnWidth(5, 130);
  sheet.setColumnWidth(6, 140);
  sheet.setColumnWidth(7, 150);

  sheet.setFrozenRows(13);
}

function formatKpiCard(sheet, fullRangeA1, valueRangeA1, numberFormat, valueColor) {
  const fullRange = sheet.getRange(fullRangeA1);
  fullRange
    .setBackground('#f8fafc')
    .setBorder(true, true, true, true, false, false, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);

  const firstRowRange = sheet.getRange(fullRangeA1.split(':')[0]);
  firstRowRange
    .setFontSize(8)
    .setFontWeight('bold')
    .setFontColor('#64748b')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  const valRange = sheet.getRange(valueRangeA1);
  valRange
    .setFontSize(15)
    .setFontWeight('bold')
    .setFontColor(valueColor || '#0f172a')
    .setNumberFormat(numberFormat)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  const lastCell = fullRangeA1.split(':')[1];
  const colLetter = lastCell.charAt(0);
  const rowNum = lastCell.slice(1);
  const subtextRange = sheet.getRange(`${colLetter}${rowNum}`);
  subtextRange
    .setFontSize(8)
    .setFontColor('#94a3b8')
    .setFontStyle('italic')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    buildMonthlyAnalysisSheet,
    applyAnalysisSheetFormatting,
    formatKpiCard
  };
}
