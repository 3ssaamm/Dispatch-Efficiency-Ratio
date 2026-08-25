/**
 * @fileoverview Sheet builder for monthly Fleet Dispatch Efficiency Analysis.
 * Creates/refreshes the '[Month] - Analysis' tab with top KPI cards,
 * live dynamic formulas, driver allocation table, and premium styling.
 */

/**
 * Builds or refreshes the Analysis tab for a given month dataset.
 * 
 * @param {Object} monthData - Output bundle from extractDriverDataFromSummary
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet] - Optional target sheet (defaults to Active)
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function buildMonthlyAnalysisSheet(monthData, targetSpreadsheet) {
  const ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  const monthName = monthData.monthName;
  const year = monthData.year;
  const tabName = `${monthName} ${year} - Analysis`;

  let sheet = ss.getSheetByName(tabName);
  // Also check if existing tab is named just "[Month] - Analysis"
  if (!sheet) {
    sheet = ss.getSheetByName(`${monthName} - Analysis`);
  }

  if (sheet) {
    sheet.clear();
    // Clear all conditional formatting & data validations
    sheet.clearConditionalFormatRules();
  } else {
    sheet = ss.insertSheet(tabName);
  }

  // Calculate calendar dispatch labor hours
  const scheduleResult = calculateMonthlyDispatchHours(monthData.monthIndex, year);
  const standardDispatchHours = scheduleResult.totalHours;

  const drivers = monthData.drivers || [];
  const driverCount = drivers.length;
  const startDataRow = 11;
  const endDataRow = startDataRow + Math.max(driverCount - 1, 0);
  const totalRow = endDataRow + 1;

  // Build 2D values array for the entire sheet for high-speed batch writing
  const totalRowsNeeded = Math.max(totalRow + 3, 25);
  const totalCols = 7; // Columns A to G
  const matrix = Array.from({ length: totalRowsNeeded }, () => Array(totalCols).fill(''));

  // Row 1: Title Banner
  matrix[0][0] = `🚀 ${monthName.toUpperCase()} ${year} — FLEET DISPATCH EFFICIENCY & LABOR RATIO ANALYSIS`;

  // Row 2: Subtitle / Metadata
  matrix[1][0] = `Generated: ${new Date().toLocaleString()} | Source File: ${monthData.fileName} | Working Days: ${scheduleResult.daysInMonth} days`;

  // Row 3-5: Top KPI Cards (Row 1 of Cards)
  // Card 1: TOTAL FLEET DRIVER HOURS (Cols A:B)
  matrix[2][0] = 'TOTAL FLEET DRIVER HOURS';
  matrix[3][0] = driverCount > 0 ? `=SUM(B${startDataRow}:B${endDataRow})` : 0;
  matrix[4][0] = 'Active Road Hours';

  // Card 2: STANDARD DISPATCH HOURS (Cols C:D)
  matrix[2][2] = 'STANDARD DISPATCH HOURS';
  matrix[3][2] = standardDispatchHours;
  matrix[4][2] = `Calendar Base (${scheduleResult.daysInMonth} days, 4 Dispatchers)`;

  // Card 3: OVERTIME / ADJUSTMENTS (Col E)
  matrix[2][4] = 'OVERTIME / ADJ (HRS)';
  matrix[3][4] = 0; // Editable input cell
  matrix[4][4] = '✏️ Editable Adjustment';

  // Card 4: TOTAL DISPATCH HOURS (Cols F:G)
  matrix[2][5] = 'TOTAL DISPATCH HOURS';
  matrix[3][5] = `=$C$4+$E$4`;
  matrix[4][5] = 'Standard + Overtime';

  // Row 6-8: Ratio KPI Cards (Row 2 of Cards)
  // Card 5: FLEET DISPATCH RATIO (Cols A:B)
  matrix[5][0] = 'FLEET DISPATCH RATIO (OVERHEAD)';
  matrix[6][0] = `=IF($A$4>0, $F$4/$A$4, 0)`;
  matrix[7][0] = 'Dispatch Hours / Driver Hour';

  // Card 6: SUPPORT MINS / ROAD HOUR (Cols C:D)
  matrix[5][2] = 'SUPPORT MINS / ROAD HOUR';
  matrix[6][2] = `=$A$7*60`;
  matrix[7][2] = 'Minutes of Dispatch per 1 Road Hr';

  // Card 7: DRIVER LEVERAGE RATIO (Cols E:G)
  matrix[5][4] = 'DRIVER LEVERAGE RATIO';
  matrix[6][4] = `=IF($F$4>0, $A$4/$F$4, 0)`;
  matrix[7][4] = 'Driver Road Hours per 1 Dispatch Hr';

  // Row 9: Section Divider / Title
  matrix[8][0] = 'DRIVER PERFORMANCE & DISPATCH LABOR ALLOCATION';

  // Row 10: Table Headers
  matrix[9][0] = 'Driver Name';
  matrix[9][1] = 'Driver Active Hours';
  matrix[9][2] = '% Share of Fleet Hours';
  matrix[9][3] = 'Allocated Dispatch Hours';
  matrix[9][4] = 'Completed Trips';
  matrix[9][5] = 'Trips / Driver Hour';
  matrix[9][6] = 'Trips / Dispatch Hour';

  // Rows 11+ : Data Rows
  if (driverCount > 0) {
    for (let i = 0; i < driverCount; i++) {
      const r = startDataRow + i; // 1-indexed row in sheet
      const d = drivers[i];
      const rowIdx = r - 1; // 0-indexed in matrix

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

  // Batch write all data into the sheet
  sheet.getRange(1, 1, totalRowsNeeded, totalCols).setValues(matrix);

  // Apply Styling and Formatting
  applyAnalysisSheetFormatting(sheet, startDataRow, endDataRow, totalRow, driverCount);

  return sheet;
}

/**
 * Applies visual styling, borders, card backgrounds, alignments, and number formats.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} startDataRow
 * @param {number} endDataRow
 * @param {number} totalRow
 * @param {number} driverCount
 */
function applyAnalysisSheetFormatting(sheet, startDataRow, endDataRow, totalRow, driverCount) {
  const theme = (typeof CONFIG !== 'undefined' && CONFIG.THEME) ? CONFIG.THEME : {
    headerBg: '#1e293b',
    headerColor: '#ffffff',
    kpiCardBg: '#f8fafc',
    kpiCardBorder: '#cbd5e1',
    kpiValueColor: '#0f172a',
    zebraBg: '#f1f5f9'
  };

  // Row 1: Title Banner
  sheet.getRange('A1:G1').merge()
    .setBackground(theme.headerBg)
    .setFontColor(theme.headerColor)
    .setFontSize(13)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 38);

  // Row 2: Subtitle
  sheet.getRange('A2:G2').merge()
    .setBackground('#334155')
    .setFontColor('#cbd5e1')
    .setFontSize(9)
    .setFontStyle('italic')
    .setHorizontalAlignment('left');
  sheet.setRowHeight(2, 22);

  // Row 3-5: KPI Cards (Row 1)
  // Card 1: A3:B5
  sheet.getRange('A3:B3').merge().setValue('TOTAL FLEET DRIVER HOURS');
  sheet.getRange('A4:B4').merge();
  sheet.getRange('A5:B5').merge().setValue('Active Road Hours');
  formatKpiCard(sheet, 'A3:B5', 'A4:B4', '#,##0.0 "hrs"', '#1e3a8a');

  // Card 2: C3:D5
  sheet.getRange('C3:D3').merge().setValue('STANDARD DISPATCH HOURS');
  sheet.getRange('C4:D4').merge();
  sheet.getRange('C5:D5').merge();
  formatKpiCard(sheet, 'C3:D5', 'C4:D4', '#,##0.0 "hrs"', '#1e293b');

  // Card 3: E3:E5 (Overtime Editable Input)
  formatKpiCard(sheet, 'E3:E5', 'E4:E4', '#,##0.0 "hrs"', '#b45309');
  sheet.getRange('E3:E5').setBackground('#fef3c7'); // Soft amber highlight for input cell
  sheet.getRange('E4').setNote('👉 EDIT THIS CELL to add overtime or adjustment dispatch hours.');

  // Card 4: F3:G5
  sheet.getRange('F3:G3').merge().setValue('TOTAL DISPATCH HOURS');
  sheet.getRange('F4:G4').merge();
  sheet.getRange('F5:G5').merge().setValue('Standard + Overtime');
  formatKpiCard(sheet, 'F3:G5', 'F4:G4', '#,##0.0 "hrs"', '#0f172a');

  // Row 6-8: KPI Cards (Row 2)
  // Card 5: A6:B8 (Fleet Dispatch Ratio)
  sheet.getRange('A6:B6').merge().setValue('FLEET DISPATCH RATIO');
  sheet.getRange('A7:B7').merge();
  sheet.getRange('A8:B8').merge().setValue('Dispatch Hours / Road Hr');
  formatKpiCard(sheet, 'A6:B8', 'A7:B7', '0.000', '#0369a1');

  // Card 6: C6:D8 (Support Mins / Road Hr)
  sheet.getRange('C6:D6').merge().setValue('SUPPORT MINS / ROAD HR');
  sheet.getRange('C7:D7').merge();
  sheet.getRange('C8:D8').merge().setValue('Dispatch Mins per 1 Road Hr');
  formatKpiCard(sheet, 'C6:D8', 'C7:D7', '0.0 "mins"', '#4338ca');

  // Card 7: E6:G8 (Driver Leverage Ratio)
  sheet.getRange('E6:G6').merge().setValue('DRIVER LEVERAGE RATIO');
  sheet.getRange('E7:G7').merge();
  sheet.getRange('E8:G8').merge().setValue('Road Hrs driven per 1 Dispatch Hr');
  formatKpiCard(sheet, 'E6:G8', 'E7:G7', '0.00 "x"', '#15803d');

  sheet.setRowHeight(3, 20);
  sheet.setRowHeight(4, 32);
  sheet.setRowHeight(5, 18);
  sheet.setRowHeight(6, 20);
  sheet.setRowHeight(7, 32);
  sheet.setRowHeight(8, 18);

  // Row 9: Table Section Header
  sheet.getRange('A9:G9').merge()
    .setBackground('#0f172a')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('left');
  sheet.setRowHeight(9, 26);

  // Row 10: Table Column Headers
  const tableHeaderRange = sheet.getRange('A10:G10');
  tableHeaderRange
    .setBackground('#334155')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
  sheet.setRowHeight(10, 30);

  // Format Data Rows
  if (driverCount > 0) {
    const dataRange = sheet.getRange(startDataRow, 1, driverCount, 7);
    dataRange.setFontSize(10).setVerticalAlignment('middle');

    // Alignments
    sheet.getRange(startDataRow, 1, driverCount, 1).setHorizontalAlignment('left'); // Driver Name
    sheet.getRange(startDataRow, 2, driverCount, 6).setHorizontalAlignment('right'); // Numbers

    // Number Formats
    sheet.getRange(startDataRow, 2, driverCount, 1).setNumberFormat('#,##0.0'); // Driver Hours
    sheet.getRange(startDataRow, 3, driverCount, 1).setNumberFormat('0.0%');    // % Share
    sheet.getRange(startDataRow, 4, driverCount, 1).setNumberFormat('#,##0.0'); // Allocated Hours
    sheet.getRange(startDataRow, 5, driverCount, 1).setNumberFormat('#,##0');   // Trips
    sheet.getRange(startDataRow, 6, driverCount, 1).setNumberFormat('0.00');    // Trips / Driver Hr
    sheet.getRange(startDataRow, 7, driverCount, 1).setNumberFormat('0.00');    // Trips / Disp Hr

    // Alternating Row Backgrounds (Zebra striping)
    for (let r = startDataRow; r <= endDataRow; r++) {
      if (r % 2 === 0) {
        sheet.getRange(r, 1, 1, 7).setBackground(theme.zebraBg);
      } else {
        sheet.getRange(r, 1, 1, 7).setBackground('#ffffff');
      }
      sheet.setRowHeight(r, 22);
    }

    // Grid Borders for table
    dataRange.setBorder(true, true, true, true, true, true, '#e2e8f0', SpreadsheetApp.BorderStyle.SOLID);

    // Total Row Styling
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

    // Total row borders: top thin, bottom double
    totalRange.setBorder(true, true, true, true, false, false, '#475569', SpreadsheetApp.BorderStyle.SOLID);
    sheet.setRowHeight(totalRow, 26);
  }

  // Column Widths
  sheet.setColumnWidth(1, 190); // Driver Name
  sheet.setColumnWidth(2, 140); // Driver Active Hours
  sheet.setColumnWidth(3, 140); // % Share
  sheet.setColumnWidth(4, 160); // Allocated Disp Hours
  sheet.setColumnWidth(5, 130); // Completed Trips
  sheet.setColumnWidth(6, 140); // Trips / Driver Hr
  sheet.setColumnWidth(7, 150); // Trips / Disp Hr

  // Hide excessive empty gridlines/columns
  sheet.setFrozenRows(10);
}

/**
 * Helper to style a 3-row KPI Card component.
 */
function formatKpiCard(sheet, fullRangeA1, valueRangeA1, numberFormat, valueColor) {
  const fullRange = sheet.getRange(fullRangeA1);
  fullRange
    .setBackground('#f8fafc')
    .setBorder(true, true, true, true, false, false, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);

  // Label row (1st row)
  const firstRowRange = sheet.getRange(fullRangeA1.split(':')[0]);
  firstRowRange
    .setFontSize(8)
    .setFontWeight('bold')
    .setFontColor('#64748b')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // Value row (2nd row)
  const valRange = sheet.getRange(valueRangeA1);
  valRange
    .setFontSize(16)
    .setFontWeight('bold')
    .setFontColor(valueColor || '#0f172a')
    .setNumberFormat(numberFormat)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // Subtext row (3rd row)
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

// Node.js module export for testing
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    buildMonthlyAnalysisSheet,
    applyAnalysisSheetFormatting,
    formatKpiCard
  };
}
