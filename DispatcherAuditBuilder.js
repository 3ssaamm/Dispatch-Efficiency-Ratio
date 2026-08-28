/**
 * @fileoverview Dispatcher Hours Audit Sheet Builder.
 * Creates a dedicated '🕒 Dispatcher Hours Audit' tab in Google Sheets
 * to verify and audit exact working hours and overtime fetched from the timesheet
 * (Muhammad, Mariam, Nourween) + fixed schedule (Nour 9h Mon-Fri).
 * Zero hours logged = 0.0 hrs (dispatcher was not active/working).
 */

/**
 * Builds or refreshes the '🕒 Dispatcher Hours Audit' sheet tab.
 * 
 * @param {Array<Object>} [allMonthsData]
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function buildDispatcherAuditSheet(allMonthsData, targetSpreadsheet) {
  const ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  const tabName = (typeof CONFIG !== 'undefined' && CONFIG.AUDIT_SHEET_NAME) ? CONFIG.AUDIT_SHEET_NAME : '🕒 Dispatcher Hours Audit';

  let sheet = ss.getSheetByName(tabName);
  if (!sheet) sheet = ss.getSheetByName('Dispatcher Hours Audit');

  if (sheet) {
    sheet.clear();
    sheet.clearConditionalFormatRules();
  } else {
    sheet = ss.insertSheet(tabName, 2);
  }

  let months = allMonthsData;
  if (!months || months.length === 0) {
    months = collectDataFromExistingAnalysisSheets(ss);
  }

  months.sort((a, b) => (a.year - b.year) || (a.monthIndex - b.monthIndex));
  const monthCount = months.length;

  const startAuditRow = 12;
  const endAuditRow = startAuditRow + Math.max(monthCount - 1, 0);
  const totalAuditRow = endAuditRow + 1;

  const totalRowsNeeded = Math.max(totalAuditRow + 15, 35);
  const totalCols = 11;
  const matrix = Array.from({ length: totalRowsNeeded }, () => Array(totalCols).fill(''));

  // Row 1: Title Banner
  matrix[0][0] = '🕒 DISPATCHER WORKING HOURS AUDIT & TIMESHEET RECONCILIATION';

  // Row 2: Subtitle
  matrix[1][0] = `Timesheet Source: Muhammad, Mariam, Nourween | Fixed Shift: Nour (9h/day Mon-Fri) | Excluded: Mohanad, Abdulrahman`;

  // Rows 3-5: Summary Cards
  // Card 1: TOTAL DISPATCH LABOR HOURS (Cols A:C)
  matrix[2][0] = 'TOTAL DISPATCH LABOR HOURS';
  matrix[3][0] = monthCount > 0 ? `=SUM(I${startAuditRow}:I${endAuditRow})` : 0;
  matrix[4][0] = 'All Active Months (Regular + Overtime)';

  // Card 2: TOTAL REGULAR HOURS (Cols D:E)
  matrix[2][3] = 'TOTAL REGULAR HOURS (WT)';
  matrix[3][3] = monthCount > 0 ? `=SUM(G${startAuditRow}:G${endAuditRow})` : 0;
  matrix[4][3] = 'Timesheet WT + Nour Fixed Shift';

  // Card 3: TOTAL OVERTIME HOURS (Cols F:G)
  matrix[2][5] = 'TOTAL OVERTIME HOURS (OT)';
  matrix[3][5] = monthCount > 0 ? `=SUM(H${startAuditRow}:H${endAuditRow})` : 0;
  matrix[4][5] = 'Timesheet Overtime Logged';

  // Card 4: ACTIVE DISPATCHERS (Cols H:K)
  matrix[2][7] = 'ACTIVE DISPATCH TEAM (4)';
  matrix[3][7] = '4 Dispatchers';
  matrix[4][7] = 'Muhammad, Mariam, Nourween, Nour';

  // Row 10: Section 1 Header
  matrix[9][0] = '1. MONTH-BY-MONTH DISPATCHER HOURS AUDIT BREAKDOWN';

  // Row 11: Table Headers
  matrix[10][0] = 'Month & Year';
  matrix[10][1] = 'Labor Data Source';
  matrix[10][2] = 'Muhammad (WT / OT)';
  matrix[10][3] = 'Mariam (WT / OT)';
  matrix[10][4] = 'Nourween (WT / OT)';
  matrix[10][5] = 'Nour (Fixed 9h/day)';
  matrix[10][6] = 'Total Regular WT';
  matrix[10][7] = 'Total Overtime OT';
  matrix[10][8] = 'Total Dispatch Hrs';
  matrix[10][9] = 'Excluded Staff Hours';
  matrix[10][10] = 'Verification Status';

  // Rows 12+: Monthly Breakdown Rows
  if (monthCount > 0) {
    for (let i = 0; i < monthCount; i++) {
      const m = months[i];
      const r = startAuditRow + i;
      const rowIdx = r - 1;

      // Fetch verified timesheet details
      const laborInfo = fetchDispatcherHoursFromTimesheet(m.monthIndex, m.year);

      matrix[rowIdx][0] = `${m.monthName} ${m.year}`;
      
      if (laborInfo.tabFound) {
        matrix[rowIdx][1] = `✅ Timesheet ("${laborInfo.tabName}")`;
      } else if (laborInfo.permissionError) {
        matrix[rowIdx][1] = '🔒 Permission Restricted (Calendar Fallback)';
      } else {
        matrix[rowIdx][1] = '📅 Calendar Schedule Fallback';
      }

      // Find individual staff
      const muhammad = laborInfo.timesheetStaff.find(s => s.name.toLowerCase().includes('muhammad') || s.name.toLowerCase().includes('mohamed'));
      const mariam = laborInfo.timesheetStaff.find(s => s.name.toLowerCase().includes('mariam'));
      const nourween = laborInfo.timesheetStaff.find(s => s.name.toLowerCase().includes('nourween'));
      const nour = laborInfo.fixedStaff.find(s => s.name.toLowerCase() === 'nour');

      if (laborInfo.tabFound) {
        // If timesheet was read, zero hours means the dispatcher was NOT working
        matrix[rowIdx][2] = muhammad ? `${muhammad.regularHours.toFixed(1)} WT + ${muhammad.overtimeHours.toFixed(1)} OT` : '0.0 hrs (Not Working)';
        matrix[rowIdx][3] = mariam ? `${mariam.regularHours.toFixed(1)} WT + ${mariam.overtimeHours.toFixed(1)} OT` : '0.0 hrs (Not Working)';
        matrix[rowIdx][4] = nourween ? `${nourween.regularHours.toFixed(1)} WT + ${nourween.overtimeHours.toFixed(1)} OT` : '0.0 hrs (Not Working)';
        matrix[rowIdx][5] = nour ? `${nour.regularHours.toFixed(1)} hrs (${nour.notes})` : '0.0 hrs (Not Working)';
      } else {
        // Calendar Fallback
        matrix[rowIdx][2] = '10h/day Mon-Fri (Sched)';
        matrix[rowIdx][3] = '10h/day Mon-Wed,Fri,Sat (Sched)';
        matrix[rowIdx][4] = '10h/day Mon-Fri (Sched)';
        matrix[rowIdx][5] = nour ? `${nour.regularHours.toFixed(1)} hrs (${nour.notes})` : '9h/day Mon-Fri';
      }

      matrix[rowIdx][6] = laborInfo.standardHours;
      matrix[rowIdx][7] = laborInfo.overtimeHours;
      matrix[rowIdx][8] = laborInfo.totalHours;

      // Excluded staff
      const excludedTotal = laborInfo.excludedStaff.reduce((sum, ex) => sum + ex.total, 0);
      matrix[rowIdx][9] = laborInfo.excludedStaff.length > 0 
        ? `${excludedTotal.toFixed(1)} hrs (${laborInfo.excludedStaff.map(ex => `${ex.name}: ${ex.total}`).join(', ')}) [EXCLUDED]`
        : '0.0 hrs (Excluded)';

      if (laborInfo.tabFound) {
        matrix[rowIdx][10] = '✅ Verified from Timesheet';
      } else if (laborInfo.permissionError) {
        matrix[rowIdx][10] = '⚠️ Move Timesheet to Drive Folder to sync';
      } else {
        matrix[rowIdx][10] = 'Calculated via Schedule';
      }
    }

    // Total Row
    const totIdx = totalAuditRow - 1;
    matrix[totIdx][0] = 'TOTAL FLEET DISPATCH LABOR';
    matrix[totIdx][1] = `${monthCount} months`;
    matrix[totIdx][2] = 'Muhammad';
    matrix[totIdx][3] = 'Mariam';
    matrix[totIdx][4] = 'Nourween';
    matrix[totIdx][5] = 'Nour (9h Fixed)';
    matrix[totIdx][6] = `=SUM(G${startAuditRow}:G${endAuditRow})`;
    matrix[totIdx][7] = `=SUM(H${startAuditRow}:H${endAuditRow})`;
    matrix[totIdx][8] = `=SUM(I${startAuditRow}:I${endAuditRow})`;
    matrix[totIdx][9] = 'Excluded Staff Not Counted';
    matrix[totIdx][10] = 'Complete';
  }

  // Section 2: Dispatcher Team Roster Reference
  const sec2TitleRow = totalAuditRow + 3;
  const sec2HeaderRow = sec2TitleRow + 1;
  const sec2StartRow = sec2HeaderRow + 1;

  matrix[sec2TitleRow - 1][0] = '2. DISPATCH TEAM ROSTER & CONFIGURATION RULES';

  matrix[sec2HeaderRow - 1][0] = 'Staff Name';
  matrix[sec2HeaderRow - 1][1] = 'Role & Status';
  matrix[sec2HeaderRow - 1][2] = 'Labor Source';
  matrix[sec2HeaderRow - 1][3] = 'Daily Working Shift';
  matrix[sec2HeaderRow - 1][4] = 'Working Days';
  matrix[sec2HeaderRow - 1][5] = 'Off Days';
  matrix[sec2HeaderRow - 1][6] = 'Calculated in Fleet Ratio?';
  matrix[sec2HeaderRow - 1][7] = 'Notes & Rule Summary';

  const rosterData = [
    ['Muhammad', 'Active Dispatcher', 'Working Time Timesheet', '10.0 hrs/day', 'Mon, Tue, Wed, Thu, Fri', 'Sat, Sun', '✅ YES (Counted)', 'Fetched directly from Timesheet (0 hrs if not working)'],
    ['Mariam', 'Active Dispatcher', 'Working Time Timesheet', '10.0 hrs/day', 'Mon, Tue, Wed, Fri, Sat', 'Thursday, Sun', '✅ YES (Counted)', 'Works Saturday (10h), Thu/Sun OFF (0 hrs if not working)'],
    ['Nourween', 'Active Dispatcher', 'Working Time Timesheet', '10.0 hrs/day', 'Mon, Tue, Wed, Thu, Fri', 'Sat, Sun', '✅ YES (Counted)', 'Fetched directly from Timesheet (0 hrs if not working)'],
    ['Nour', 'Active Dispatcher', 'Fixed Calendar Schedule', '9.0 hrs/day', 'Mon, Tue, Wed, Thu, Fri', 'Sat, Sun', '✅ YES (Counted)', 'Fixed 9h/day Mon-Fri (Not in timesheet)'],
    ['Mohanad (Muhanad)', 'Support Staff', 'Working Time Timesheet', 'Varied', 'As logged', 'As logged', '❌ NO (Excluded)', 'Excluded from dispatch labor overhead'],
    ['Abdulrahman', 'Support Staff', 'Working Time Timesheet', 'Varied', 'As logged', 'As logged', '❌ NO (Excluded)', 'Excluded from dispatch labor overhead']
  ];

  for (let k = 0; k < rosterData.length; k++) {
    const rIdx = sec2StartRow + k - 1;
    for (let c = 0; c < rosterData[k].length; c++) {
      matrix[rIdx][c] = rosterData[k][c];
    }
  }

  // Batch write all data
  sheet.getRange(1, 1, totalRowsNeeded, totalCols).setValues(matrix);

  // Apply Styling
  formatDispatcherAuditSheet(sheet, {
    monthCount: monthCount,
    startAuditRow: startAuditRow,
    endAuditRow: endAuditRow,
    totalAuditRow: totalAuditRow,
    sec2TitleRow: sec2TitleRow,
    sec2HeaderRow: sec2HeaderRow,
    sec2StartRow: sec2StartRow,
    rosterCount: rosterData.length
  });

  return sheet;
}

function formatDispatcherAuditSheet(sheet, cfg) {
  const theme = (typeof CONFIG !== 'undefined' && CONFIG.THEME) ? CONFIG.THEME : {
    headerBg: '#1e293b',
    headerColor: '#ffffff',
    zebraBg: '#f1f5f9'
  };

  // Title Banner
  sheet.getRange('A1:K1').merge()
    .setBackground(theme.headerBg)
    .setFontColor(theme.headerColor)
    .setFontSize(13)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 38);

  // Subtitle
  sheet.getRange('A2:K2').merge()
    .setBackground('#334155')
    .setFontColor('#cbd5e1')
    .setFontSize(9)
    .setFontStyle('italic')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(2, 22);

  // Summary Cards
  sheet.getRange('A3:C3').merge().setValue('TOTAL DISPATCH LABOR HOURS');
  sheet.getRange('A4:C4').merge();
  sheet.getRange('A5:C5').merge().setValue('All Active Months (Regular + Overtime)');
  formatKpiCard(sheet, 'A3:C5', 'A4:C4', '#,##0.0 "hrs"', '#0f172a');

  sheet.getRange('D3:E3').merge().setValue('TOTAL REGULAR HOURS (WT)');
  sheet.getRange('D4:E4').merge();
  sheet.getRange('D5:E5').merge().setValue('Timesheet WT + Nour Fixed Shift');
  formatKpiCard(sheet, 'D3:E5', 'D4:E4', '#,##0.0 "hrs"', '#1e293b');

  sheet.getRange('F3:G3').merge().setValue('TOTAL OVERTIME HOURS (OT)');
  sheet.getRange('F4:G4').merge();
  sheet.getRange('F5:G5').merge().setValue('Timesheet Overtime Logged');
  formatKpiCard(sheet, 'F3:G5', 'F4:G4', '#,##0.0 "hrs"', '#b45309');

  sheet.getRange('H3:K3').merge().setValue('ACTIVE DISPATCH TEAM (4)');
  sheet.getRange('H4:K4').merge().setValue('4 Active Dispatchers');
  sheet.getRange('H5:K5').merge().setValue('Muhammad, Mariam, Nourween, Nour');
  formatKpiCard(sheet, 'H3:K5', 'H4:K4', '@', '#15803d');

  sheet.setRowHeight(3, 20);
  sheet.setRowHeight(4, 32);
  sheet.setRowHeight(5, 18);

  // Section 1 Header
  sheet.getRange('A10:K10').merge()
    .setBackground('#0f172a')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('left');
  sheet.setRowHeight(10, 26);

  // Table Headers
  sheet.getRange('A11:K11')
    .setBackground('#334155')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(9)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
  sheet.setRowHeight(11, 28);

  if (cfg.monthCount > 0) {
    const dataRange = sheet.getRange(cfg.startAuditRow, 1, cfg.monthCount, 11);
    dataRange.setFontSize(9).setVerticalAlignment('middle');

    sheet.getRange(cfg.startAuditRow, 1, cfg.monthCount, 6).setHorizontalAlignment('left');
    sheet.getRange(cfg.startAuditRow, 7, cfg.monthCount, 3).setHorizontalAlignment('right');
    sheet.getRange(cfg.startAuditRow, 10, cfg.monthCount, 2).setHorizontalAlignment('left');

    sheet.getRange(cfg.startAuditRow, 7, cfg.monthCount, 3).setNumberFormat('#,##0.0');

    for (let r = cfg.startAuditRow; r <= cfg.endAuditRow; r++) {
      sheet.getRange(r, 1, 1, 11).setBackground(r % 2 === 0 ? theme.zebraBg : '#ffffff');
      sheet.setRowHeight(r, 24);
    }

    dataRange.setBorder(true, true, true, true, true, true, '#e2e8f0', SpreadsheetApp.BorderStyle.SOLID);

    // Total Row
    const totRange = sheet.getRange(cfg.totalAuditRow, 1, 1, 11);
    totRange.setBackground('#e2e8f0').setFontWeight('bold').setFontSize(9).setVerticalAlignment('middle');
    sheet.getRange(cfg.totalAuditRow, 7, 1, 3).setHorizontalAlignment('right').setNumberFormat('#,##0.0');
    totRange.setBorder(true, true, true, true, false, false, '#475569', SpreadsheetApp.BorderStyle.SOLID);
    sheet.setRowHeight(cfg.totalAuditRow, 26);
  }

  // Section 2 Header
  sheet.getRange(cfg.sec2TitleRow, 1, 1, 11).merge()
    .setBackground('#0f172a')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('left');
  sheet.setRowHeight(cfg.sec2TitleRow, 26);

  sheet.getRange(cfg.sec2HeaderRow, 1, 1, 11)
    .setBackground('#334155')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(9)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
  sheet.setRowHeight(cfg.sec2HeaderRow, 28);

  const rosterRange = sheet.getRange(cfg.sec2StartRow, 1, cfg.rosterCount, 11);
  rosterRange.setFontSize(9).setVerticalAlignment('middle');
  sheet.getRange(cfg.sec2StartRow, 1, cfg.rosterCount, 8).setHorizontalAlignment('left');

  for (let r = cfg.sec2StartRow; r < cfg.sec2StartRow + cfg.rosterCount; r++) {
    sheet.getRange(r, 1, 1, 11).setBackground(r % 2 === 0 ? theme.zebraBg : '#ffffff');
    sheet.setRowHeight(r, 22);
  }
  rosterRange.setBorder(true, true, true, true, true, true, '#e2e8f0', SpreadsheetApp.BorderStyle.SOLID);

  // Column Widths
  sheet.setColumnWidth(1, 150); // Month / Name
  sheet.setColumnWidth(2, 220); // Source / Role
  sheet.setColumnWidth(3, 190); // Muhammad / Source
  sheet.setColumnWidth(4, 210); // Mariam / Shift
  sheet.setColumnWidth(5, 190); // Nourween / Days
  sheet.setColumnWidth(6, 210); // Nour / Off Days
  sheet.setColumnWidth(7, 130); // Regular WT
  sheet.setColumnWidth(8, 130); // Overtime OT
  sheet.setColumnWidth(9, 140); // Total Dispatch
  sheet.setColumnWidth(10, 240); // Excluded Staff
  sheet.setColumnWidth(11, 160); // Verification
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    buildDispatcherAuditSheet
  };
}
