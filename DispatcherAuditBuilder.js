/**
 * @fileoverview Dispatcher Hours Audit Sheet Builder.
 * Creates a dedicated '🕒 Dispatcher Hours Audit' tab in Google Sheets
 * to verify and audit exact working hours and overtime fetched from the timesheet
 * (Muhammad, Mariam, Nourween, Mohanad) + fixed schedule (Nour 9h Mon-Fri).
 * Computes individual cumulative totals across all months in the Total row.
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
  const totalCols = 12;
  const matrix = Array.from({ length: totalRowsNeeded }, () => Array(totalCols).fill(''));

  // Row 1: Title Banner
  matrix[0][0] = '🕒 DISPATCHER WORKING HOURS AUDIT & TIMESHEET RECONCILIATION';

  // Row 2: Subtitle
  matrix[1][0] = `Timesheet Source: Muhammad, Mariam, Nourween, Mohanad | Fixed Shift: Nour (9h/day Mon-Fri) | Excluded: Abdulrahman, Fares`;

  // Rows 3-5: Summary Cards
  // Card 1: TOTAL DISPATCH LABOR HOURS (Cols A:C)
  matrix[2][0] = 'TOTAL DISPATCH LABOR HOURS';
  matrix[3][0] = monthCount > 0 ? `=SUM(J${startAuditRow}:J${endAuditRow})` : 0;
  matrix[4][0] = 'All Active Months (Regular + Overtime)';

  // Card 2: TOTAL REGULAR HOURS (Cols D:F)
  matrix[2][3] = 'TOTAL REGULAR HOURS (WT)';
  matrix[3][3] = monthCount > 0 ? `=SUM(H${startAuditRow}:H${endAuditRow})` : 0;
  matrix[4][3] = 'Timesheet WT + Nour Fixed Shift';

  // Card 3: TOTAL OVERTIME HOURS (Cols G:I)
  matrix[2][6] = 'TOTAL OVERTIME HOURS (OT)';
  matrix[3][6] = monthCount > 0 ? `=SUM(I${startAuditRow}:I${endAuditRow})` : 0;
  matrix[4][6] = 'Timesheet Overtime Logged';

  // Card 4: ACTIVE DISPATCHERS (Cols J:L)
  matrix[2][9] = 'ACTIVE DISPATCH TEAM (5)';
  matrix[3][9] = '5 Dispatchers';
  matrix[4][9] = 'Muhammad, Mariam, Nourween, Mohanad, Nour';

  // Row 10: Section 1 Header
  matrix[9][0] = '1. MONTH-BY-MONTH DISPATCHER HOURS AUDIT BREAKDOWN';

  // Row 11: Table Headers
  matrix[10][0] = 'Month & Year';
  matrix[10][1] = 'Labor Data Source';
  matrix[10][2] = 'Muhammad (WT / OT)';
  matrix[10][3] = 'Mariam (WT / OT)';
  matrix[10][4] = 'Nourween (WT / OT)';
  matrix[10][5] = 'Mohanad (WT / OT)';
  matrix[10][6] = 'Nour (Fixed 9h/day)';
  matrix[10][7] = 'Total Regular WT';
  matrix[10][8] = 'Total Overtime OT';
  matrix[10][9] = 'Total Dispatch Hrs';
  matrix[10][10] = 'Excluded Staff Hours';
  matrix[10][11] = 'Verification Status';

  // Accumulators for per-dispatcher totals across all months
  let cumMuhammadWT = 0, cumMuhammadOT = 0;
  let cumMariamWT = 0, cumMariamOT = 0;
  let cumNourweenWT = 0, cumNourweenOT = 0;
  let cumMohanadWT = 0, cumMohanadOT = 0;
  let cumNourHours = 0;

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
      } else if (laborInfo.source === 'no_timesheet') {
        matrix[rowIdx][1] = '❌ No Timesheet (0.0 hrs)';
      } else if (laborInfo.permissionError) {
        matrix[rowIdx][1] = '🔒 Permission Restricted';
      } else {
        matrix[rowIdx][1] = '📅 Calendar Fallback';
      }

      // Find individual staff
      const muhammad = laborInfo.timesheetStaff.find(s => s.name.toLowerCase().includes('muhammad') || s.name.toLowerCase().includes('mohamed'));
      const mariam = laborInfo.timesheetStaff.find(s => s.name.toLowerCase().includes('mariam'));
      const nourween = laborInfo.timesheetStaff.find(s => s.name.toLowerCase().includes('nourween'));
      const mohanad = laborInfo.timesheetStaff.find(s => s.name.toLowerCase().includes('mohanad') || s.name.toLowerCase().includes('muhanad'));
      const nour = laborInfo.fixedStaff.find(s => s.name.toLowerCase() === 'nour');

      if (laborInfo.tabFound) {
        // Muhammad
        if (muhammad && (muhammad.regularHours > 0 || muhammad.overtimeHours > 0)) {
          matrix[rowIdx][2] = `${muhammad.regularHours.toFixed(1)} WT + ${muhammad.overtimeHours.toFixed(1)} OT`;
          cumMuhammadWT += muhammad.regularHours;
          cumMuhammadOT += muhammad.overtimeHours;
        } else {
          matrix[rowIdx][2] = '0.0 hrs (Not Working)';
        }

        // Mariam
        if (mariam && (mariam.regularHours > 0 || mariam.overtimeHours > 0)) {
          matrix[rowIdx][3] = `${mariam.regularHours.toFixed(1)} WT + ${mariam.overtimeHours.toFixed(1)} OT`;
          cumMariamWT += mariam.regularHours;
          cumMariamOT += mariam.overtimeHours;
        } else {
          matrix[rowIdx][3] = '0.0 hrs (Not Working)';
        }

        // Nourween
        if (nourween && (nourween.regularHours > 0 || nourween.overtimeHours > 0)) {
          matrix[rowIdx][4] = `${nourween.regularHours.toFixed(1)} WT + ${nourween.overtimeHours.toFixed(1)} OT`;
          cumNourweenWT += nourween.regularHours;
          cumNourweenOT += nourween.overtimeHours;
        } else {
          matrix[rowIdx][4] = '0.0 hrs (Not Working)';
        }

        // Mohanad
        if (mohanad && (mohanad.regularHours > 0 || mohanad.overtimeHours > 0)) {
          matrix[rowIdx][5] = `${mohanad.regularHours.toFixed(1)} WT + ${mohanad.overtimeHours.toFixed(1)} OT`;
          cumMohanadWT += mohanad.regularHours;
          cumMohanadOT += mohanad.overtimeHours;
        } else {
          matrix[rowIdx][5] = '0.0 hrs (Not Working)';
        }

        // Nour
        if (nour && nour.regularHours > 0) {
          matrix[rowIdx][6] = `${nour.regularHours.toFixed(1)} hrs (${nour.notes})`;
          cumNourHours += nour.regularHours;
        } else {
          matrix[rowIdx][6] = '0.0 hrs (Not Working)';
        }
      } else {
        matrix[rowIdx][2] = '0.0 hrs (No Timesheet)';
        matrix[rowIdx][3] = '0.0 hrs (No Timesheet)';
        matrix[rowIdx][4] = '0.0 hrs (No Timesheet)';
        matrix[rowIdx][5] = '0.0 hrs (No Timesheet)';
        matrix[rowIdx][6] = '0.0 hrs (No Timesheet)';
      }

      matrix[rowIdx][7] = laborInfo.standardHours;
      matrix[rowIdx][8] = laborInfo.overtimeHours;
      matrix[rowIdx][9] = laborInfo.totalHours;

      // Excluded staff (Abdulrahman, Fares)
      const excludedTotal = laborInfo.excludedStaff.reduce((sum, ex) => sum + ex.total, 0);
      matrix[rowIdx][10] = laborInfo.excludedStaff.length > 0 
        ? `${excludedTotal.toFixed(1)} hrs (${laborInfo.excludedStaff.map(ex => `${ex.name}: ${ex.total}`).join(', ')}) [EXCLUDED]`
        : '0.0 hrs (Excluded)';

      if (laborInfo.tabFound) {
        matrix[rowIdx][11] = '✅ Verified from Timesheet';
      } else if (laborInfo.source === 'no_timesheet') {
        matrix[rowIdx][11] = 'No Timesheet Found (0 hrs)';
      } else if (laborInfo.permissionError) {
        matrix[rowIdx][11] = '⚠️ Move Timesheet to Drive Folder to sync';
      } else {
        matrix[rowIdx][11] = 'Calculated via Schedule';
      }
    }

    // Total Row: Calculate individual dispatcher totals!
    const totIdx = totalAuditRow - 1;
    matrix[totIdx][0] = 'TOTAL FLEET DISPATCH LABOR';
    matrix[totIdx][1] = `${monthCount} months`;
    matrix[totIdx][2] = `${cumMuhammadWT.toFixed(1)} WT + ${cumMuhammadOT.toFixed(1)} OT (${(cumMuhammadWT + cumMuhammadOT).toFixed(1)}h)`;
    matrix[totIdx][3] = `${cumMariamWT.toFixed(1)} WT + ${cumMariamOT.toFixed(1)} OT (${(cumMariamWT + cumMariamOT).toFixed(1)}h)`;
    matrix[totIdx][4] = `${cumNourweenWT.toFixed(1)} WT + ${cumNourweenOT.toFixed(1)} OT (${(cumNourweenWT + cumNourweenOT).toFixed(1)}h)`;
    matrix[totIdx][5] = `${cumMohanadWT.toFixed(1)} WT + ${cumMohanadOT.toFixed(1)} OT (${(cumMohanadWT + cumMohanadOT).toFixed(1)}h)`;
    matrix[totIdx][6] = `${cumNourHours.toFixed(1)} hrs`;
    matrix[totIdx][7] = `=SUM(H${startAuditRow}:H${endAuditRow})`;
    matrix[totIdx][8] = `=SUM(I${startAuditRow}:I${endAuditRow})`;
    matrix[totIdx][9] = `=SUM(J${startAuditRow}:J${endAuditRow})`;
    matrix[totIdx][10] = 'Excluded Staff Not Counted';
    matrix[totIdx][11] = 'Complete';
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
    ['Mohanad', 'Active Dispatcher', 'Working Time Timesheet', 'Logged Hours', 'As logged in Timesheet', 'As logged', '✅ YES (Counted)', 'Fetched directly from Timesheet (WT + OT)'],
    ['Nour', 'Active Dispatcher', 'Fixed Calendar Schedule', '9.0 hrs/day', 'Mon, Tue, Wed, Thu, Fri', 'Sat, Sun', '✅ YES (Counted)', 'Fixed 9h/day Mon-Fri (Not in timesheet)'],
    ['Abdulrahman', 'Support Staff', 'Working Time Timesheet', 'Varied', 'As logged', 'As logged', '❌ NO (Excluded)', 'Excluded from dispatch labor overhead'],
    ['Fares', 'Support Staff', 'Working Time Timesheet', 'Varied', 'As logged', 'As logged', '❌ NO (Excluded)', 'Disregarded / Excluded from dispatch labor']
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
  sheet.getRange('A1:L1').merge()
    .setBackground(theme.headerBg)
    .setFontColor(theme.headerColor)
    .setFontSize(13)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 38);

  // Subtitle
  sheet.getRange('A2:L2').merge()
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

  sheet.getRange('D3:F3').merge().setValue('TOTAL REGULAR HOURS (WT)');
  sheet.getRange('D4:F4').merge();
  sheet.getRange('D5:F5').merge().setValue('Timesheet WT + Nour Fixed Shift');
  formatKpiCard(sheet, 'D3:F5', 'D4:F4', '#,##0.0 "hrs"', '#1e293b');

  sheet.getRange('G3:I3').merge().setValue('TOTAL OVERTIME HOURS (OT)');
  sheet.getRange('G4:I4').merge();
  sheet.getRange('G5:I5').merge().setValue('Timesheet Overtime Logged');
  formatKpiCard(sheet, 'G3:I5', 'G4:I4', '#,##0.0 "hrs"', '#b45309');

  sheet.getRange('J3:L3').merge().setValue('ACTIVE DISPATCH TEAM (5)');
  sheet.getRange('J4:L4').merge().setValue('5 Active Dispatchers');
  sheet.getRange('J5:L5').merge().setValue('Muhammad, Mariam, Nourween, Mohanad, Nour');
  formatKpiCard(sheet, 'J3:L5', 'J4:L4', '@', '#15803d');

  sheet.setRowHeight(3, 20);
  sheet.setRowHeight(4, 32);
  sheet.setRowHeight(5, 18);

  // Section 1 Header
  sheet.getRange('A10:L10').merge()
    .setBackground('#0f172a')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('left');
  sheet.setRowHeight(10, 26);

  // Table Headers
  sheet.getRange('A11:L11')
    .setBackground('#334155')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(9)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
  sheet.setRowHeight(11, 28);

  if (cfg.monthCount > 0) {
    const dataRange = sheet.getRange(cfg.startAuditRow, 1, cfg.monthCount, 12);
    dataRange.setFontSize(9).setVerticalAlignment('middle');

    sheet.getRange(cfg.startAuditRow, 1, cfg.monthCount, 7).setHorizontalAlignment('left');
    sheet.getRange(cfg.startAuditRow, 8, cfg.monthCount, 3).setHorizontalAlignment('right');
    sheet.getRange(cfg.startAuditRow, 11, cfg.monthCount, 2).setHorizontalAlignment('left');

    sheet.getRange(cfg.startAuditRow, 8, cfg.monthCount, 3).setNumberFormat('#,##0.0');

    for (let r = cfg.startAuditRow; r <= cfg.endAuditRow; r++) {
      sheet.getRange(r, 1, 1, 12).setBackground(r % 2 === 0 ? theme.zebraBg : '#ffffff');
      sheet.setRowHeight(r, 24);
    }

    dataRange.setBorder(true, true, true, true, true, true, '#e2e8f0', SpreadsheetApp.BorderStyle.SOLID);

    // Total Row
    const totRange = sheet.getRange(cfg.totalAuditRow, 1, 1, 12);
    totRange.setBackground('#e2e8f0').setFontWeight('bold').setFontSize(9).setVerticalAlignment('middle');
    sheet.getRange(cfg.totalAuditRow, 8, 1, 3).setHorizontalAlignment('right').setNumberFormat('#,##0.0');
    totRange.setBorder(true, true, true, true, false, false, '#475569', SpreadsheetApp.BorderStyle.SOLID);
    sheet.setRowHeight(cfg.totalAuditRow, 26);
  }

  // Section 2 Header
  sheet.getRange(cfg.sec2TitleRow, 1, 1, 12).merge()
    .setBackground('#0f172a')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('left');
  sheet.setRowHeight(cfg.sec2TitleRow, 26);

  sheet.getRange(cfg.sec2HeaderRow, 1, 1, 12)
    .setBackground('#334155')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(9)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
  sheet.setRowHeight(cfg.sec2HeaderRow, 28);

  const rosterRange = sheet.getRange(cfg.sec2StartRow, 1, cfg.rosterCount, 12);
  rosterRange.setFontSize(9).setVerticalAlignment('middle');
  sheet.getRange(cfg.sec2StartRow, 1, cfg.rosterCount, 8).setHorizontalAlignment('left');

  for (let r = cfg.sec2StartRow; r < cfg.sec2StartRow + cfg.rosterCount; r++) {
    sheet.getRange(r, 1, 1, 12).setBackground(r % 2 === 0 ? theme.zebraBg : '#ffffff');
    sheet.setRowHeight(r, 22);
  }
  rosterRange.setBorder(true, true, true, true, true, true, '#e2e8f0', SpreadsheetApp.BorderStyle.SOLID);

  // Column Widths
  sheet.setColumnWidth(1, 150); // Month / Name
  sheet.setColumnWidth(2, 210); // Source / Role
  sheet.setColumnWidth(3, 190); // Muhammad / Source
  sheet.setColumnWidth(4, 190); // Mariam / Shift
  sheet.setColumnWidth(5, 190); // Nourween / Days
  sheet.setColumnWidth(6, 190); // Mohanad / Off Days
  sheet.setColumnWidth(7, 180); // Nour / Fixed
  sheet.setColumnWidth(8, 130); // Regular WT
  sheet.setColumnWidth(9, 130); // Overtime OT
  sheet.setColumnWidth(10, 140); // Total Dispatch
  sheet.setColumnWidth(11, 240); // Excluded Staff
  sheet.setColumnWidth(12, 160); // Verification
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    buildDispatcherAuditSheet
  };
}
