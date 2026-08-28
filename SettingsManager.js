/**
 * @fileoverview Settings Manager for Fleet Dispatch Efficiency Engine.
 * Creates and reads the 'Settings' sheet tab to allow configuring
 * Folder IDs, shift schedules, and defaults directly inside Google Sheets.
 */

function getEffectiveDriveFolderId() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return CONFIG.DRIVE_FOLDER_ID;

  const settingsSheet = ss.getSheetByName(CONFIG.SETTINGS_SHEET_NAME || 'Settings');
  if (settingsSheet) {
    const val = settingsSheet.getRange('B4').getValue();
    if (val && String(val).trim().length > 0) {
      const strVal = String(val).trim();
      const match = strVal.match(/folders\/([a-zA-Z0-9_-]+)/);
      return match ? match[1] : strVal;
    }
  }

  return CONFIG.DRIVE_FOLDER_ID;
}

function initSettingsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheetName = CONFIG.SETTINGS_SHEET_NAME || 'Settings';
  let sheet = ss.getSheetByName(settingsSheetName);

  if (!sheet) {
    sheet = ss.insertSheet(settingsSheetName);
  }

  sheet.clear();
  sheet.clearConditionalFormatRules();

  // Banner
  sheet.getRange('A1:D1').merge()
    .setValue('⚙️ FLEET DISPATCH EFFICIENCY ENGINE — CONFIGURATION & SETTINGS')
    .setBackground('#1e293b')
    .setFontColor('#ffffff')
    .setFontSize(12)
    .setFontWeight('bold')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 36);

  // Section 1: Source Files & Timesheet
  sheet.getRange('A3:D3').merge()
    .setValue('1. GOOGLE DRIVE SOURCE & DISPATCHER TIMESHEET')
    .setBackground('#334155')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(10);
  sheet.setRowHeight(3, 24);

  sheet.getRange('A4').setValue('Balance Drive Folder:').setFontWeight('bold');
  sheet.getRange('B4:D4').merge().setValue(CONFIG.DRIVE_FOLDER_URL);

  sheet.getRange('A5').setValue('Working Time Sheet:').setFontWeight('bold');
  sheet.getRange('B5:D5').merge().setValue(CONFIG.WORKING_TIME_SPREADSHEET_URL);

  sheet.getRange('A6').setValue('Active Dispatchers:').setFontWeight('bold');
  sheet.getRange('B6:D6').merge().setValue('Muhammad, Mariam (Works Sat, Thu OFF), Nour (Nourween)');

  sheet.getRange('A7').setValue('Excluded Staff:').setFontWeight('bold');
  sheet.getRange('B7:D7').merge().setValue('Mohanad, Abdulrahman');

  // Section 2: Dispatcher Shift Schedule
  sheet.getRange('A9:D9').merge()
    .setValue('2. DISPATCHER SHIFTS & WEEKLY SCHEDULE (CALENDAR BASE)')
    .setBackground('#334155')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(10);
  sheet.setRowHeight(9, 24);

  const scheduleHeaders = ['Dispatcher', 'Daily Shift Hours', 'Working Days', 'Off Days'];
  for (let c = 0; c < 4; c++) {
    sheet.getRange(10, c + 1).setValue(scheduleHeaders[c])
      .setBackground('#475569')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
  }

  const shiftData = [
    ['Muhammad', '10.0 hrs', 'Mon, Tue, Wed, Thu, Fri', 'Sat, Sun'],
    ['Mariam', '10.0 hrs', 'Mon, Tue, Wed, Fri, Sat', 'Thursday, Sun'],
    ['Nour (Nourween)', '9.0 hrs', 'Mon, Tue, Wed, Thu, Fri', 'Sat, Sun']
  ];

  for (let r = 0; r < shiftData.length; r++) {
    for (let c = 0; c < 4; c++) {
      sheet.getRange(11 + r, c + 1).setValue(shiftData[r][c])
        .setHorizontalAlignment(c === 0 ? 'left' : 'center');
    }
  }

  // Section 3: Daily Fleet Labor Breakdown
  sheet.getRange('A15:D15').merge()
    .setValue('3. DAILY DISPATCH LABOR HOURS PER DAY OF WEEK')
    .setBackground('#334155')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(10);
  sheet.setRowHeight(15, 24);

  const daysHeader = ['Day of Week', 'Active Dispatchers', 'Daily Total Hours', 'Notes'];
  for (let c = 0; c < 4; c++) {
    sheet.getRange(16, c + 1).setValue(daysHeader[c])
      .setBackground('#475569')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
  }

  const dailyTotals = [
    ['Monday', '3 (Muhammad, Mariam, Nour)', '29.0 hrs', '10 + 10 + 9'],
    ['Tuesday', '3 (Muhammad, Mariam, Nour)', '29.0 hrs', '10 + 10 + 9'],
    ['Wednesday', '3 (Muhammad, Mariam, Nour)', '29.0 hrs', '10 + 10 + 9'],
    ['Thursday', '2 (Muhammad, Nour)', '19.0 hrs', '10 + 9 (Mariam OFF)'],
    ['Friday', '3 (Muhammad, Mariam, Nour)', '29.0 hrs', '10 + 10 + 9'],
    ['Saturday', '1 (Mariam)', '10.0 hrs', 'Mariam 10h (Muhammad & Nour OFF)'],
    ['Sunday', '0', '0.0 hrs', 'Fleet OFF']
  ];

  for (let r = 0; r < dailyTotals.length; r++) {
    for (let c = 0; c < 4; c++) {
      sheet.getRange(17 + r, c + 1).setValue(dailyTotals[r][c])
        .setHorizontalAlignment(c === 0 ? 'left' : 'center');
    }
  }

  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 240);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 200);

  SpreadsheetApp.getActiveSpreadsheet().toast('Settings tab initialized successfully!', 'Fleet Tools', 4);
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    getEffectiveDriveFolderId,
    initSettingsSheet
  };
}
