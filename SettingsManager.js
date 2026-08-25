/**
 * @fileoverview Settings Manager for Fleet Dispatch Efficiency Engine.
 * Creates and reads the 'Settings' sheet tab to allow configuring
 * Folder IDs, shift schedules, and defaults directly inside Google Sheets.
 */

/**
 * Gets the effective Google Drive Folder ID (from Settings tab or CONFIG constant).
 * @returns {string} Folder ID
 */
function getEffectiveDriveFolderId() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return CONFIG.DRIVE_FOLDER_ID;

  const settingsSheet = ss.getSheetByName(CONFIG.SETTINGS_SHEET_NAME || 'Settings');
  if (settingsSheet) {
    const val = settingsSheet.getRange('B4').getValue();
    if (val && String(val).trim().length > 0) {
      // If user pasted a full URL, extract the ID
      const strVal = String(val).trim();
      const match = strVal.match(/folders\/([a-zA-Z0-9_-]+)/);
      return match ? match[1] : strVal;
    }
  }

  return CONFIG.DRIVE_FOLDER_ID;
}

/**
 * Initializes or updates the 'Settings' sheet tab in the active spreadsheet.
 */
function initSettingsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheetName = CONFIG.SETTINGS_SHEET_NAME || 'Settings';
  let sheet = ss.getSheetByName(settingsSheetName);

  if (!sheet) {
    sheet = ss.insertSheet(settingsSheetName);
  }

  sheet.clear();
  sheet.clearConditionalFormatRules();

  // Header Banner
  sheet.getRange('A1:D1').merge()
    .setValue('⚙️ FLEET DISPATCH EFFICIENCY ENGINE — CONFIGURATION & SETTINGS')
    .setBackground('#1e293b')
    .setFontColor('#ffffff')
    .setFontSize(12)
    .setFontWeight('bold')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 36);

  // Section 1: Drive Configuration
  sheet.getRange('A3:D3').merge()
    .setValue('1. GOOGLE DRIVE SOURCE FOLDER')
    .setBackground('#334155')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(10);
  sheet.setRowHeight(3, 24);

  sheet.getRange('A4').setValue('Drive Folder ID / URL:').setFontWeight('bold');
  sheet.getRange('B4:D4').merge().setValue(CONFIG.DRIVE_FOLDER_URL);
  sheet.getRange('B4').setNote('Enter Google Drive Folder URL or Folder ID where monthly balance files (and 2025/2026 subfolders) are located.');

  sheet.getRange('A5').setValue('Target Sheet Tab:').setFontWeight('bold');
  sheet.getRange('B5').setValue(CONFIG.SOURCE_SHEET_NAME);

  sheet.getRange('A6').setValue('File Naming Pattern:').setFontWeight('bold');
  sheet.getRange('B6:D6').merge().setValue('[Month] - Drivers Daily Balance');

  // Section 2: Dispatcher Shift Schedule
  sheet.getRange('A8:D8').merge()
    .setValue('2. DISPATCHER SHIFTS & WEEKLY SCHEDULE')
    .setBackground('#334155')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(10);
  sheet.setRowHeight(8, 24);

  const scheduleHeaders = ['Dispatcher', 'Daily Shift Hours', 'Working Days', 'Off Days'];
  for (let c = 0; c < 4; c++) {
    sheet.getRange(9, c + 1).setValue(scheduleHeaders[c])
      .setBackground('#475569')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
  }

  const shiftData = [
    ['Dispatcher 1', '10.0 hrs', 'Mon, Tue, Wed, Thu, Fri', 'Sat, Sun'],
    ['Dispatcher 2', '10.0 hrs', 'Mon, Tue, Wed, Thu, Fri', 'Sat, Sun'],
    ['Dispatcher 3', '10.0 hrs', 'Mon, Tue, Wed, Thu, Fri', 'Sat, Sun'],
    ['Dispatcher 4', '9.0 hrs',  'Mon, Tue, Wed, Thu, Fri', 'Sat, Sun']
  ];

  for (let r = 0; r < shiftData.length; r++) {
    for (let c = 0; c < 4; c++) {
      sheet.getRange(10 + r, c + 1).setValue(shiftData[r][c])
        .setHorizontalAlignment(c === 0 ? 'left' : 'center');
    }
  }

  // Section 3: Daily Fleet Labor Breakdown
  sheet.getRange('A15:D15').merge()
    .setValue('3. FLEET DISPATCH LABOR HOURS PER DAY OF WEEK')
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
    ['Monday', '4 (Disp 1, 2, 3, 4)', '39.0 hrs', '10 + 10 + 10 + 9'],
    ['Tuesday', '4 (Disp 1, 2, 3, 4)', '39.0 hrs', '10 + 10 + 10 + 9'],
    ['Wednesday', '4 (Disp 1, 2, 3, 4)', '39.0 hrs', '10 + 10 + 10 + 9'],
    ['Thursday', '4 (Disp 1, 2, 3, 4)', '39.0 hrs', '10 + 10 + 10 + 9'],
    ['Friday', '4 (Disp 1, 2, 3, 4)', '39.0 hrs', '10 + 10 + 10 + 9'],
    ['Saturday', '0', '0.0 hrs', 'Fleet OFF'],
    ['Sunday', '0', '0.0 hrs', 'Fleet OFF']
  ];

  for (let r = 0; r < dailyTotals.length; r++) {
    for (let c = 0; c < 4; c++) {
      sheet.getRange(17 + r, c + 1).setValue(dailyTotals[r][c])
        .setHorizontalAlignment(c === 0 ? 'left' : 'center');
    }
  }

  // Formatting borders & widths
  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 200);

  SpreadsheetApp.getActiveSpreadsheet().toast('Settings tab initialized successfully!', 'Fleet Tools', 4);
}

// Node.js module export for testing
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    getEffectiveDriveFolderId,
    initSettingsSheet
  };
}
