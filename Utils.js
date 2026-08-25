/**
 * @fileoverview Utility and Logger helpers for Fleet Dispatch Efficiency Engine.
 */

/**
 * Appends an entry to the 'Execution Log' sheet for auditing and troubleshooting.
 * 
 * @param {string} action - e.g. "Sync All", "Sync Selected Month", "Recalculate"
 * @param {string} status - "SUCCESS", "WARNING", "ERROR", "INFO"
 * @param {string} message - Summary message
 * @param {string} [details] - Detailed log / warnings
 */
function logExecution(action, status, message, details) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;

    const logSheetName = (typeof CONFIG !== 'undefined' && CONFIG.LOG_SHEET_NAME) ? CONFIG.LOG_SHEET_NAME : 'Execution Log';
    let logSheet = ss.getSheetByName(logSheetName);

    if (!logSheet) {
      logSheet = ss.insertSheet(logSheetName);
      logSheet.getRange('A1:E1').setValues([['Timestamp', 'Action', 'Status', 'Message', 'Details']])
        .setBackground('#1e293b')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
      logSheet.setRowHeight(1, 28);
      logSheet.setColumnWidth(1, 170);
      logSheet.setColumnWidth(2, 160);
      logSheet.setColumnWidth(3, 110);
      logSheet.setColumnWidth(4, 300);
      logSheet.setColumnWidth(5, 350);
      logSheet.setFrozenRows(1);
    }

    const timestamp = new Date().toLocaleString();
    const newRow = [timestamp, action, status, message, details || ''];
    logSheet.appendRow(newRow);

    const lastRow = logSheet.getLastRow();
    const statusCell = logSheet.getRange(lastRow, 3);
    if (status === 'SUCCESS') statusCell.setFontColor('#16a34a').setFontWeight('bold');
    else if (status === 'WARNING') statusCell.setFontColor('#d97706').setFontWeight('bold');
    else if (status === 'ERROR') statusCell.setFontColor('#dc2626').setFontWeight('bold');

  } catch (err) {
    Logger.log(`Failed to write to Execution Log: ${err.message}`);
  }
}

/**
 * Safe toast display for Google Sheets UI.
 */
function showToast(message, title, timeoutSeconds) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title || 'Fleet Dispatch Tools', timeoutSeconds || 5);
  } catch (e) {
    Logger.log(`Toast: [${title}] ${message}`);
  }
}

/**
 * Safe UI alert dialog.
 */
function showAlert(message, title) {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert(title || 'Fleet Dispatch Tools', message, ui.ButtonSet.OK);
  } catch (e) {
    Logger.log(`Alert: [${title}] ${message}`);
  }
}

// Node.js module export for testing
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    logExecution,
    showToast,
    showAlert
  };
}
