/**
 * @fileoverview Custom UI Menu and top-level orchestrator functions.
 * Adds the 'Fleet Dispatch Tools' menu to Google Sheets and executes sync tasks.
 */

/**
 * Standard trigger that runs when the spreadsheet is opened.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚗 Fleet Dispatch Tools')
    .addItem('🔄 Sync All Monthly Files from Drive', 'menuSyncAllMonthlyFiles')
    .addItem('📅 Sync Selected Month', 'menuSyncSelectedMonth')
    .addSeparator()
    .addItem('📊 Open / Refresh "Master Summary" Tab', 'menuBuildMasterSummaryTab')
    .addItem('🕒 Open / Refresh "Dispatcher Hours Audit" Tab', 'menuBuildDispatcherAuditTab')
    .addItem('⚡ Recalculate All Dispatch Ratios', 'menuRecalculateRatios')
    .addSeparator()
    .addItem('📖 Open / Refresh "Read Me & Guide" Tab', 'menuBuildReadMeTab')
    .addItem('⚙️ Open / Reset Settings Tab', 'initSettingsSheet')
    .addItem('🧪 Generate Sample Data in Drive (Demo)', 'menuGenerateSampleData')
    .addToUi();
}

/**
 * Menu Handler: Builds or refreshes the Dispatcher Hours Audit tab.
 */
function menuBuildDispatcherAuditTab() {
  showToast('Building Dispatcher Working Hours Audit...', 'Dispatcher Audit', 6);
  const sheet = buildDispatcherAuditSheet();
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  showToast('Dispatcher Hours Audit refreshed successfully!', 'Audit Ready', 5);
}

/**
 * Menu Handler: Builds or refreshes the Master Summary tab.
 */
function menuBuildMasterSummaryTab() {
  showToast('Building Master Summary across all months...', 'Master Summary', 6);
  const sheet = buildMasterSummarySheet();
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  showToast('Master Summary tab refreshed successfully!', 'Master Summary Ready', 5);
}

/**
 * Menu Handler: Builds or refreshes the Read Me & Guide tab.
 */
function menuBuildReadMeTab() {
  const sheet = buildReadMeSheet();
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  showToast('Read Me & Guide tab refreshed successfully!', 'Guide Ready', 5);
}

/**
 * Cleans up legacy sheet tab names by removing '- Analysis' suffix.
 */
function cleanupLegacyTabNames(ss) {
  const sheets = ss.getSheets();
  for (const s of sheets) {
    const oldName = s.getName();
    if (oldName.includes('- Analysis')) {
      const newName = oldName.replace('- Analysis', '').trim();
      const existing = ss.getSheetByName(newName);
      if (!existing) {
        s.setName(newName);
      }
    }
  }
}

/**
 * Menu Handler: Syncs all monthly balance spreadsheets found in the Drive folder & subfolders.
 */
function menuSyncAllMonthlyFiles() {
  showToast('Scanning Drive folder for monthly balance spreadsheets...', 'Sync in Progress', 10);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Ensure Read Me tab exists
    buildReadMeSheet(ss);

    // Clean up any legacy '- Analysis' tab names
    cleanupLegacyTabNames(ss);

    const matchedFiles = findMonthlyBalanceFiles();

    if (!matchedFiles || matchedFiles.length === 0) {
      const folderId = getEffectiveDriveFolderId();
      const msg = `No files matching "[Month] - Drivers Daily Balance" were found in Drive folder ID: ${folderId}.\n\nPlease ensure your files are in the folder or 2025/2026 subfolders.`;
      logExecution('Sync All', 'WARNING', 'No matching monthly balance files found in Drive.', `Folder: ${folderId}`);
      showAlert(msg, 'No Files Found');
      return;
    }

    let successCount = 0;
    let warningCount = 0;
    const summaryLog = [];
    const allMonthsData = [];

    for (const item of matchedFiles) {
      try {
        const monthData = extractDriverDataFromSummary(item.fileId, item.monthIndex, item.year);
        buildMonthlyAnalysisSheet(monthData, ss);
        allMonthsData.push(monthData);

        let warnMsg = '';
        if (monthData.warnings && monthData.warnings.length > 0) {
          warningCount++;
          warnMsg = monthData.warnings.join('; ');
        } else {
          successCount++;
        }

        const noteZero = (monthData.totalDriverHours === 0) ? ' [0 driver hrs - Excluded from Master Summary]' : '';
        summaryLog.push(`• ${item.monthName} ${item.year} (${item.folderName}): ${monthData.driverCount} unique drivers, ${monthData.totalDriverHours.toFixed(1)} hrs ${warnMsg ? '⚠️ ' + warnMsg : '✅'}${noteZero}`);
      } catch (fileErr) {
        summaryLog.push(`• ${item.fileName} (${item.folderName}): ❌ ERROR - ${fileErr.message}`);
        logExecution('Sync File', 'ERROR', `Failed to process ${item.fileName}`, fileErr.message);
      }
    }

    // Build/Refresh Master Summary tab and Dispatcher Audit tab
    if (allMonthsData.length > 0) {
      buildMasterSummarySheet(allMonthsData, ss);
      buildDispatcherAuditSheet(allMonthsData, ss);
    }

    const logStatus = warningCount > 0 ? 'WARNING' : 'SUCCESS';
    logExecution('Sync All', logStatus, `Processed ${matchedFiles.length} file(s) from Drive.`, summaryLog.join('\n'));

    showAlert(
      `Sync Complete!\n\nProcessed ${matchedFiles.length} monthly spreadsheet(s):\n\n` +
      summaryLog.join('\n') +
      `\n\n✅ Clean monthly tabs generated (e.g. "August 2026", "July 2026").\n✅ "📊 Master Summary" tab updated.\n✅ "🕒 Dispatcher Hours Audit" tab created.\n✅ "📖 Read Me & Guide" tab available.`,
      'Sync Summary'
    );

  } catch (err) {
    logExecution('Sync All', 'ERROR', 'Fatal error during Drive sync', err.message);
    showAlert(`Error during sync: ${err.message}`, 'Sync Failed');
  }
}

/**
 * Menu Handler: Syncs a single specific month chosen by the user.
 */
function menuSyncSelectedMonth() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.prompt(
    'Sync Selected Month',
    'Enter the Month & Year to sync (e.g. "August 2026", "September 2025", or "August"):',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const inputMonth = response.getResponseText().trim();
  if (!inputMonth) {
    showAlert('Please enter a valid month name.', 'Invalid Input');
    return;
  }

  const parsed = parseMonthAndYear(inputMonth);
  showToast(`Searching for ${parsed.monthName} ${parsed.year} balance file...`, 'Syncing Month', 8);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    buildReadMeSheet(ss);

    const matchedFiles = findMonthlyBalanceFiles();
    const targetFile = matchedFiles.find(f => 
      f.monthIndex === parsed.monthIndex && (f.year === parsed.year || matchedFiles.length === 1)
    );

    if (!targetFile) {
      showAlert(
        `Could not find a balance spreadsheet for "${parsed.monthName} ${parsed.year}".\n\nPlease verify that the file exists in your Drive folder (or 2025/2026 subfolders) and is named "${parsed.monthName} - Drivers Daily Balance".`,
        'File Not Found'
      );
      return;
    }

    const monthData = extractDriverDataFromSummary(targetFile.fileId, targetFile.monthIndex, targetFile.year);
    const sheet = buildMonthlyAnalysisSheet(monthData, ss);

    // Refresh Master Summary and Dispatcher Audit
    buildMasterSummarySheet(null, ss);
    buildDispatcherAuditSheet(null, ss);

    SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);

    logExecution('Sync Selected Month', 'SUCCESS', `Successfully synced ${targetFile.monthName} ${targetFile.year}`, `${monthData.driverCount} unique drivers`);
    showAlert(
      `Successfully synced ${targetFile.monthName} ${targetFile.year}!\n\n` +
      `• Unique Drivers: ${monthData.driverCount}\n` +
      `• Fleet Active Hours: ${monthData.totalDriverHours.toFixed(1)} hrs\n` +
      `• Completed Trips: ${monthData.totalTrips}\n` +
      `• Tab: "${sheet.getName()}"\n` +
      `• Master Summary & Dispatcher Audit updated!`,
      'Month Synced'
    );

  } catch (err) {
    logExecution('Sync Selected Month', 'ERROR', `Error syncing ${inputMonth}`, err.message);
    showAlert(`Error syncing month: ${err.message}`, 'Sync Error');
  }
}

/**
 * Menu Handler: Recalculates dispatch ratios on all existing monthly tabs.
 */
function menuRecalculateRatios() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let updatedCount = 0;
  const systemNames = ['master summary', 'read me', 'user guide', 'settings', 'execution log', 'dispatcher hours audit', 'sheet1'];

  for (const sheet of sheets) {
    const name = sheet.getName();
    const lower = name.toLowerCase();
    const isSystem = systemNames.some(sys => lower.includes(sys));
    if (isSystem) continue;

    const isMonth = MONTH_NAMES.some(m => lower.includes(m.toLowerCase()));
    if (isMonth) {
      const cleanName = name.replace('- Analysis', '').trim();
      const parsed = parseMonthAndYear(cleanName);
      const laborInfo = fetchDispatcherHoursFromTimesheet(parsed.monthIndex, parsed.year);

      const labelC3 = sheet.getRange('C3').getValue();
      if (labelC3 && String(labelC3).includes('STANDARD DISPATCH')) {
        sheet.getRange('C4').setValue(laborInfo.standardHours);
        if (laborInfo.overtimeHours > 0) {
          sheet.getRange('E4').setValue(laborInfo.overtimeHours);
        }
        updatedCount++;
      }
    }
  }

  buildMasterSummarySheet(null, ss);
  buildDispatcherAuditSheet(null, ss);

  logExecution('Recalculate Ratios', 'SUCCESS', `Recalculated dispatch labor on ${updatedCount} monthly tab(s).`);
  showAlert(`Successfully recalculated dispatch schedules across ${updatedCount} monthly sheet(s) and refreshed Master Summary & Dispatcher Audit.`, 'Recalculation Complete');
}

/**
 * Menu Handler: Generates a sample monthly balance spreadsheet for testing.
 */
function menuGenerateSampleData() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.alert(
    'Generate Sample Monthly Balance Sheet',
    'Would you like to generate a sample "August 2026 - Drivers Daily Balance" spreadsheet in your configured Google Drive folder for testing?',
    ui.ButtonSet.YES_NO
  );

  if (res !== ui.Button.YES) return;

  try {
    showToast('Creating sample spreadsheet in Drive...', 'Generating Demo Data', 8);
    const newFile = createSampleMonthlySpreadsheet('August', 2026);
    showAlert(`Sample file created successfully in Google Drive:\n\nFile Name: "${newFile.getName()}"\nURL: ${newFile.getUrl()}\n\nYou can now run "Sync All Monthly Files from Drive" to test!`, 'Demo Data Ready');
  } catch (err) {
    showAlert(`Failed to generate sample data: ${err.message}`, 'Error');
  }
}
