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
    .addItem('⚡ Recalculate All Dispatch Ratios', 'menuRecalculateRatios')
    .addSeparator()
    .addItem('⚙️ Open / Reset Settings Tab', 'initSettingsSheet')
    .addItem('🧪 Generate Sample Data in Drive (Demo)', 'menuGenerateSampleData')
    .addItem('📖 User Guide & Metric Formulas', 'menuShowUserGuide')
    .addToUi();
}

/**
 * Menu Handler: Syncs all monthly balance spreadsheets found in the Drive folder & subfolders.
 */
function menuSyncAllMonthlyFiles() {
  const ui = SpreadsheetApp.getUi();
  showToast('Scanning Drive folder for monthly balance spreadsheets...', 'Sync in Progress', 10);

  try {
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

    for (const item of matchedFiles) {
      try {
        const monthData = extractDriverDataFromSummary(item.fileId, item.monthIndex, item.year);
        buildMonthlyAnalysisSheet(monthData);

        let warnMsg = '';
        if (monthData.warnings && monthData.warnings.length > 0) {
          warningCount++;
          warnMsg = monthData.warnings.join('; ');
        } else {
          successCount++;
        }

        summaryLog.push(`• ${item.monthName} ${item.year} (${item.folderName}): ${monthData.driverCount} drivers, ${monthData.totalDriverHours.toFixed(1)} hrs ${warnMsg ? '⚠️ ' + warnMsg : '✅'}`);
      } catch (fileErr) {
        summaryLog.push(`• ${item.fileName} (${item.folderName}): ❌ ERROR - ${fileErr.message}`);
        logExecution('Sync File', 'ERROR', `Failed to process ${item.fileName}`, fileErr.message);
      }
    }

    const logStatus = warningCount > 0 ? 'WARNING' : 'SUCCESS';
    logExecution('Sync All', logStatus, `Processed ${matchedFiles.length} file(s) from Drive.`, summaryLog.join('\n'));

    showAlert(
      `Sync Complete!\n\nProcessed ${matchedFiles.length} monthly spreadsheet(s):\n\n` +
      summaryLog.join('\n') +
      `\n\nEach monthly analysis tab has been created/updated with live KPI cards and driver ratios.`,
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
    const sheet = buildMonthlyAnalysisSheet(monthData);

    SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);

    logExecution('Sync Selected Month', 'SUCCESS', `Successfully synced ${targetFile.monthName} ${targetFile.year}`, `${monthData.driverCount} drivers`);
    showAlert(
      `Successfully synced ${targetFile.monthName} ${targetFile.year}!\n\n` +
      `• Drivers Loaded: ${monthData.driverCount}\n` +
      `• Fleet Active Hours: ${monthData.totalDriverHours.toFixed(1)} hrs\n` +
      `• Completed Trips: ${monthData.totalTrips}\n` +
      `• Tab: "${sheet.getName()}"`,
      'Month Synced'
    );

  } catch (err) {
    logExecution('Sync Selected Month', 'ERROR', `Error syncing ${inputMonth}`, err.message);
    showAlert(`Error syncing month: ${err.message}`, 'Sync Error');
  }
}

/**
 * Menu Handler: Recalculates dispatch ratios on all existing '[Month] - Analysis' tabs.
 */
function menuRecalculateRatios() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let updatedCount = 0;

  for (const sheet of sheets) {
    const name = sheet.getName();
    if (name.includes('- Analysis')) {
      // Re-trigger calculation by checking formulas and updating formatting if necessary
      const monthPart = name.replace('- Analysis', '').trim();
      const parsed = parseMonthAndYear(monthPart);
      const schedule = calculateMonthlyDispatchHours(parsed.monthIndex, parsed.year);

      // Refresh standard dispatch hours in cell C4 if present
      const labelC3 = sheet.getRange('C3').getValue();
      if (labelC3 && String(labelC3).includes('STANDARD DISPATCH')) {
        sheet.getRange('C4').setValue(schedule.totalHours);
        updatedCount++;
      }
    }
  }

  logExecution('Recalculate Ratios', 'SUCCESS', `Recalculated dispatch labor on ${updatedCount} analysis tab(s).`);
  showAlert(`Successfully recalculated dispatch schedules across ${updatedCount} analysis sheet(s).`, 'Recalculation Complete');
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

/**
 * Menu Handler: Displays an informational guide about the calculations.
 */
function menuShowUserGuide() {
  const guideText = 
    `📊 FLEET DISPATCH EFFICIENCY ENGINE — GUIDE\n\n` +
    `1. Calendar Shift Schedule (4 Dispatchers):\n` +
    `   • Monday – Friday: 39 hrs/day (3 × 10h + 1 × 9h)\n` +
    `   • Saturday & Sunday: 0 hrs/day (OFF)\n\n` +
    `2. Key Metrics & Formulas:\n` +
    `   • Total Dispatch Hours = Standard Dispatch Hours + Overtime\n` +
    `   • Fleet Dispatch Ratio = Total Dispatch Hours / Total Driver Hours\n` +
    `   • Support Mins / Road Hr = Fleet Dispatch Ratio × 60\n` +
    `   • Driver Leverage Ratio = Total Driver Hours / Total Dispatch Hours\n` +
    `   • Allocated Dispatch Hours = Driver Active Hours × Fleet Dispatch Ratio\n` +
    `   • Trips Supported / Disp Hr = Completed Trips / Allocated Dispatch Hours\n\n` +
    `3. Interactive Features:\n` +
    `   • You can edit cell E4 (Overtime / Adj) in any Analysis tab to instantly recalculate all ratios in real time!`;

  showAlert(guideText, 'Fleet Dispatch Engine Guide');
}
