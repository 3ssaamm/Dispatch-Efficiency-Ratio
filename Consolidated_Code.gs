/**
 * ============================================================================
 * GOOGLE APPS SCRIPT & GOOGLE SHEETS FLEET DISPATCH EFFICIENCY ENGINE
 * ============================================================================
 * Author: Antigravity AI
 * Folder ID: 1iDd2ME2b6coX8GO6Dkom5u9jldkRSwVg
 * Folder URL: https://drive.google.com/drive/folders/1iDd2ME2b6coX8GO6Dkom5u9jldkRSwVg?usp=sharing
 * ============================================================================
 */

// ============================================================================
// 1. CONFIGURATION & CONSTANTS
// ============================================================================

const CONFIG = {
  // Master Google Drive Folder containing monthly balance files & subfolders (2025, 2026)
  DRIVE_FOLDER_ID: '1iDd2ME2b6coX8GO6Dkom5u9jldkRSwVg',
  DRIVE_FOLDER_URL: 'https://drive.google.com/drive/folders/1iDd2ME2b6coX8GO6Dkom5u9jldkRSwVg?usp=sharing',

  // Target sheet tab name inside each monthly workbook
  SOURCE_SHEET_NAME: 'Summary',

  // File naming regex / pattern (Matches "August - Drivers Daily Balance", "August 2026 - Drivers Daily Balance", etc.)
  FILE_NAME_PATTERN: /^(.*)\s*-\s*Drivers Daily Balance$/i,

  // Dispatcher shift schedules and daily labor hours (4 Dispatchers)
  // - Disp 1: 10h Mon-Fri (Sat/Sun OFF)
  // - Disp 2: 10h Mon-Fri (Sat/Sun OFF)
  // - Disp 3: 10h Mon-Fri (Sat/Sun OFF)
  // - Disp 4: 9h Mon-Fri (Sat/Sun OFF)
  // Total Weekday = 10 + 10 + 10 + 9 = 39 hrs/day. Sat & Sun = 0 hrs/day.
  DISPATCH_HOURS_BY_DAY_OF_WEEK: {
    0: 0,   // Sunday: Fleet-wide OFF
    1: 39,  // Monday: 39 hrs
    2: 39,  // Tuesday: 39 hrs
    3: 39,  // Wednesday: 39 hrs
    4: 39,  // Thursday: 39 hrs
    5: 39,  // Friday: 39 hrs
    6: 0    // Saturday: OFF
  },

  // Detailed shift rosters for auditing & settings display
  DISPATCHERS: [
    { id: 1, name: 'Dispatcher 1', dailyHours: 10, workDays: [1, 2, 3, 4, 5], offDays: [0, 6] },
    { id: 2, name: 'Dispatcher 2', dailyHours: 10, workDays: [1, 2, 3, 4, 5], offDays: [0, 6] },
    { id: 3, name: 'Dispatcher 3', dailyHours: 10, workDays: [1, 2, 3, 4, 5], offDays: [0, 6] },
    { id: 4, name: 'Dispatcher 4', dailyHours: 9,  workDays: [1, 2, 3, 4, 5], offDays: [0, 6] }
  ],

  // Column header aliases for dynamic column lookup in the 'Summary' tab
  HEADER_ALIASES: {
    driver: ['driver', 'driver name', 'driver id', 'name', 'drivername', 'pilot', 'chauffeur'],
    hours: ['hours', 'active hours', 'driver hours', 'total hours', 'hrs', 'active hrs', 'working hours', 'road hours'],
    trips: ['trips', 'completed trips', 'total trips', 'trip count', 'rides', 'completed rides'],
    credit: ['credit', 'credits', 'credit card', 'cc', 'cc amount'],
    cash: ['cash', 'cash amount', 'cash collected'],
    balance: ['balance', 'net balance', 'total balance', 'final balance'],
    noShow: ['no show', 'no-show', 'noshow', 'no shows', 'no_show', 'cancelled']
  },

  // UI Theme Styling for Output Sheets
  THEME: {
    headerBg: '#1e293b',       // Deep slate navy
    headerColor: '#ffffff',    // White text
    kpiCardBg: '#f8fafc',      // Light slate
    kpiCardBorder: '#cbd5e1',  // Slate border
    kpiValueColor: '#0f172a',  // Dark charcoal
    kpiLabelColor: '#64748b',  // Muted slate
    zebraBg: '#f1f5f9',        // Light alternate row
    accentBlue: '#2563eb',     // Blue highlight
    accentGreen: '#16a34a',    // Green highlight
    accentAmber: '#d97706',    // Amber highlight
    borderLight: '#e2e8f0'     // Row dividers
  },

  SETTINGS_SHEET_NAME: 'Settings',
  LOG_SHEET_NAME: 'Execution Log'
};


// ============================================================================
// 2. DISPATCHER SCHEDULE & CALENDAR ENGINE
// ============================================================================

const MONTH_NAMES = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December'
];

const MONTH_MAP = {
  'january': 0, 'jan': 0,
  'february': 1, 'feb': 1,
  'march': 2, 'mar': 2,
  'april': 3, 'apr': 3,
  'may': 4,
  'june': 5, 'jun': 5,
  'july': 6, 'jul': 6,
  'august': 7, 'aug': 7,
  'september': 8, 'sep': 8, 'sept': 8,
  'october': 9, 'oct': 9,
  'november': 10, 'nov': 10,
  'december': 11, 'dec': 11
};

/**
 * Parses month name/string and year from a label or filename.
 */
function parseMonthAndYear(monthStr, fallbackYear) {
  if (!monthStr || typeof monthStr !== 'string') {
    const now = new Date();
    return {
      monthIndex: now.getMonth(),
      monthName: MONTH_NAMES[now.getMonth()],
      year: fallbackYear || now.getFullYear()
    };
  }

  const clean = monthStr.trim();
  const yearMatch = clean.match(/\b(20\d\d)\b/);
  const year = yearMatch ? parseInt(yearMatch[1], 10) : (fallbackYear || new Date().getFullYear());

  const words = clean.toLowerCase().replace(/[^a-z]/g, ' ').split(/\s+/).filter(Boolean);
  let monthIndex = -1;

  for (const word of words) {
    if (MONTH_MAP.hasOwnProperty(word)) {
      monthIndex = MONTH_MAP[word];
      break;
    }
  }

  if (monthIndex === -1) {
    const lower = clean.toLowerCase();
    for (const [name, idx] of Object.entries(MONTH_MAP)) {
      if (lower.includes(name)) {
        monthIndex = idx;
        break;
      }
    }
  }

  if (monthIndex === -1) {
    monthIndex = new Date().getMonth();
  }

  return {
    monthIndex: monthIndex,
    monthName: MONTH_NAMES[monthIndex],
    year: year
  };
}

/**
 * Calculates total standard dispatch labor hours for a specific month and year.
 */
function calculateMonthlyDispatchHours(monthIndex, year, customDailyHours) {
  const schedule = customDailyHours || CONFIG.DISPATCH_HOURS_BY_DAY_OF_WEEK;
  const daysInMonth = new Date(year, monthIndex + 1, 0).getDate();
  let totalHours = 0;
  const dayCounts = { 0: 0, 1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0 };
  const dailyLog = [];

  for (let day = 1; day <= daysInMonth; day++) {
    const dateObj = new Date(year, monthIndex, day);
    const dayOfWeek = dateObj.getDay();
    const hoursToday = schedule[dayOfWeek] || 0;

    dayCounts[dayOfWeek] = (dayCounts[dayOfWeek] || 0) + 1;
    totalHours += hoursToday;

    dailyLog.push({
      date: `${year}-${String(monthIndex + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`,
      dayOfWeek: dayOfWeek,
      dayName: ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'][dayOfWeek],
      hours: hoursToday
    });
  }

  return {
    totalHours: totalHours,
    daysInMonth: daysInMonth,
    breakdown: {
      mondays: dayCounts[1],
      tuesdays: dayCounts[2],
      wednesdays: dayCounts[3],
      thursdays: dayCounts[4],
      fridays: dayCounts[5],
      saturdays: dayCounts[6],
      sundays: dayCounts[0]
    },
    dailyLog: dailyLog
  };
}


// ============================================================================
// 3. DRIVE SYNCHRONIZATION & SUMMARY TAB INGESTION
// ============================================================================

/**
 * Finds all monthly driver daily balance files in the configured Drive folder and its subfolders.
 */
function findMonthlyBalanceFiles(folderId) {
  const targetFolderId = folderId || getEffectiveDriveFolderId();
  
  if (!targetFolderId) {
    throw new Error('Google Drive Folder ID is not configured. Please check Settings or Config.js.');
  }

  let rootFolder;
  try {
    rootFolder = DriveApp.getFolderById(targetFolderId);
  } catch (err) {
    throw new Error(`Failed to access Google Drive folder with ID: "${targetFolderId}". Error: ${err.message}`);
  }

  const matchedFiles = [];
  
  // 1. Process files directly in the root folder
  scanFolderForBalanceFiles(rootFolder, matchedFiles, null);

  // 2. Process subfolders (e.g. "2025", "2026", etc.)
  const subfolders = rootFolder.getFolders();
  while (subfolders.hasNext()) {
    const subfolder = subfolders.next();
    const folderName = subfolder.getName();
    const yearMatch = folderName.match(/\b(20\d\d)\b/);
    const subfolderYear = yearMatch ? parseInt(yearMatch[1], 10) : null;
    
    scanFolderForBalanceFiles(subfolder, matchedFiles, subfolderYear);
  }

  return matchedFiles;
}

function scanFolderForBalanceFiles(folder, matchedFiles, fallbackYear) {
  const files = folder.getFiles();
  const pattern = CONFIG.FILE_NAME_PATTERN || /^(.*)\s*-\s*Drivers Daily Balance$/i;

  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();
    const match = fileName.match(pattern);

    if (match) {
      const rawMonthPart = match[1];
      const parsed = parseMonthAndYear(rawMonthPart, fallbackYear);

      matchedFiles.push({
        file: file,
        fileId: file.getId(),
        fileName: fileName,
        folderName: folder.getName(),
        year: parsed.year,
        monthName: parsed.monthName,
        monthIndex: parsed.monthIndex
      });
    }
  }
}

/**
 * Reads and extracts driver records from the 'Summary' tab of a monthly balance spreadsheet.
 */
function extractDriverDataFromSummary(fileId, monthIndex, year) {
  let spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openById(fileId);
  } catch (e) {
    throw new Error(`Unable to open spreadsheet (${fileId}): ${e.message}`);
  }

  const targetSheetName = CONFIG.SOURCE_SHEET_NAME || 'Summary';
  let sheet = spreadsheet.getSheetByName(targetSheetName);

  if (!sheet) {
    const allSheets = spreadsheet.getSheets();
    for (const s of allSheets) {
      if (s.getName().trim().toLowerCase() === targetSheetName.toLowerCase()) {
        sheet = s;
        break;
      }
    }
  }

  if (!sheet) {
    throw new Error(`Sheet tab named "${targetSheetName}" was not found in file: "${spreadsheet.getName()}".`);
  }

  const data = sheet.getDataRange().getValues();
  if (!data || data.length < 2) {
    return {
      fileName: spreadsheet.getName(),
      monthIndex: monthIndex,
      monthName: MONTH_NAMES[monthIndex],
      year: year,
      drivers: [],
      totalDriverHours: 0,
      totalTrips: 0,
      warnings: ['Sheet is empty or has no data rows.']
    };
  }

  const headerInfo = locateHeaders(data);
  if (!headerInfo || headerInfo.colMap.driver === undefined) {
    throw new Error(`Could not find a valid "Driver" column in "${targetSheetName}" tab of "${spreadsheet.getName()}".`);
  }

  const colMap = headerInfo.colMap;
  const startRow = headerInfo.headerRowIndex + 1;
  const drivers = [];
  const warnings = [];
  let totalDriverHours = 0;
  let totalTrips = 0;

  if (colMap.hours === undefined) {
    warnings.push(`Warning: "Hours" column not detected in "${spreadsheet.getName()}". Defaulting driver hours to 0.`);
  }

  for (let r = startRow; r < data.length; r++) {
    const row = data[r];
    const rawDriverName = row[colMap.driver];

    if (!rawDriverName || (typeof rawDriverName !== 'string' && typeof rawDriverName !== 'number')) continue;
    const driverName = String(rawDriverName).trim();
    if (!driverName || isSummaryRowName(driverName)) continue;

    let hours = 0;
    if (colMap.hours !== undefined && row[colMap.hours] !== undefined && row[colMap.hours] !== '') {
      const parsedHours = parseNumericValue(row[colMap.hours]);
      if (!isNaN(parsedHours) && parsedHours >= 0) {
        hours = parsedHours;
      } else {
        warnings.push(`Invalid hours value "${row[colMap.hours]}" for driver "${driverName}" on row ${r + 1}. Set to 0.`);
      }
    }

    let trips = 0;
    if (colMap.trips !== undefined && row[colMap.trips] !== undefined && row[colMap.trips] !== '') {
      const parsedTrips = parseNumericValue(row[colMap.trips]);
      if (!isNaN(parsedTrips) && parsedTrips >= 0) {
        trips = parsedTrips;
      }
    }

    const credit = (colMap.credit !== undefined) ? (parseNumericValue(row[colMap.credit]) || 0) : 0;
    const cash = (colMap.cash !== undefined) ? (parseNumericValue(row[colMap.cash]) || 0) : 0;
    const balance = (colMap.balance !== undefined) ? (parseNumericValue(row[colMap.balance]) || 0) : 0;
    const noShow = (colMap.noShow !== undefined) ? (parseNumericValue(row[colMap.noShow]) || 0) : 0;

    drivers.push({
      driverName: driverName,
      hours: hours,
      trips: trips,
      credit: credit,
      cash: cash,
      balance: balance,
      noShow: noShow
    });

    totalDriverHours += hours;
    totalTrips += trips;
  }

  return {
    fileId: fileId,
    fileName: spreadsheet.getName(),
    monthIndex: monthIndex,
    monthName: MONTH_NAMES[monthIndex],
    year: year,
    drivers: drivers,
    totalDriverHours: totalDriverHours,
    totalTrips: totalTrips,
    driverCount: drivers.length,
    warnings: warnings
  };
}

function locateHeaders(rows) {
  const aliases = CONFIG.HEADER_ALIASES;
  const maxRowsToInspect = Math.min(rows.length, 5);

  for (let r = 0; r < maxRowsToInspect; r++) {
    const row = rows[r];
    const colMap = {};

    for (let c = 0; c < row.length; c++) {
      const cellVal = String(row[c] || '').trim().toLowerCase();
      if (!cellVal) continue;

      for (const [key, aliasList] of Object.entries(aliases)) {
        if (colMap[key] !== undefined) continue;

        for (const alias of aliasList) {
          if (cellVal === alias || cellVal.includes(alias)) {
            colMap[key] = c;
            break;
          }
        }
      }
    }

    if (colMap.driver !== undefined) {
      return {
        headerRowIndex: r,
        colMap: colMap
      };
    }
  }

  return null;
}

function isSummaryRowName(name) {
  const lower = name.toLowerCase().trim();
  const summaryKeywords = ['total', 'grand total', 'subtotal', 'average', 'summary', 'fleet total', 'fleet sum'];
  return summaryKeywords.includes(lower) || lower.startsWith('total ') || lower.endsWith(' total');
}

function parseNumericValue(val) {
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  if (!val) return 0;

  const str = String(val).replace(/[$,]/g, '').trim();
  const num = parseFloat(str);
  return isNaN(num) ? 0 : num;
}


// ============================================================================
// 4. ANALYSIS SHEET BUILDER & DYNAMIC FORMULAS
// ============================================================================

function buildMonthlyAnalysisSheet(monthData, targetSpreadsheet) {
  const ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  const monthName = monthData.monthName;
  const year = monthData.year;
  const tabName = `${monthName} ${year} - Analysis`;

  let sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    sheet = ss.getSheetByName(`${monthName} - Analysis`);
  }

  if (sheet) {
    sheet.clear();
    sheet.clearConditionalFormatRules();
  } else {
    sheet = ss.insertSheet(tabName);
  }

  const scheduleResult = calculateMonthlyDispatchHours(monthData.monthIndex, year);
  const standardDispatchHours = scheduleResult.totalHours;

  const drivers = monthData.drivers || [];
  const driverCount = drivers.length;
  const startDataRow = 11;
  const endDataRow = startDataRow + Math.max(driverCount - 1, 0);
  const totalRow = endDataRow + 1;

  const totalRowsNeeded = Math.max(totalRow + 3, 25);
  const totalCols = 7;
  const matrix = Array.from({ length: totalRowsNeeded }, () => Array(totalCols).fill(''));

  // Row 1: Title Banner
  matrix[0][0] = `🚀 ${monthName.toUpperCase()} ${year} — FLEET DISPATCH EFFICIENCY & LABOR RATIO ANALYSIS`;

  // Row 2: Subtitle
  matrix[1][0] = `Generated: ${new Date().toLocaleString()} | Source File: ${monthData.fileName} | Calendar Days: ${scheduleResult.daysInMonth} days`;

  // Row 3-5: Top KPI Cards
  // Card 1: TOTAL FLEET DRIVER HOURS
  matrix[2][0] = 'TOTAL FLEET DRIVER HOURS';
  matrix[3][0] = driverCount > 0 ? `=SUM(B${startDataRow}:B${endDataRow})` : 0;
  matrix[4][0] = 'Active Road Hours';

  // Card 2: STANDARD DISPATCH HOURS
  matrix[2][2] = 'STANDARD DISPATCH HOURS';
  matrix[3][2] = standardDispatchHours;
  matrix[4][2] = `Calendar Base (${scheduleResult.daysInMonth} days, 4 Dispatchers)`;

  // Card 3: OVERTIME / ADJUSTMENTS
  matrix[2][4] = 'OVERTIME / ADJ (HRS)';
  matrix[3][4] = 0;
  matrix[4][4] = '✏️ Editable Adjustment';

  // Card 4: TOTAL DISPATCH HOURS
  matrix[2][5] = 'TOTAL DISPATCH HOURS';
  matrix[3][5] = `=$C$4+$E$4`;
  matrix[4][5] = 'Standard + Overtime';

  // Row 6-8: Ratio KPI Cards
  // Card 5: FLEET DISPATCH RATIO
  matrix[5][0] = 'FLEET DISPATCH RATIO (OVERHEAD)';
  matrix[6][0] = `=IF($A$4>0, $F$4/$A$4, 0)`;
  matrix[7][0] = 'Dispatch Hours / Driver Hour';

  // Card 6: SUPPORT MINS / ROAD HOUR
  matrix[5][2] = 'SUPPORT MINS / ROAD HOUR';
  matrix[6][2] = `=$A$7*60`;
  matrix[7][2] = 'Minutes of Dispatch per 1 Road Hr';

  // Card 7: DRIVER LEVERAGE RATIO
  matrix[5][4] = 'DRIVER LEVERAGE RATIO';
  matrix[6][4] = `=IF($F$4>0, $A$4/$F$4, 0)`;
  matrix[7][4] = 'Driver Road Hours per 1 Dispatch Hr';

  // Row 9: Section Divider
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
  applyAnalysisSheetFormatting(sheet, startDataRow, endDataRow, totalRow, driverCount);

  return sheet;
}

function applyAnalysisSheetFormatting(sheet, startDataRow, endDataRow, totalRow, driverCount) {
  const theme = CONFIG.THEME;

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
    .setBackground('#334155')
    .setFontColor('#cbd5e1')
    .setFontSize(9)
    .setFontStyle('italic')
    .setHorizontalAlignment('left');
  sheet.setRowHeight(2, 22);

  // KPI Cards
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
  sheet.getRange('E4').setNote('👉 EDIT THIS CELL to add overtime or adjustment dispatch hours.');

  sheet.getRange('F3:G3').merge().setValue('TOTAL DISPATCH HOURS');
  sheet.getRange('F4:G4').merge();
  sheet.getRange('F5:G5').merge().setValue('Standard + Overtime');
  formatKpiCard(sheet, 'F3:G5', 'F4:G4', '#,##0.0 "hrs"', '#0f172a');

  // Ratio Cards
  sheet.getRange('A6:B6').merge().setValue('FLEET DISPATCH RATIO');
  sheet.getRange('A7:B7').merge();
  sheet.getRange('A8:B8').merge().setValue('Dispatch Hours / Road Hr');
  formatKpiCard(sheet, 'A6:B8', 'A7:B7', '0.000', '#0369a1');

  sheet.getRange('C6:D6').merge().setValue('SUPPORT MINS / ROAD HR');
  sheet.getRange('C7:D7').merge();
  sheet.getRange('C8:D8').merge().setValue('Dispatch Mins per 1 Road Hr');
  formatKpiCard(sheet, 'C6:D8', 'C7:D7', '0.0 "mins"', '#4338ca');

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

  // Section Header
  sheet.getRange('A9:G9').merge()
    .setBackground('#0f172a')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('left');
  sheet.setRowHeight(9, 26);

  // Table Headers
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
      if (r % 2 === 0) {
        sheet.getRange(r, 1, 1, 7).setBackground(theme.zebraBg);
      } else {
        sheet.getRange(r, 1, 1, 7).setBackground('#ffffff');
      }
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

  sheet.setFrozenRows(10);
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
    .setFontSize(16)
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


// ============================================================================
// 5. SETTINGS MANAGER
// ============================================================================

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

  // Section 1
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

  // Section 2
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

  // Section 3
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

  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 200);

  SpreadsheetApp.getActiveSpreadsheet().toast('Settings tab initialized successfully!', 'Fleet Tools', 4);
}


// ============================================================================
// 6. UI MENU & TRIGGER HANDLERS
// ============================================================================

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

function menuSyncAllMonthlyFiles() {
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

function menuRecalculateRatios() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let updatedCount = 0;

  for (const sheet of sheets) {
    const name = sheet.getName();
    if (name.includes('- Analysis')) {
      const monthPart = name.replace('- Analysis', '').trim();
      const parsed = parseMonthAndYear(monthPart);
      const schedule = calculateMonthlyDispatchHours(parsed.monthIndex, parsed.year);

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


// ============================================================================
// 7. UTILITIES & MOCK GENERATOR
// ============================================================================

function logExecution(action, status, message, details) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;

    const logSheetName = CONFIG.LOG_SHEET_NAME || 'Execution Log';
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

function showToast(message, title, timeoutSeconds) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title || 'Fleet Dispatch Tools', timeoutSeconds || 5);
  } catch (e) {
    Logger.log(`Toast: [${title}] ${message}`);
  }
}

function showAlert(message, title) {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert(title || 'Fleet Dispatch Tools', message, ui.ButtonSet.OK);
  } catch (e) {
    Logger.log(`Alert: [${title}] ${message}`);
  }
}

function createSampleMonthlySpreadsheet(monthName, year) {
  const mName = monthName || 'August';
  const y = year || 2026;
  const fileName = `${mName} ${y} - Drivers Daily Balance`;

  const folderId = getEffectiveDriveFolderId();
  let folder;
  try {
    folder = DriveApp.getFolderById(folderId);
  } catch (e) {
    throw new Error(`Cannot access Drive folder (${folderId}) to create sample file: ${e.message}`);
  }

  const newSs = SpreadsheetApp.create(fileName);
  const targetSheetName = CONFIG.SOURCE_SHEET_NAME || 'Summary';

  let sheet = newSs.getActiveSheet();
  sheet.setName(targetSheetName);

  const sampleDrivers = [
    { name: 'Marcus Vance (101)', hours: 168.5, trips: 224, credit: 3200, cash: 450, balance: 3650, noShow: 3 },
    { name: 'Elena Rostova (102)', hours: 155.0, trips: 210, credit: 2950, cash: 620, balance: 3570, noShow: 1 },
    { name: 'David Kim (103)', hours: 172.0, trips: 245, credit: 3410, cash: 510, balance: 3920, noShow: 2 },
    { name: 'Fatima Al-Mansoor (104)', hours: 144.5, trips: 198, credit: 2800, cash: 390, balance: 3190, noShow: 4 },
    { name: 'Carlos Mendez (105)', hours: 160.0, trips: 215, credit: 3100, cash: 480, balance: 3580, noShow: 2 },
    { name: 'Aisha Diallo (106)', hours: 138.0, trips: 180, credit: 2600, cash: 310, balance: 2910, noShow: 1 },
    { name: 'James O\'Connor (107)', hours: 180.5, trips: 260, credit: 3750, cash: 700, balance: 4450, noShow: 5 },
    { name: 'Priya Sharma (108)', hours: 152.0, trips: 205, credit: 2900, cash: 420, balance: 3320, noShow: 0 },
    { name: 'Lucas Silva (109)', hours: 140.0, trips: 190, credit: 2750, cash: 360, balance: 3110, noShow: 2 },
    { name: 'Tariq Hassan (110)', hours: 165.0, trips: 230, credit: 3300, cash: 540, balance: 3840, noShow: 3 },
    { name: 'Nathalie Dupont (111)', hours: 125.5, trips: 165, credit: 2400, cash: 290, balance: 2690, noShow: 1 },
    { name: 'Viktor Ivanov (112)', hours: 158.0, trips: 218, credit: 3150, cash: 490, balance: 3640, noShow: 2 },
    { name: 'Grace Hopper (113)', hours: 148.0, trips: 200, credit: 2850, cash: 410, balance: 3260, noShow: 1 },
    { name: 'Kenji Sato (114)', hours: 135.5, trips: 175, credit: 2500, cash: 350, balance: 2850, noShow: 0 },
    { name: 'Alan Turing (115)', hours: 0.0, trips: 0, credit: 0, cash: 0, balance: 0, noShow: 0 }
  ];

  const headers = ['Driver Name', 'Active Hours', 'Completed Trips', 'Credit Card ($)', 'Cash ($)', 'Net Balance ($)', 'No Shows'];
  const rows = [headers];
  let sumHours = 0;
  let sumTrips = 0;
  let sumCredit = 0;
  let sumCash = 0;
  let sumBal = 0;
  let sumNoShow = 0;

  for (const d of sampleDrivers) {
    rows.push([d.name, d.hours, d.trips, d.credit, d.cash, d.balance, d.noShow]);
    sumHours += d.hours;
    sumTrips += d.trips;
    sumCredit += d.credit;
    sumCash += d.cash;
    sumBal += d.balance;
    sumNoShow += d.noShow;
  }

  rows.push(['Total Fleet Sum', sumHours, sumTrips, sumCredit, sumCash, sumBal, sumNoShow]);
  sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);

  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#1e293b')
    .setFontColor('#ffffff')
    .setFontWeight('bold');

  sheet.getRange(2, 2, rows.length - 1, 1).setNumberFormat('#,##0.0');
  sheet.getRange(2, 3, rows.length - 1, 1).setNumberFormat('#,##0');
  sheet.getRange(2, 4, rows.length - 1, 3).setNumberFormat('$#,##0.00');

  const driveFile = DriveApp.getFileById(newSs.getId());
  driveFile.moveTo(folder);

  return driveFile;
}
