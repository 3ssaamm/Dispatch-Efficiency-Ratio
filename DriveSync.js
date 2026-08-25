/**
 * @fileoverview Google Drive file discovery and data extraction engine.
 * Recursively inspects root folder & year subfolders (2025, 2026),
 * dynamically maps headers in the 'Summary' tab, aggregates daily records,
 * and seamlessly NORMALIZES & MERGES single-names (e.g. "Brian") with full-names ("Brian Macancela").
 */

// Node.js fallback import for testing
let _CONFIG = (typeof CONFIG !== 'undefined') ? CONFIG : null;
if (!_CONFIG && typeof require !== 'undefined') {
  try {
    _CONFIG = require('./Config').CONFIG;
  } catch (e) {}
}

/**
 * Finds all monthly driver daily balance files in the configured Drive folder and its subfolders.
 * @param {string} [folderId] - Google Drive folder ID
 * @returns {Array<{ file: GoogleAppsScript.Drive.File, folderName: string, year: number, monthName: string, monthIndex: number }>}
 */
function findMonthlyBalanceFiles(folderId) {
  const activeCfg = (typeof CONFIG !== 'undefined') ? CONFIG : _CONFIG;
  const targetFolderId = folderId || (typeof getEffectiveDriveFolderId === 'function' ? getEffectiveDriveFolderId() : activeCfg.DRIVE_FOLDER_ID);
  
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

/**
 * Helper to scan a single folder for matching balance files.
 */
function scanFolderForBalanceFiles(folder, matchedFiles, fallbackYear) {
  const activeCfg = (typeof CONFIG !== 'undefined') ? CONFIG : _CONFIG;
  const files = folder.getFiles();
  const pattern = (activeCfg && activeCfg.FILE_NAME_PATTERN) ? activeCfg.FILE_NAME_PATTERN : /^(.*)\s*-\s*Drivers Daily Balance$/i;

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
 * Normalizes a driver name and merges single-name variations with canonical full names.
 * E.g. "Brian" -> "Brian Macancela", "Angel" -> "Angel Yoy", "Nikolay" -> "Nikolay Iankov"
 * 
 * @param {string} rawName - Name from spreadsheet
 * @param {Object} [dynamicAliasMap] - Optional map of first-name to full-name
 * @returns {string} Canonical driver name
 */
function normalizeDriverName(rawName, dynamicAliasMap) {
  if (!rawName) return '';
  let clean = String(rawName).trim();
  
  // Strip trailing driver IDs like "(101)" or "#12"
  clean = clean.replace(/\s*\(\d+\)\s*$/, '').replace(/\s*#\d+\s*$/, '').trim();

  // Normalize multiple consecutive spaces
  clean = clean.replace(/\s+/g, ' ');

  const lower = clean.toLowerCase();
  const activeCfg = (typeof CONFIG !== 'undefined') ? CONFIG : _CONFIG;

  // 1. Check explicit alias map from CONFIG
  if (activeCfg && activeCfg.DRIVER_ALIASES && activeCfg.DRIVER_ALIASES[lower]) {
    return activeCfg.DRIVER_ALIASES[lower];
  }

  // 2. Check dynamic multi-month alias map
  if (dynamicAliasMap && dynamicAliasMap[lower]) {
    return dynamicAliasMap[lower];
  }

  return clean;
}

/**
 * Reads and extracts driver records from the 'Summary' tab of a monthly balance spreadsheet.
 * AGGREGATES all daily rows for the same driver into a single comprehensive monthly record
 * with driver name merging.
 * 
 * @param {string} fileId - ID of the monthly Google Spreadsheet
 * @param {number} monthIndex - Month index (0-11)
 * @param {number} year - 4-digit Year
 * @param {Object} [globalAliasMap] - Optional global name alias map
 * @returns {Object} Extracted monthly data bundle
 */
function extractDriverDataFromSummary(fileId, monthIndex, year, globalAliasMap) {
  let spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openById(fileId);
  } catch (e) {
    throw new Error(`Unable to open spreadsheet (${fileId}): ${e.message}`);
  }

  const activeCfg = (typeof CONFIG !== 'undefined') ? CONFIG : _CONFIG;
  const targetSheetName = (activeCfg && activeCfg.SOURCE_SHEET_NAME) ? activeCfg.SOURCE_SHEET_NAME : 'Summary';
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
  const warnings = [];

  // Pass 1: Scan all driver names in this workbook to detect full names (e.g. "Brian Macancela")
  const localFullNameMap = {};
  for (let r = startRow; r < data.length; r++) {
    const rawVal = data[r][colMap.driver];
    if (rawVal && typeof rawVal === 'string') {
      const nameClean = rawVal.trim().replace(/\s*\(\d+\)\s*$/, '').replace(/\s*#\d+\s*$/, '').trim();
      const parts = nameClean.split(/\s+/);
      if (parts.length >= 2 && !isSummaryRowName(nameClean)) {
        const firstNameLower = parts[0].toLowerCase();
        if (!localFullNameMap[firstNameLower]) {
          localFullNameMap[firstNameLower] = nameClean;
        }
      }
    }
  }

  // Combine local map and global alias map
  const activeAliasMap = Object.assign({}, localFullNameMap, globalAliasMap || {});

  const driverMap = {}; // Key: normalized lower-cased canonical driver name
  let totalRawRowsProcessed = 0;

  for (let r = startRow; r < data.length; r++) {
    const row = data[r];
    const rawDriverName = row[colMap.driver];

    if (!rawDriverName || (typeof rawDriverName !== 'string' && typeof rawDriverName !== 'number')) continue;
    const rawClean = String(rawDriverName).trim();
    if (!rawClean || isSummaryRowName(rawClean)) continue;

    // Normalize and resolve single names (e.g. "Brian" -> "Brian Macancela")
    const canonicalName = normalizeDriverName(rawClean, activeAliasMap);

    let hours = 0;
    if (colMap.hours !== undefined && row[colMap.hours] !== undefined && row[colMap.hours] !== '') {
      const parsedHours = parseNumericValue(row[colMap.hours]);
      if (!isNaN(parsedHours) && parsedHours >= 0) {
        hours = parsedHours;
      } else {
        warnings.push(`Invalid hours value "${row[colMap.hours]}" for driver "${canonicalName}" on row ${r + 1}.`);
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

    const key = canonicalName.toLowerCase();
    if (!driverMap[key]) {
      driverMap[key] = {
        driverName: canonicalName,
        hours: 0,
        trips: 0,
        credit: 0,
        cash: 0,
        balance: 0,
        noShow: 0,
        dailyEntries: 0
      };
    }

    driverMap[key].hours += hours;
    driverMap[key].trips += trips;
    driverMap[key].credit += credit;
    driverMap[key].cash += cash;
    driverMap[key].balance += balance;
    driverMap[key].noShow += noShow;
    driverMap[key].dailyEntries += 1;
    totalRawRowsProcessed++;
  }

  const uniqueDrivers = Object.values(driverMap);
  let totalDriverHours = 0;
  let totalTrips = 0;

  for (const d of uniqueDrivers) {
    d.hours = Math.round(d.hours * 10) / 10;
    d.credit = Math.round(d.credit * 100) / 100;
    d.cash = Math.round(d.cash * 100) / 100;
    d.balance = Math.round(d.balance * 100) / 100;

    totalDriverHours += d.hours;
    totalTrips += d.trips;
  }

  uniqueDrivers.sort((a, b) => b.hours - a.hours || a.driverName.localeCompare(b.driverName));

  return {
    fileId: fileId,
    fileName: spreadsheet.getName(),
    monthIndex: monthIndex,
    monthName: MONTH_NAMES[monthIndex],
    year: year,
    drivers: uniqueDrivers,
    totalDriverHours: Math.round(totalDriverHours * 10) / 10,
    totalTrips: totalTrips,
    driverCount: uniqueDrivers.length,
    rawRowsProcessed: totalRawRowsProcessed,
    warnings: warnings
  };
}

function locateHeaders(rows) {
  const activeCfg = (typeof CONFIG !== 'undefined') ? CONFIG : _CONFIG;
  const aliases = (activeCfg && activeCfg.HEADER_ALIASES) ? activeCfg.HEADER_ALIASES : {
    driver: ['driver', 'driver name', 'name', 'driver id'],
    hours: ['hours', 'active hours', 'driver hours', 'total hours', 'hrs'],
    trips: ['trips', 'completed trips', 'total trips'],
    credit: ['credit', 'credits', 'cc'],
    cash: ['cash'],
    balance: ['balance', 'net balance'],
    noShow: ['no show', 'no-show', 'noshow']
  };

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

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    findMonthlyBalanceFiles,
    scanFolderForBalanceFiles,
    normalizeDriverName,
    extractDriverDataFromSummary,
    locateHeaders,
    isSummaryRowName,
    parseNumericValue
  };
}
