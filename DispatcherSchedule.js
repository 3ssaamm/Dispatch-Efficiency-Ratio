/**
 * @fileoverview Dispatcher labor hours calculation & timesheet integration engine.
 * Computes exact monthly dispatch hours across all timesheet formats:
 * - 2026 Format: Management Summary block or explicit WT & OT columns (Muhammad, Mariam, Nourween).
 * - 2025 Format (e.g. September 25, October 25): Overtime-only columns with base shift schedules.
 * - Nour: Fixed 9h/day Mon-Fri.
 * - Excludes: Mohanad, Abdulrahman, Fares.
 */

// Node.js fallback import for testing
let _CONFIG_DS = (typeof CONFIG !== 'undefined') ? CONFIG : null;
if (!_CONFIG_DS && typeof require !== 'undefined') {
  try {
    _CONFIG_DS = require('./Config').CONFIG;
  } catch (e) {}
}

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
 * Calculates standard base shift hours for an individual dispatcher in a given month.
 */
function calculateBaseContractHours(dispatcherName, monthIndex, year) {
  const daysInMonth = new Date(year, monthIndex + 1, 0).getDate();
  const lower = dispatcherName.toLowerCase();
  let totalHours = 0;
  let workDaysCount = 0;

  for (let day = 1; day <= daysInMonth; day++) {
    const dow = new Date(year, monthIndex, day).getDay();

    if (lower.includes('muhammad') || lower.includes('mohamed')) {
      // Muhammad: 10h Mon-Fri (dow 1..5)
      if (dow >= 1 && dow <= 5) {
        totalHours += 10;
        workDaysCount++;
      }
    } else if (lower.includes('mariam')) {
      // Mariam: 10h Mon, Tue, Wed, Fri, Sat (dow 1,2,3,5,6) | Thu (4) & Sun (0) OFF
      if (dow === 1 || dow === 2 || dow === 3 || dow === 5 || dow === 6) {
        totalHours += 10;
        workDaysCount++;
      }
    } else if (lower.includes('nourween')) {
      // Nourween: 10h Mon-Fri (dow 1..5)
      if (dow >= 1 && dow <= 5) {
        totalHours += 10;
        workDaysCount++;
      }
    } else if (lower.includes('nour')) {
      // Nour: 9h Mon-Fri (dow 1..5)
      if (dow >= 1 && dow <= 5) {
        totalHours += 9;
        workDaysCount++;
      }
    }
  }

  return { hours: totalHours, days: workDaysCount };
}

/**
 * Calculates calendar schedule fallback when timesheet is not present.
 */
function calculateMonthlyDispatchHours(monthIndex, year, customDailyHours) {
  const activeCfg = (typeof CONFIG !== 'undefined') ? CONFIG : _CONFIG_DS;
  const schedule = customDailyHours || (activeCfg && activeCfg.DISPATCH_HOURS_BY_DAY_OF_WEEK) || {
    0: 0,   // Sun
    1: 39,  // Mon
    2: 39,  // Tue
    3: 39,  // Wed
    4: 29,  // Thu
    5: 39,  // Fri
    6: 10   // Sat (Mariam 10h)
  };

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

/**
 * Calculates fixed schedule hours for Nour (9h Mon-Fri).
 */
function calculateNourFixedHours(monthIndex, year) {
  const daysInMonth = new Date(year, monthIndex + 1, 0).getDate();
  let activeDays = 0;

  for (let day = 1; day <= daysInMonth; day++) {
    const dow = new Date(year, monthIndex, day).getDay();
    if (dow >= 1 && dow <= 5) {
      activeDays++;
    }
  }

  return {
    hours: activeDays * 9.0,
    activeDays: activeDays
  };
}

/**
 * Opens the Working Time spreadsheet via multiple access methods.
 */
function getTimesheetSpreadsheet() {
  const activeCfg = (typeof CONFIG !== 'undefined') ? CONFIG : _CONFIG_DS;
  const timesheetId = activeCfg && activeCfg.WORKING_TIME_SPREADSHEET_ID;

  if (typeof SpreadsheetApp === 'undefined') return null;

  // Method 1: DriveApp.getFileById
  if (typeof DriveApp !== 'undefined' && timesheetId) {
    try {
      const file = DriveApp.getFileById(timesheetId);
      if (file) return SpreadsheetApp.open(file);
    } catch (e) {
      Logger.log(`DriveApp openById failed: ${e.message}`);
    }
  }

  // Method 2: SpreadsheetApp.openById
  if (timesheetId) {
    try {
      return SpreadsheetApp.openById(timesheetId);
    } catch (e) {
      Logger.log(`SpreadsheetApp.openById failed: ${e.message}`);
    }
  }

  // Method 3: Search master Drive folder for 'Working Time'
  if (typeof DriveApp !== 'undefined') {
    try {
      const folderId = (typeof getEffectiveDriveFolderId === 'function') 
        ? getEffectiveDriveFolderId() 
        : (activeCfg && activeCfg.DRIVE_FOLDER_ID);

      if (folderId) {
        const folder = DriveApp.getFolderById(folderId);
        const files = folder.getFiles();
        while (files.hasNext()) {
          const f = files.next();
          const fName = f.getName().toLowerCase();
          if (fName.includes('working') || fName.includes('time') || fName.includes('timesheet')) {
            return SpreadsheetApp.open(f);
          }
        }
      }
    } catch (e) {
      Logger.log(`Drive folder search failed: ${e.message}`);
    }
  }

  return null;
}

/**
 * Directly fetches dispatcher working hours and overtime from the Working Time spreadsheet.
 * Supports both 2026 full summary format and 2025 overtime-only formats (Sep 25, Oct 25).
 * 
 * @param {number} monthIndex - Month index (0-11)
 * @param {number} year - 4-digit Year
 * @returns {Object} Comprehensive dispatch labor bundle
 */
function fetchDispatcherHoursFromTimesheet(monthIndex, year) {
  const activeCfg = (typeof CONFIG !== 'undefined') ? CONFIG : _CONFIG_DS;
  const nourCalc = calculateNourFixedHours(monthIndex, year);
  const monthName = MONTH_NAMES[monthIndex];
  const shortYear = String(year).slice(-2); // e.g. "25" or "26"

  let timesheetSs = null;
  try {
    timesheetSs = getTimesheetSpreadsheet();
  } catch (err) {
    Logger.log(`getTimesheetSpreadsheet error: ${err.message}`);
  }

  if (!timesheetSs) {
    const cal = calculateMonthlyDispatchHours(monthIndex, year);
    return {
      standardHours: cal.totalHours,
      overtimeHours: 0,
      totalHours: cal.totalHours,
      source: 'calendar',
      monthName: monthName,
      year: year,
      tabFound: false,
      tabName: '',
      permissionError: true,
      timesheetStaff: [],
      fixedStaff: [{ name: 'Nour', regularHours: nourCalc.hours, overtimeHours: 0, total: nourCalc.hours, notes: `${nourCalc.activeDays} weekdays × 9h` }],
      excludedStaff: []
    };
  }

  try {
    const allSheets = timesheetSs.getSheets();
    
    // Tab name search patterns (e.g. "October 25", "October 2025", "Oct 25", "October '25")
    const targetNames = [
      `${monthName} ${shortYear}`,
      `${monthName} ${year}`,
      `${monthName}`,
      `${monthName.slice(0, 3)} ${shortYear}`,
      `${monthName.slice(0, 3)} ${year}`,
      `${monthName.slice(0, 3)}`
    ];

    let sheet = null;
    for (const s of allSheets) {
      const sName = s.getName().trim().toLowerCase().replace(/['’]/g, '');
      for (const t of targetNames) {
        const tClean = t.toLowerCase().replace(/['’]/g, '');
        if (sName === tClean) {
          sheet = s;
          break;
        }
      }
      if (sheet) break;
    }

    if (!sheet) {
      // Month tab was not found in timesheet
      return {
        standardHours: 0,
        overtimeHours: 0,
        totalHours: 0,
        source: 'no_timesheet',
        monthName: monthName,
        year: year,
        tabFound: false,
        tabName: 'No Timesheet Tab',
        permissionError: false,
        timesheetStaff: [],
        fixedStaff: [{ name: 'Nour', regularHours: 0, overtimeHours: 0, total: 0, notes: '0.0 hrs' }],
        excludedStaff: []
      };
    }

    const data = sheet.getDataRange().getValues();
    const activeTimesheetNames = (activeCfg && activeCfg.ACTIVE_TIMESHEET_DISPATCHERS) || ['muhammad', 'mohamed', 'mariam', 'nourween'];
    const excludedNames = (activeCfg && activeCfg.EXCLUDED_DISPATCHER_NAMES) || ['mohanad', 'muhanad', 'abdulrahman', 'abdelrahman', 'abdo', 'fares'];

    let timesheetStandardWT = 0;
    let timesheetOvertimeOT = 0;
    const timesheetStaff = [];
    const excludedStaff = [];

    // ------------------------------------------------------------------------
    // PARSER STRATEGY 1: Management Summary Block (Cols J to N in 2026 tabs)
    // ------------------------------------------------------------------------
    let foundSummaryBlock = false;
    for (let r = 0; r < Math.min(data.length, 10); r++) {
      for (let c = 0; c < data[r].length; c++) {
        const cell = String(data[r][c] || '').trim().toLowerCase();
        if (cell.includes('management summary') || cell === 'employee name') {
          const headerRow = cell === 'employee name' ? r : r + 1;
          const nameCol = c;
          const wtCol = nameCol + 1;
          const otCol = nameCol + 2;

          for (let sr = headerRow + 1; sr < data.length; sr++) {
            const empRaw = String(data[sr][nameCol] || '').trim();
            if (!empRaw || empRaw.toLowerCase().includes('grand total') || empRaw.toLowerCase().includes('total')) continue;

            const empLower = empRaw.toLowerCase();
            const isExcluded = excludedNames.some(ex => empLower.includes(ex));
            const isActive = activeTimesheetNames.some(al => empLower.includes(al));

            const wt = parseNumericValueSafe(data[sr][wtCol]);
            const ot = parseNumericValueSafe(data[sr][otCol]);

            if (isActive && !isExcluded) {
              timesheetStandardWT += wt;
              timesheetOvertimeOT += ot;
              timesheetStaff.push({
                name: empRaw,
                regularHours: wt,
                overtimeHours: ot,
                total: wt + ot,
                source: 'Timesheet Summary'
              });
              foundSummaryBlock = true;
            } else if (isExcluded) {
              excludedStaff.push({
                name: empRaw,
                regularHours: wt,
                overtimeHours: ot,
                total: wt + ot,
                reason: 'Excluded Staff'
              });
            }
          }
          break;
        }
      }
      if (foundSummaryBlock) break;
    }

    // ------------------------------------------------------------------------
    // PARSER STRATEGY 2: Column Inspection (Handles both 2026 WT/OT and 2025 OT-only tabs)
    // ------------------------------------------------------------------------
    if (!foundSummaryBlock) {
      // Find the TOTAL row
      let totalRowIdx = -1;
      for (let r = data.length - 1; r >= 0; r--) {
        if (String(data[r][0] || '').trim().toLowerCase() === 'total') {
          totalRowIdx = r;
          break;
        }
      }

      if (totalRowIdx !== -1) {
        const headerRow = data[0];
        const staffMap = {};

        for (let c = 1; c < headerRow.length; c++) {
          const colHeader = String(headerRow[c] || '').trim();
          const colLower = colHeader.toLowerCase();
          if (!colHeader) continue;

          const isExcluded = excludedNames.some(ex => colLower.includes(ex));
          const val = parseNumericValueSafe(data[totalRowIdx][c]);

          if (isExcluded) {
            // Identify excluded staff member
            let exName = colHeader.replace(/\s*(wt|ot|dr|working|overtime)\s*$/i, '').trim();
            let existing = excludedStaff.find(ex => ex.name.toLowerCase() === exName.toLowerCase());
            if (!existing) {
              existing = { name: exName, regularHours: 0, overtimeHours: 0, total: 0, reason: 'Excluded Staff' };
              excludedStaff.push(existing);
            }
            if (colLower.includes('ot') || colLower.includes('overtime')) {
              existing.overtimeHours += val;
            } else {
              existing.regularHours += val;
            }
            existing.total = existing.regularHours + existing.overtimeHours;
            continue;
          }

          // Check active dispatchers (Muhammad, Mariam, Nourween)
          for (const actName of activeTimesheetNames) {
            if (colLower.includes(actName)) {
              if (!staffMap[actName]) {
                staffMap[actName] = {
                  name: colHeader.replace(/\s*(wt|ot|working|overtime)\s*$/i, '').trim() || actName,
                  hasExplicitWT: false,
                  regularHours: 0,
                  overtimeHours: 0
                };
              }

              if (colLower.includes('wt') || colLower.includes('working')) {
                staffMap[actName].regularHours += val;
                staffMap[actName].hasExplicitWT = true;
              } else if (colLower.includes('ot') || colLower.includes('overtime')) {
                staffMap[actName].overtimeHours += val;
              }
            }
          }
        }

        // Process found active staff
        for (const [actKey, staffObj] of Object.entries(staffMap)) {
          // If no explicit WT column was present (e.g. Sep 25, Oct 25 where only OT was logged):
          // Calculate standard base contract hours!
          if (!staffObj.hasExplicitWT) {
            const baseCalc = calculateBaseContractHours(staffObj.name, monthIndex, year);
            staffObj.regularHours = baseCalc.hours;
          }

          timesheetStandardWT += staffObj.regularHours;
          timesheetOvertimeOT += staffObj.overtimeHours;
          timesheetStaff.push({
            name: staffObj.name,
            regularHours: staffObj.regularHours,
            overtimeHours: staffObj.overtimeHours,
            total: staffObj.regularHours + staffObj.overtimeHours,
            source: staffObj.hasExplicitWT ? 'Timesheet Columns' : 'Base Shift + Timesheet OT'
          });
          foundSummaryBlock = true;
        }
      }
    }

    // Add Nour (Fixed 9h Mon-Fri)
    const nourHours = nourCalc.hours;
    const totalStandard = timesheetStandardWT + nourHours;
    const totalOvertime = timesheetOvertimeOT;

    return {
      standardHours: Math.round(totalStandard * 10) / 10,
      overtimeHours: Math.round(totalOvertime * 10) / 10,
      totalHours: Math.round((totalStandard + totalOvertime) * 10) / 10,
      source: 'timesheet',
      monthName: monthName,
      year: year,
      tabFound: true,
      tabName: sheet.getName(),
      permissionError: false,
      timesheetStaff: timesheetStaff,
      fixedStaff: [{ 
        name: 'Nour', 
        regularHours: nourHours, 
        overtimeHours: 0, 
        total: nourHours, 
        notes: `${nourCalc.activeDays} weekdays × 9h/day` 
      }],
      excludedStaff: excludedStaff
    };

  } catch (err) {
    Logger.log(`Timesheet fetch error for ${monthIndex}/${year}: ${err.message}`);
    return {
      standardHours: 0,
      overtimeHours: 0,
      totalHours: 0,
      source: 'error',
      monthName: MONTH_NAMES[monthIndex],
      year: year,
      tabFound: false,
      tabName: '',
      permissionError: true,
      permissionErrorMessage: err.message,
      timesheetStaff: [],
      fixedStaff: [{ name: 'Nour', regularHours: 0, overtimeHours: 0, total: 0, notes: '0.0 hrs' }],
      excludedStaff: []
    };
  }
}

/**
 * Diagnostic tool: Tests timesheet connection and reports detailed status.
 */
function testTimesheetConnection() {
  let userEmail = '';
  try {
    userEmail = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || 'Current Google User';
  } catch (e) {
    userEmail = 'Current Google User';
  }

  showToast('Testing Timesheet connection...', 'Diagnostic Test', 5);

  const timesheetSs = getTimesheetSpreadsheet();

  if (!timesheetSs) {
    const errorMsg = 
      `❌ Could not access the "Working Time" spreadsheet.\n\n` +
      `👤 Current Running User: ${userEmail}\n` +
      `📄 Target Sheet ID: ${CONFIG.WORKING_TIME_SPREADSHEET_ID}\n\n` +
      `🔧 HOW TO FIX (Choose either):\n` +
      `1. Open the Working Time sheet:\n` +
      `   ${CONFIG.WORKING_TIME_SPREADSHEET_URL}\n` +
      `   Click "Share" -> Change General Access to "Anyone with the link can view".\n\n` +
      `2. OR Move the "Working Time" spreadsheet into your Google Drive folder:\n` +
      `   ${CONFIG.DRIVE_FOLDER_URL}`;

    showAlert(errorMsg, 'Timesheet Access Test: FAILED');
    return false;
  }

  const title = timesheetSs.getName();
  const sheets = timesheetSs.getSheets();
  const sheetNames = sheets.map(s => `"${s.getName()}"`).join(', ');

  const octTest = fetchDispatcherHoursFromTimesheet(9, 2025); // October 2025
  const staffList = octTest.timesheetStaff.map(s => `• ${s.name}: ${s.regularHours} WT + ${s.overtimeHours} OT`).join('\n');
  const exclList = octTest.excludedStaff.map(s => `• ${s.name}: ${s.total} hrs [EXCLUDED]`).join('\n');

  const successMsg = 
    `✅ Timesheet Connection SUCCESSFUL!\n\n` +
    `📄 File Title: "${title}"\n` +
    `📑 Tabs Found: ${sheetNames}\n\n` +
    `👥 October 2025 Test Parse:\n` +
    (staffList || '• None found in summary table') + `\n` +
    `• Nour (Fixed Shift): ${octTest.fixedStaff[0]?.regularHours || 0} hrs\n\n` +
    `🚫 Excluded Staff (Mohanad, Fares, Abdulrahman):\n` +
    (exclList || '• None') + `\n\n` +
    `📊 October 2025 Totals:\n` +
    `• Regular WT: ${octTest.standardHours} hrs\n` +
    `• Overtime OT: ${octTest.overtimeHours} hrs\n` +
    `• Total Dispatch: ${octTest.totalHours} hrs`;

  showAlert(successMsg, 'Timesheet Access Test: SUCCESS');
  return true;
}

function parseNumericValueSafe(val) {
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  if (!val) return 0;
  const num = parseFloat(String(val).replace(/[$,£]/g, '').trim());
  return isNaN(num) ? 0 : num;
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    MONTH_NAMES,
    MONTH_MAP,
    parseMonthAndYear,
    calculateBaseContractHours,
    calculateMonthlyDispatchHours,
    calculateNourFixedHours,
    fetchDispatcherHoursFromTimesheet,
    getTimesheetSpreadsheet
  };
}
