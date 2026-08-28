/**
 * @fileoverview Dispatcher labor hours calculation & timesheet integration engine.
 * Computes exact monthly dispatch hours by:
 * 1. Fetching verified timesheet hours for Muhammad (10h), Mariam (10h), and Nourween (10h).
 * 2. Adding fixed schedule hours for Nour (9h Mon-Fri).
 * 3. Excluding Mohanad and Abdulrahman.
 * 4. Gracefully falling back to calendar schedule if timesheet permissions are restricted.
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
 * Calculates standard dispatch hours based on calendar schedule fallback:
 * - Muhammad: 10h Mon-Fri (Sat/Sun OFF)
 * - Mariam: 10h Mon-Wed, Fri, Sat (Thu & Sun OFF)
 * - Nourween: 10h Mon-Fri (Sat/Sun OFF)
 * - Nour: 9h Mon-Fri (Sat/Sun OFF)
 * 
 * Daily breakdown:
 * - Mon, Tue, Wed, Fri: 10 + 10 + 10 + 9 = 39.0 hrs/day
 * - Thursday: 10 + 0 + 10 + 9 = 29.0 hrs/day
 * - Saturday: 0 + 10 (Mariam) + 0 + 0 = 10.0 hrs/day
 * - Sunday: 0.0 hrs/day
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
 * Calculates fixed schedule hours for Nour (9h Mon-Fri) for any month/year.
 * @param {number} monthIndex
 * @param {number} year
 * @returns {{ hours: number, activeDays: number }}
 */
function calculateNourFixedHours(monthIndex, year) {
  const daysInMonth = new Date(year, monthIndex + 1, 0).getDate();
  let activeDays = 0;

  for (let day = 1; day <= daysInMonth; day++) {
    const dow = new Date(year, monthIndex, day).getDay();
    if (dow >= 1 && dow <= 5) { // Mon, Tue, Wed, Thu, Fri
      activeDays++;
    }
  }

  return {
    hours: activeDays * 9.0,
    activeDays: activeDays
  };
}

/**
 * Directly fetches dispatcher working hours and overtime from the Working Time spreadsheet.
 * Fetches: Muhammad, Mariam, Nourween from timesheet.
 * Adds: Nour (Fixed 9h/day Mon-Fri).
 * Excludes: Mohanad, Abdulrahman.
 * 
 * @param {number} monthIndex - Month index (0-11)
 * @param {number} year - 4-digit Year
 * @returns {Object} Comprehensive dispatch labor bundle
 */
function fetchDispatcherHoursFromTimesheet(monthIndex, year) {
  const activeCfg = (typeof CONFIG !== 'undefined') ? CONFIG : _CONFIG_DS;
  const timesheetId = activeCfg && activeCfg.WORKING_TIME_SPREADSHEET_ID;

  // Calculate Nour's fixed hours (9h Mon-Fri)
  const nourCalc = calculateNourFixedHours(monthIndex, year);

  if (!timesheetId || typeof SpreadsheetApp === 'undefined') {
    const cal = calculateMonthlyDispatchHours(monthIndex, year);
    return {
      standardHours: cal.totalHours,
      overtimeHours: 0,
      totalHours: cal.totalHours,
      source: 'calendar',
      monthName: MONTH_NAMES[monthIndex],
      year: year,
      tabFound: false,
      tabName: '',
      permissionError: false,
      timesheetStaff: [],
      fixedStaff: [{ name: 'Nour', regularHours: nourCalc.hours, overtimeHours: 0, total: nourCalc.hours, notes: `${nourCalc.activeDays} weekdays × 9h` }],
      excludedStaff: []
    };
  }

  try {
    let timesheetSs = null;
    try {
      timesheetSs = SpreadsheetApp.openById(timesheetId);
    } catch (permErr) {
      // Gracefully handle permission restriction without bubbling error dialog
      const cal = calculateMonthlyDispatchHours(monthIndex, year);
      return {
        standardHours: cal.totalHours,
        overtimeHours: 0,
        totalHours: cal.totalHours,
        source: 'calendar',
        monthName: MONTH_NAMES[monthIndex],
        year: year,
        tabFound: false,
        tabName: '',
        permissionError: true,
        permissionErrorMessage: permErr.message,
        timesheetStaff: [],
        fixedStaff: [{ name: 'Nour', regularHours: nourCalc.hours, overtimeHours: 0, total: nourCalc.hours, notes: `${nourCalc.activeDays} weekdays × 9h` }],
        excludedStaff: []
      };
    }

    if (!timesheetSs) {
      const cal = calculateMonthlyDispatchHours(monthIndex, year);
      return {
        standardHours: cal.totalHours,
        overtimeHours: 0,
        totalHours: cal.totalHours,
        source: 'calendar',
        monthName: MONTH_NAMES[monthIndex],
        year: year,
        tabFound: false,
        tabName: '',
        permissionError: false,
        timesheetStaff: [],
        fixedStaff: [{ name: 'Nour', regularHours: nourCalc.hours, overtimeHours: 0, total: nourCalc.hours, notes: `${nourCalc.activeDays} weekdays × 9h` }],
        excludedStaff: []
      };
    }

    const monthName = MONTH_NAMES[monthIndex];
    const targetNames = [
      `${monthName} ${year}`,
      `${monthName}`,
      `${monthName.slice(0, 3)} ${year}`,
      `${monthName.slice(0, 3)}`
    ];

    let sheet = null;
    const allSheets = timesheetSs.getSheets();
    for (const s of allSheets) {
      const sName = s.getName().trim().toLowerCase();
      for (const t of targetNames) {
        if (sName === t.toLowerCase() || sName.startsWith(monthName.toLowerCase())) {
          sheet = s;
          break;
        }
      }
      if (sheet) break;
    }

    if (!sheet) {
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
        permissionError: false,
        timesheetStaff: [],
        fixedStaff: [{ name: 'Nour', regularHours: nourCalc.hours, overtimeHours: 0, total: nourCalc.hours, notes: `${nourCalc.activeDays} weekdays × 9h` }],
        excludedStaff: []
      };
    }

    const data = sheet.getDataRange().getValues();
    const activeTimesheetNames = (activeCfg && activeCfg.ACTIVE_TIMESHEET_DISPATCHERS) || ['muhammad', 'mohamed', 'mariam', 'nourween'];
    const excludedNames = (activeCfg && activeCfg.EXCLUDED_DISPATCHER_NAMES) || ['mohanad', 'muhanad', 'abdulrahman', 'abdelrahman', 'abdo'];

    let timesheetStandardWT = 0;
    let timesheetOvertimeOT = 0;
    const timesheetStaff = [];
    const excludedStaff = [];

    let foundSummaryTable = false;
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
            const isActiveTimesheet = activeTimesheetNames.some(al => empLower.includes(al));

            const wt = parseNumericValueSafe(data[sr][wtCol]);
            const ot = parseNumericValueSafe(data[sr][otCol]);

            if (isActiveTimesheet && !isExcluded) {
              timesheetStandardWT += wt;
              timesheetOvertimeOT += ot;
              timesheetStaff.push({
                name: empRaw,
                regularHours: wt,
                overtimeHours: ot,
                total: wt + ot,
                source: 'Timesheet'
              });
              foundSummaryTable = true;
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
      if (foundSummaryTable) break;
    }

    if (!foundSummaryTable) {
      let totalRowIdx = -1;
      for (let r = data.length - 1; r >= 0; r--) {
        if (String(data[r][0] || '').trim().toLowerCase() === 'total') {
          totalRowIdx = r;
          break;
        }
      }

      if (totalRowIdx !== -1) {
        const headerRow = data[0];
        for (let c = 1; c < headerRow.length; c++) {
          const colHeader = String(headerRow[c] || '').trim().toLowerCase();
          const isExcluded = excludedNames.some(ex => colHeader.includes(ex));
          const isActiveTimesheet = activeTimesheetNames.some(al => colHeader.includes(al));

          if (isActiveTimesheet && !isExcluded) {
            const val = parseNumericValueSafe(data[totalRowIdx][c]);
            if (colHeader.includes('wt') || colHeader.includes('working')) {
              timesheetStandardWT += val;
            } else if (colHeader.includes('ot') || colHeader.includes('overtime')) {
              timesheetOvertimeOT += val;
            }
          }
        }
        if (timesheetStandardWT > 0) foundSummaryTable = true;
      }
    }

    if (foundSummaryTable && timesheetStandardWT > 0) {
      const totalStandard = timesheetStandardWT + nourCalc.hours;
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
        fixedStaff: [{ name: 'Nour', regularHours: nourCalc.hours, overtimeHours: 0, total: nourCalc.hours, notes: `${nourCalc.activeDays} weekdays × 9h/day` }],
        excludedStaff: excludedStaff
      };
    }

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
      permissionError: false,
      timesheetStaff: [],
      fixedStaff: [{ name: 'Nour', regularHours: nourCalc.hours, overtimeHours: 0, total: nourCalc.hours, notes: `${nourCalc.activeDays} weekdays × 9h/day` }],
      excludedStaff: []
    };

  } catch (err) {
    Logger.log(`Timesheet fetch error for ${monthIndex}/${year}: ${err.message}`);
    const cal = calculateMonthlyDispatchHours(monthIndex, year);
    return {
      standardHours: cal.totalHours,
      overtimeHours: 0,
      totalHours: cal.totalHours,
      source: 'calendar',
      monthName: MONTH_NAMES[monthIndex],
      year: year,
      tabFound: false,
      tabName: '',
      permissionError: true,
      permissionErrorMessage: err.message,
      timesheetStaff: [],
      fixedStaff: [{ name: 'Nour', regularHours: nourCalc.hours, overtimeHours: 0, total: nourCalc.hours, notes: `${nourCalc.activeDays} weekdays × 9h/day` }],
      excludedStaff: []
    };
  }
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
    calculateMonthlyDispatchHours,
    calculateNourFixedHours,
    fetchDispatcherHoursFromTimesheet
  };
}
