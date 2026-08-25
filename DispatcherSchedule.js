/**
 * @fileoverview Dispatcher labor hours calculation engine.
 * Computes exact monthly dispatch hours by evaluating actual calendar days
 * of any target month/year against individual dispatcher shift configurations.
 */

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
 * E.g., "August", "August 2026", "2025 - August", "Aug 2025"
 * @param {string} monthStr - Month text
 * @param {number} [fallbackYear] - Default year if not found in string
 * @returns {{ monthIndex: number, monthName: string, year: number }}
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

  // Extract alphabetical words to match month name
  const words = clean.toLowerCase().replace(/[^a-z]/g, ' ').split(/\s+/).filter(Boolean);
  let monthIndex = -1;

  for (const word of words) {
    if (MONTH_MAP.hasOwnProperty(word)) {
      monthIndex = MONTH_MAP[word];
      break;
    }
  }

  // If not matched, try substring search
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
    monthIndex = new Date().getMonth(); // default to current
  }

  return {
    monthIndex: monthIndex,
    monthName: MONTH_NAMES[monthIndex],
    year: year
  };
}

/**
 * Calculates total standard dispatch labor hours for a specific month and year.
 * Evaluates every calendar day of the month against weekday shift schedules.
 * 
 * @param {number} monthIndex - 0-indexed month (0 = January, 11 = December)
 * @param {number} year - 4-digit calendar year (e.g. 2025, 2026)
 * @param {Object} [customDailyHours] - Optional map of day-of-week to hours (0=Sun..6=Sat)
 * @returns {{ totalHours: number, daysInMonth: number, breakdown: Object, dailyLog: Array }}
 */
function calculateMonthlyDispatchHours(monthIndex, year, customDailyHours) {
  const schedule = customDailyHours || (typeof CONFIG !== 'undefined' ? CONFIG.DISPATCH_HOURS_BY_DAY_OF_WEEK : {
    0: 0,
    1: 39,
    2: 39,
    3: 39,
    4: 39,
    5: 39,
    6: 0
  });

  // Calculate days in the target month (handles leap years accurately)
  const daysInMonth = new Date(year, monthIndex + 1, 0).getDate();
  let totalHours = 0;
  const dayCounts = { 0: 0, 1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0 };
  const dailyLog = [];

  for (let day = 1; day <= daysInMonth; day++) {
    const dateObj = new Date(year, monthIndex, day);
    const dayOfWeek = dateObj.getDay(); // 0 = Sun, 1 = Mon, ..., 6 = Sat
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

// Node.js module export for testing
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    MONTH_NAMES,
    MONTH_MAP,
    parseMonthAndYear,
    calculateMonthlyDispatchHours
  };
}
