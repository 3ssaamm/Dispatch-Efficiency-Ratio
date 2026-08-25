/**
 * @fileoverview Configuration file for Fleet Dispatch Efficiency Engine.
 * Contains global settings, Drive folder definitions, shift configurations,
 * header mappings, and UI theme styling.
 */

const CONFIG = {
  // Master Google Drive Folder containing monthly balance files & subfolders (2025, 2026)
  DRIVE_FOLDER_ID: '1iDd2ME2b6coX8GO6Dkom5u9jldkRSwVg',
  DRIVE_FOLDER_URL: 'https://drive.google.com/drive/folders/1iDd2ME2b6coX8GO6Dkom5u9jldkRSwVg?usp=sharing',

  // Target sheet tab name inside each monthly workbook
  SOURCE_SHEET_NAME: 'Summary',

  // File naming regex / pattern
  // Matches "August - Drivers Daily Balance", "August 2026 - Drivers Daily Balance", etc.
  FILE_NAME_PATTERN: /^(.*)\s*-\s*Drivers Daily Balance$/i,

  // Dispatcher shift schedules and daily labor hours
  // 4 Dispatchers:
  // - Disp 1: 10h Mon-Fri (Sat/Sun OFF)
  // - Disp 2: 10h Mon-Fri (Sat/Sun OFF)
  // - Disp 3: 10h Mon-Fri (Sat/Sun OFF)
  // - Disp 4: 9h Mon-Fri (Sat/Sun OFF) [Per user schedule specification]
  // Sunday = 0, Monday = 1, Tuesday = 2, Wednesday = 3, Thursday = 4, Friday = 5, Saturday = 6
  DISPATCH_HOURS_BY_DAY_OF_WEEK: {
    0: 0,   // Sunday: Fleet-wide OFF
    1: 39,  // Monday: 10 + 10 + 10 + 9 = 39 hrs
    2: 39,  // Tuesday: 10 + 10 + 10 + 9 = 39 hrs
    3: 39,  // Wednesday: 10 + 10 + 10 + 9 = 39 hrs
    4: 39,  // Thursday: 10 + 10 + 10 + 9 = 39 hrs
    5: 39,  // Friday: 10 + 10 + 10 + 9 = 39 hrs
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
    fontFamily: 'Inter, Roboto, Arial, sans-serif',
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

  // Default Sheet Names
  SETTINGS_SHEET_NAME: 'Settings',
  LOG_SHEET_NAME: 'Execution Log'
};

// If in Node.js environment for testing
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { CONFIG };
}
