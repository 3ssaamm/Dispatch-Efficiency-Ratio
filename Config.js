/**
 * @fileoverview Configuration file for Fleet Dispatch Efficiency Engine.
 * Contains global settings, Drive folder definitions, shift configurations,
 * timesheet integration, header mappings, driver alias merging, UI theme styling,
 * and 20-hour operating window & real-time concurrency constants.
 */

const CONFIG = {
  // Master Google Drive Folder containing monthly balance files & subfolders (2025, 2026)
  DRIVE_FOLDER_ID: '1iDd2ME2b6coX8GO6Dkom5u9jldkRSwVg',
  DRIVE_FOLDER_URL: 'https://drive.google.com/drive/folders/1iDd2ME2b6coX8GO6Dkom5u9jldkRSwVg?usp=sharing',

  // Timesheet Spreadsheet for Exact Dispatcher Working Hours & Overtime
  WORKING_TIME_SPREADSHEET_ID: '1bDHi8a2TuopuWyM1vOH3cA-Vx3umsQqosipCWAhr9Fc',
  WORKING_TIME_SPREADSHEET_URL: 'https://docs.google.com/spreadsheets/d/1bDHi8a2TuopuWyM1vOH3cA-Vx3umsQqosipCWAhr9Fc/edit?gid=557642234#gid=557642234',

  // Target sheet tab name inside each monthly balance workbook
  SOURCE_SHEET_NAME: 'Summary',

  // File naming regex / pattern
  FILE_NAME_PATTERN: /^(.*)\s*-\s*Drivers Daily Balance$/i,

  // --------------------------------------------------------------------------
  // OPERATING WINDOW & REAL-TIME CONCURRENCY CONSTANTS
  // --------------------------------------------------------------------------
  // Dispatch operation window: 4:00 AM to 12:00 AM midnight (20 active hours/day)
  DAILY_OPERATING_WINDOW_HOURS: 20.0,

  // Average on-duty dispatchers active concurrently at any given hour:
  // 39 scheduled weekday hours / 20 operating window hours = 1.95 dispatchers
  AVG_CONCURRENT_DISPATCHERS: 1.95,

  // --------------------------------------------------------------------------
  // DISPATCHER ROSTER & SCHEDULE RULES
  // --------------------------------------------------------------------------
  // 1. In Timesheet:
  //    - Muhammad: 10h Mon-Fri (Sat/Sun OFF)
  //    - Mariam: 10h Mon-Wed, Fri, Sat (Thursday & Sunday OFF)
  //    - Nourween: 10h Mon-Fri (Sat/Sun OFF) [Active 2026]
  //    - Mohanad: Counted from Timesheet (WT + OT)
  // 2. Fixed (NOT in timesheet):
  //    - Nour: Fixed 9h/day Mon-Fri (Sat/Sun OFF)
  // 3. Excluded Staff:
  //    - Abdulrahman
  //    - Fares (Disregarded)
  // --------------------------------------------------------------------------
  ACTIVE_TIMESHEET_DISPATCHERS: ['muhammad', 'mohamed', 'mariam', 'nourween', 'mohanad', 'muhanad'],
  FIXED_SCHEDULE_DISPATCHERS: [
    { name: 'Nour', dailyHours: 9, workDays: [1, 2, 3, 4, 5], offDays: [0, 6], notes: 'Fixed 9h/day Mon-Fri (Not in timesheet)' }
  ],
  EXCLUDED_DISPATCHER_NAMES: ['abdulrahman', 'abdelrahman', 'abdo', 'fares'],

  // Calendar Shift Schedule (Fallback when timesheet is not present):
  // - Mon, Tue, Wed, Fri: Muhammad (10h) + Mariam (10h) + Nourween (10h) + Nour (9h) = 39.0 hrs/day
  // - Thursday: Muhammad (10h) + Mariam (OFF) + Nourween (10h) + Nour (9h) = 29.0 hrs/day
  // - Saturday: Muhammad (OFF) + Mariam (10h) + Nourween (OFF) + Nour (OFF) = 10.0 hrs/day
  // - Sunday: 0.0 hrs/day
  DISPATCH_HOURS_BY_DAY_OF_WEEK: {
    0: 0,   // Sunday: OFF
    1: 39,  // Monday: 10 + 10 + 10 + 9 = 39 hrs
    2: 39,  // Tuesday: 10 + 10 + 10 + 9 = 39 hrs
    3: 39,  // Wednesday: 10 + 10 + 10 + 9 = 39 hrs
    4: 29,  // Thursday: 10 + 0 (Mariam OFF) + 10 + 9 = 29 hrs
    5: 39,  // Friday: 10 + 10 + 10 + 9 = 39 hrs
    6: 10   // Saturday: 0 + 10 (Mariam) + 0 + 0 = 10 hrs
  },

  DISPATCHERS: [
    { id: 1, name: 'Muhammad', dailyHours: 10, source: 'Timesheet', workDays: [1, 2, 3, 4, 5], offDays: [0, 6], notes: '10h Mon-Fri (Sat/Sun OFF)' },
    { id: 2, name: 'Mariam', dailyHours: 10, source: 'Timesheet', workDays: [1, 2, 3, 5, 6], offDays: [0, 4], notes: '10h Mon-Wed, Fri, Sat (Thu & Sun OFF)' },
    { id: 3, name: 'Nourween', dailyHours: 10, source: 'Timesheet', workDays: [1, 2, 3, 4, 5], offDays: [0, 6], notes: '10h Mon-Fri (Sat/Sun OFF) [Active 2026]' },
    { id: 4, name: 'Mohanad', dailyHours: 'Logged', source: 'Timesheet', workDays: 'Logged', offDays: 'Logged', notes: 'Logged from Timesheet' },
    { id: 5, name: 'Nour', dailyHours: 9, source: 'Fixed Schedule', workDays: [1, 2, 3, 4, 5], offDays: [0, 6], notes: 'Fixed 9h/day Mon-Fri (Sat/Sun OFF)' }
  ],

  // Driver Name Alias Mapping: Automatically merges single first names into canonical full names
  DRIVER_ALIASES: {
    'angel': 'Angel Yoy',
    'brian': 'Brian Macancela',
    'nikolay': 'Nikolay Iankov',
    'biaoming': 'Biaoming Feng',
    'oumarou': 'Oumarou Amadou',
    'koba': 'Koba Svanadze',
    'prince': 'Prince Verma',
    'amadou': 'Amadou Diallo',
    'amdou': 'Amadou Diallo',
    'benjamin': 'Benjamin Douglass'
  },

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

  SETTINGS_SHEET_NAME: 'Settings',
  LOG_SHEET_NAME: 'Execution Log',
  AUDIT_SHEET_NAME: '🕒 Dispatcher Hours Audit'
};

// If in Node.js environment for testing
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { CONFIG };
}
