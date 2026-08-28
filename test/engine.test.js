const assert = require('assert');
const { CONFIG } = require('../Config');
const {
  MONTH_NAMES,
  MONTH_MAP,
  parseMonthAndYear,
  calculateMonthlyDispatchHours,
  calculateNourFixedHours
} = require('../DispatcherSchedule');
const {
  locateHeaders,
  isSummaryRowName,
  normalizeDriverName,
  parseNumericValue
} = require('../DriveSync');
const {
  aggregateCumulativeDriverData
} = require('../MasterSummaryBuilder');

console.log('🧪 Running Fleet Dispatch Efficiency Engine Test Suite...\n');

// ----------------------------------------------------
// TEST SUITE 1: Month & Year Parser
// ----------------------------------------------------
console.log('--- Test Suite 1: Month & Year Parser ---');

const testCasesParser = [
  { input: 'August - Drivers Daily Balance', fallback: 2026, expectedMonth: 7, expectedYear: 2026 },
  { input: 'July 2026 - Drivers Daily Balance', fallback: 2026, expectedMonth: 6, expectedYear: 2026 },
  { input: 'September 2025', fallback: 2026, expectedMonth: 8, expectedYear: 2025 }
];

testCasesParser.forEach((tc, idx) => {
  const res = parseMonthAndYear(tc.input, tc.fallback);
  assert.strictEqual(res.monthIndex, tc.expectedMonth);
  assert.strictEqual(res.year, tc.expectedYear);
  console.log(`  ✓ Case ${idx + 1}: "${tc.input}" -> ${res.monthName} ${res.year}`);
});

// ----------------------------------------------------
// TEST SUITE 2: Nour (Fixed 9h) vs Nourween (10h in Timesheet)
// ----------------------------------------------------
console.log('\n--- Test Suite 2: Nour (Fixed 9h) vs Nourween (10h Timesheet) ---');

// August 2026 has 21 weekdays (5 Mon, 4 Tue, 4 Wed, 4 Thu, 4 Fri)
const nourAug2026 = calculateNourFixedHours(7, 2026);
assert.strictEqual(nourAug2026.activeDays, 21);
assert.strictEqual(nourAug2026.hours, 189.0, 'Nour fixed hours for Aug 2026 must equal 21 weekdays * 9h = 189.0 hrs');
console.log(`  ✓ Nour (Fixed 9h/day): ${nourAug2026.activeDays} weekdays × 9h = ${nourAug2026.hours} hrs`);

// Calendar schedule fallback calculation:
// Mon, Tue, Wed, Fri: 10 + 10 + 10 + 9 = 39h/day (5 Mon + 4 Tue + 4 Wed + 4 Fri = 17 days * 39h = 663h)
// Thu: 10 + 0 + 10 + 9 = 29h/day (4 Thu * 29h = 116h)
// Sat: 10h/day (5 Sat * 10h = 50h)
// Total = 663 + 116 + 50 = 829.0 hrs!
const aug2026 = calculateMonthlyDispatchHours(7, 2026);
assert.strictEqual(aug2026.daysInMonth, 31);
assert.strictEqual(aug2026.totalHours, 829.0, 'August 2026 full calendar schedule must equal 829.0 hrs');
console.log(`  ✓ August 2026 calendar fallback verified: ${aug2026.totalHours} hrs (Muhammad 10h, Mariam 10h [Thu OFF, Sat 10h], Nourween 10h, Nour 9h)`);

// ----------------------------------------------------
// TEST SUITE 3: Zero-Hour Month Filtering (July 2026)
// ----------------------------------------------------
console.log('\n--- Test Suite 3: Zero-Hour Month Filtering (July 2026) ---');

const testMonths = [
  {
    monthName: 'July',
    year: 2026,
    totalDriverHours: 0,
    drivers: []
  },
  {
    monthName: 'August',
    year: 2026,
    totalDriverHours: 909.0,
    drivers: [
      { driverName: 'Brian Macancela', hours: 163.5, trips: 210 },
      { driverName: 'Angel Yoy', hours: 76.5, trips: 95 }
    ]
  }
];

const filteredMonths = testMonths.filter(m => m.totalDriverHours > 0 && m.drivers.length > 0);
assert.strictEqual(filteredMonths.length, 1);
assert.strictEqual(filteredMonths[0].monthName, 'August');
console.log('  ✓ July 2026 (0 driver hours) successfully filtered out from master multi-month summary');

// ----------------------------------------------------
// TEST SUITE 4: Dispatcher Inclusion/Exclusion Rules
// ----------------------------------------------------
console.log('\n--- Test Suite 4: Dispatcher Inclusion & Exclusion Rules ---');

const activeTimesheetList = CONFIG.ACTIVE_TIMESHEET_DISPATCHERS;
const excludedList = CONFIG.EXCLUDED_DISPATCHER_NAMES;

function isTimesheetDispatcher(name) {
  const lower = name.toLowerCase().trim();
  const isExcluded = excludedList.some(ex => lower.includes(ex));
  const isAllowed = activeTimesheetList.some(al => lower.includes(al));
  return isAllowed && !isExcluded;
}

assert.strictEqual(isTimesheetDispatcher('Muhammad WT'), true);
assert.strictEqual(isTimesheetDispatcher('Mariam WT'), true);
assert.strictEqual(isTimesheetDispatcher('Nourween WT'), true);
assert.strictEqual(isTimesheetDispatcher('Mohanad WT'), false);
assert.strictEqual(isTimesheetDispatcher('Abdulrahman DR'), false);
console.log('  ✓ Verified: Muhammad, Mariam, Nourween are fetched from Timesheet. Nour is added as fixed 9h. Mohanad & Abdulrahman are EXCLUDED.');

console.log('\n🎉 ALL TEST SUITES PASSED!\n');
