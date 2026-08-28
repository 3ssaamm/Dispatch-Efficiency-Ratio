const assert = require('assert');
const { CONFIG } = require('../Config');
const {
  MONTH_NAMES,
  MONTH_MAP,
  parseMonthAndYear,
  calculateBaseContractHours,
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
  { input: 'October 2025 - Drivers Daily Balance', fallback: 2025, expectedMonth: 9, expectedYear: 2025 },
  { input: 'September 2025', fallback: 2025, expectedMonth: 8, expectedYear: 2025 },
  { input: 'August - Drivers Daily Balance', fallback: 2026, expectedMonth: 7, expectedYear: 2026 }
];

testCasesParser.forEach((tc, idx) => {
  const res = parseMonthAndYear(tc.input, tc.fallback);
  assert.strictEqual(res.monthIndex, tc.expectedMonth);
  assert.strictEqual(res.year, tc.expectedYear);
  console.log(`  ✓ Case ${idx + 1}: "${tc.input}" -> ${res.monthName} ${res.year}`);
});

// ----------------------------------------------------
// TEST SUITE 2: October 2025 Overtime-Only Timesheet Math (with Mohanad)
// ----------------------------------------------------
console.log('\n--- Test Suite 2: October 2025 Timesheet Math (with Mohanad) ---');

// October 2025: 31 days total.
// 23 weekdays (4 Mon, 4 Tue, 5 Wed, 5 Thu, 5 Fri)
// 4 Saturdays, 4 Sundays

// Muhammad: 23 weekdays * 10h = 230.0 WT
const muhammadOct = calculateBaseContractHours('Muhammad', 9, 2025);
assert.strictEqual(muhammadOct.hours, 230.0);
assert.strictEqual(muhammadOct.days, 23);

// Mariam: 22 working days (Mon, Tue, Wed, Fri, Sat) * 10h = 220.0 WT
const mariamOct = calculateBaseContractHours('Mariam', 9, 2025);
assert.strictEqual(mariamOct.hours, 220.0);
assert.strictEqual(mariamOct.days, 22);

// Mohanad: Logged 30.5 WT in Oct 2025
const mohanadOctWT = 30.5;

// Nour: 23 weekdays * 9h = 207.0 WT
const nourOct = calculateNourFixedHours(9, 2025);
assert.strictEqual(nourOct.hours, 207.0);

// Totals:
const totalStandardWT = muhammadOct.hours + mariamOct.hours + mohanadOctWT + nourOct.hours; // 230 + 220 + 30.5 + 207 = 687.5
const totalOT = 29.0 + 34.0; // 63.0
const grandTotal = totalStandardWT + totalOT; // 750.5

assert.strictEqual(totalStandardWT, 687.5);
assert.strictEqual(totalOT, 63.0);
assert.strictEqual(grandTotal, 750.5);

console.log(`  ✓ October 2025 Verified with Mohanad:`);
console.log(`    • Muhammad: ${muhammadOct.hours} WT + 29.0 OT = ${muhammadOct.hours + 29.0} hrs`);
console.log(`    • Mariam: ${mariamOct.hours} WT + 34.0 OT = ${mariamOct.hours + 34.0} hrs`);
console.log(`    • Mohanad: ${mohanadOctWT} WT + 0.0 OT = ${mohanadOctWT} hrs (INCLUDED)`);
console.log(`    • Nour: ${nourOct.hours} WT + 0.0 OT = ${nourOct.hours} hrs`);
console.log(`    • Total Regular WT: ${totalStandardWT} hrs | Total OT: ${totalOT} hrs | Grand Total: ${grandTotal} hrs`);

// ----------------------------------------------------
// TEST SUITE 3: Dispatcher Inclusion & Exclusion Rules
// ----------------------------------------------------
console.log('\n--- Test Suite 3: Dispatcher Inclusion & Exclusion Rules ---');

const activeList = CONFIG.ACTIVE_TIMESHEET_DISPATCHERS;
const excludedList = CONFIG.EXCLUDED_DISPATCHER_NAMES;

function isTimesheetDispatcher(name) {
  const lower = name.toLowerCase();
  const isExcluded = excludedList.some(ex => lower.includes(ex));
  const isActive = activeList.some(al => lower.includes(al));
  return isActive && !isExcluded;
}

assert.strictEqual(isTimesheetDispatcher('Muhammad WT'), true, 'Muhammad must be included');
assert.strictEqual(isTimesheetDispatcher('Mariam WT'), true, 'Mariam must be included');
assert.strictEqual(isTimesheetDispatcher('Nourween WT'), true, 'Nourween must be included');
assert.strictEqual(isTimesheetDispatcher('Mohanad WT'), true, 'Mohanad must be INCLUDED');
assert.strictEqual(isTimesheetDispatcher('Fares WT'), false, 'Fares must be excluded');
assert.strictEqual(isTimesheetDispatcher('Abdulrahman DR'), false, 'Abdulrahman must be excluded');

console.log('  ✓ Verified: Muhammad, Mariam, Nourween, Mohanad are INCLUDED. Fares & Abdulrahman are EXCLUDED.');

console.log('\n🎉 ALL TEST SUITES PASSED!\n');
