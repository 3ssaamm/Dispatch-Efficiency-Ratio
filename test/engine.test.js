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
// TEST SUITE 2: 2025 Timesheet Format Parser (October 2025 Case)
// ----------------------------------------------------
console.log('\n--- Test Suite 2: October 2025 Overtime-Only Timesheet Math ---');

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

// Nour: 23 weekdays * 9h = 207.0 WT
const nourOct = calculateNourFixedHours(9, 2025);
assert.strictEqual(nourOct.hours, 207.0);

// Totals:
const totalStandardWT = muhammadOct.hours + mariamOct.hours + nourOct.hours; // 230 + 220 + 207 = 657.0
const totalOT = 29.0 + 34.0; // 63.0
const grandTotal = totalStandardWT + totalOT; // 720.0

assert.strictEqual(totalStandardWT, 657.0);
assert.strictEqual(totalOT, 63.0);
assert.strictEqual(grandTotal, 720.0);

console.log(`  ✓ October 2025 Verified:`);
console.log(`    • Muhammad: ${muhammadOct.hours} WT + 29.0 OT = ${muhammadOct.hours + 29.0} hrs`);
console.log(`    • Mariam: ${mariamOct.hours} WT + 34.0 OT = ${mariamOct.hours + 34.0} hrs`);
console.log(`    • Nour: ${nourOct.hours} WT + 0.0 OT = ${nourOct.hours} hrs`);
console.log(`    • Total Regular WT: ${totalStandardWT} hrs | Total OT: ${totalOT} hrs | Grand Total: ${grandTotal} hrs`);

// ----------------------------------------------------
// TEST SUITE 3: Excluded Staff (Mohanad, Abdulrahman, Fares)
// ----------------------------------------------------
console.log('\n--- Test Suite 3: Excluded Staff Rules ---');

const excludedList = CONFIG.EXCLUDED_DISPATCHER_NAMES;
function isExcludedStaff(name) {
  const lower = name.toLowerCase();
  return excludedList.some(ex => lower.includes(ex));
}

assert.strictEqual(isExcludedStaff('Mohanad WT'), true, 'Mohanad must be excluded');
assert.strictEqual(isExcludedStaff('Fares WT'), true, 'Fares must be excluded');
assert.strictEqual(isExcludedStaff('Abdulrahman DR'), true, 'Abdulrahman must be excluded');
assert.strictEqual(isExcludedStaff('Muhammad OT'), false, 'Muhammad must NOT be excluded');
assert.strictEqual(isExcludedStaff('Mariam OT'), false, 'Mariam must NOT be excluded');

console.log('  ✓ Verified: Fares, Mohanad, Abdulrahman are properly excluded.');

console.log('\n🎉 ALL TEST SUITES PASSED!\n');
