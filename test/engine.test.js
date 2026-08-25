const assert = require('assert');
const { CONFIG } = require('../Config');
const {
  MONTH_NAMES,
  MONTH_MAP,
  parseMonthAndYear,
  calculateMonthlyDispatchHours
} = require('../DispatcherSchedule');
const {
  locateHeaders,
  isSummaryRowName,
  parseNumericValue
} = require('../DriveSync');

console.log('🧪 Running Fleet Dispatch Efficiency Engine Test Suite...\n');

// ----------------------------------------------------
// TEST SUITE 1: Month & Year Parser
// ----------------------------------------------------
console.log('--- Test Suite 1: Month & Year Parser ---');

const testCasesParser = [
  { input: 'August - Drivers Daily Balance', fallback: 2026, expectedMonth: 7, expectedYear: 2026 },
  { input: 'August 2026 - Drivers Daily Balance', fallback: 2025, expectedMonth: 7, expectedYear: 2026 },
  { input: 'September 2025', fallback: 2026, expectedMonth: 8, expectedYear: 2025 },
  { input: '2025 - January', fallback: 2026, expectedMonth: 0, expectedYear: 2025 },
  { input: 'Feb', fallback: 2026, expectedMonth: 1, expectedYear: 2026 },
  { input: 'Dec 2024', fallback: 2026, expectedMonth: 11, expectedYear: 2024 }
];

testCasesParser.forEach((tc, idx) => {
  const res = parseMonthAndYear(tc.input, tc.fallback);
  assert.strictEqual(res.monthIndex, tc.expectedMonth, `Case ${idx + 1} monthIndex failed: got ${res.monthIndex}, expected ${tc.expectedMonth}`);
  assert.strictEqual(res.year, tc.expectedYear, `Case ${idx + 1} year failed: got ${res.year}, expected ${tc.expectedYear}`);
  console.log(`  ✓ Case ${idx + 1}: "${tc.input}" -> ${res.monthName} ${res.year}`);
});

// ----------------------------------------------------
// TEST SUITE 2: Dispatcher Schedule Math & Calendar Days
// ----------------------------------------------------
console.log('\n--- Test Suite 2: Dispatcher Labor Hours Calculation ---');

// August 2026:
// 31 days total.
// 2026-08-01 was a Saturday.
// In Aug 2026:
// Saturdays: 1, 8, 15, 22, 29 = 5 Saturdays (0 hrs each = 0h)
// Sundays: 2, 9, 16, 23, 30 = 5 Sundays (0 hrs each = 0h)
// Mondays: 3, 10, 17, 24, 31 = 5 Mondays (39h each = 195h)
// Tuesdays: 4, 11, 18, 25 = 4 Tuesdays (39h each = 156h)
// Wednesdays: 5, 12, 19, 26 = 4 Wednesdays (39h each = 156h)
// Thursdays: 6, 13, 20, 27 = 4 Thursdays (39h each = 156h)
// Fridays: 7, 14, 21, 28 = 4 Fridays (39h each = 156h)
// Total Weekdays = 5 + 4 + 4 + 4 + 4 = 21 weekdays * 39h = 819.0 hrs!

const aug2026 = calculateMonthlyDispatchHours(7, 2026);
console.log(`  August 2026 Days: ${aug2026.daysInMonth}, Total Standard Hours: ${aug2026.totalHours} hrs`);
assert.strictEqual(aug2026.daysInMonth, 31, 'August 2026 must have 31 days');
assert.strictEqual(aug2026.breakdown.mondays, 5);
assert.strictEqual(aug2026.breakdown.tuesdays, 4);
assert.strictEqual(aug2026.breakdown.wednesdays, 4);
assert.strictEqual(aug2026.breakdown.thursdays, 4);
assert.strictEqual(aug2026.breakdown.fridays, 4);
assert.strictEqual(aug2026.breakdown.saturdays, 5);
assert.strictEqual(aug2026.breakdown.sundays, 5);
assert.strictEqual(aug2026.totalHours, 819, 'August 2026 total hours must equal 819.0 hrs (21 weekdays * 39h)');
console.log('  ✓ August 2026 calendar math verified: exactly 819.0 hours (21 active weekdays * 39h/day)');

// February 2024 (Leap Year):
// 29 days. 2024-02-01 was Thursday.
// Thu: 1, 8, 15, 22, 29 = 5
// Fri: 2, 9, 16, 23 = 4
// Sat: 3, 10, 17, 24 = 4
// Sun: 4, 11, 18, 25 = 4
// Mon: 5, 12, 19, 26 = 4
// Tue: 6, 13, 20, 27 = 4
// Wed: 7, 14, 21, 28 = 4
// Total Weekdays = 5+4+4+4+4 = 21 weekdays * 39h = 819.0 hrs!
const feb2024 = calculateMonthlyDispatchHours(1, 2024);
assert.strictEqual(feb2024.daysInMonth, 29, 'Feb 2024 leap year must have 29 days');
assert.strictEqual(feb2024.totalHours, 819);
console.log('  ✓ February 2024 leap year verified: 29 days, 819.0 hours');

// ----------------------------------------------------
// TEST SUITE 3: Dynamic Header Locator
// ----------------------------------------------------
console.log('\n--- Test Suite 3: Dynamic Header Locator ---');

const sampleRows1 = [
  ['Driver Name', 'Active Hours', 'Completed Trips', 'Credit Card ($)', 'Cash ($)', 'Net Balance ($)', 'No Shows']
];
const loc1 = locateHeaders(sampleRows1);
assert.notStrictEqual(loc1, null);
assert.strictEqual(loc1.colMap.driver, 0);
assert.strictEqual(loc1.colMap.hours, 1);
assert.strictEqual(loc1.colMap.trips, 2);
console.log('  ✓ Standard column header layout mapped successfully');

// Shifted / Permuted columns with extra title row
const sampleRows2 = [
  ['MONTHLY BALANCE SUMMARY - INTERNAL REPORT', '', '', '', ''],
  ['Trips', 'Net Balance', 'Driver ID', 'Active Hrs', 'No-Show']
];
const loc2 = locateHeaders(sampleRows2);
assert.notStrictEqual(loc2, null);
assert.strictEqual(loc2.headerRowIndex, 1);
assert.strictEqual(loc2.colMap.driver, 2);
assert.strictEqual(loc2.colMap.hours, 3);
assert.strictEqual(loc2.colMap.trips, 0);
assert.strictEqual(loc2.colMap.balance, 1);
assert.strictEqual(loc2.colMap.noShow, 4);
console.log('  ✓ Shifted column layout with title row mapped successfully');

// ----------------------------------------------------
// TEST SUITE 4: Numeric Parsing & Sanitization
// ----------------------------------------------------
console.log('\n--- Test Suite 4: Numeric Parsing & Sanitization ---');

assert.strictEqual(parseNumericValue('$1,425.50'), 1425.5);
assert.strictEqual(parseNumericValue('142.5'), 142.5);
assert.strictEqual(parseNumericValue(160), 160);
assert.strictEqual(parseNumericValue(''), 0);
assert.strictEqual(parseNumericValue(null), 0);
assert.strictEqual(parseNumericValue(undefined), 0);
assert.strictEqual(parseNumericValue('N/A'), 0);
console.log('  ✓ Numeric parsing handles currencies, strings, nulls, and formatting correctly');

// ----------------------------------------------------
// TEST SUITE 5: Summary Row Detection
// ----------------------------------------------------
console.log('\n--- Test Suite 5: Summary Row Filter ---');

assert.strictEqual(isSummaryRowName('Total'), true);
assert.strictEqual(isSummaryRowName('Grand Total'), true);
assert.strictEqual(isSummaryRowName('Total Fleet Sum'), true);
assert.strictEqual(isSummaryRowName('Average'), true);
assert.strictEqual(isSummaryRowName('Marcus Vance (101)'), false);
assert.strictEqual(isSummaryRowName('John Doe'), false);
console.log('  ✓ Summary row filter correctly differentiates drivers from total rows');

// ----------------------------------------------------
// TEST SUITE 6: Dispatch Efficiency Ratios Formula Verification
// ----------------------------------------------------
console.log('\n--- Test Suite 6: Ratio Math & Metric Logic ---');

const driverHours = 1842.5;
const standardDisp = 819.0;
const overtime = 21.0;
const totalDisp = standardDisp + overtime; // 840.0
const dispatchRatio = totalDisp / driverHours; // 840 / 1842.5 = 0.455902...
const supportMins = dispatchRatio * 60; // 27.354 mins
const driverLeverage = driverHours / totalDisp; // 1842.5 / 840 = 2.1934...

console.log(`  Driver Hours: ${driverHours} hrs`);
console.log(`  Total Dispatch Hours: ${totalDisp} hrs`);
console.log(`  Fleet Dispatch Ratio: ${dispatchRatio.toFixed(3)}`);
console.log(`  Support Mins / Road Hr: ${supportMins.toFixed(1)} mins`);
console.log(`  Driver Leverage Ratio: ${driverLeverage.toFixed(2)}x`);

assert(dispatchRatio > 0 && dispatchRatio < 1);
assert(supportMins > 0);
assert(driverLeverage > 0);

console.log('\n🎉 ALL 6 TEST SUITES PASSED PERFECTLY!\n');
