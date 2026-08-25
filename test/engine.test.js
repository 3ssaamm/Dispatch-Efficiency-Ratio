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
  { input: '2025 - January', fallback: 2026, expectedMonth: 0, expectedYear: 2025 }
];

testCasesParser.forEach((tc, idx) => {
  const res = parseMonthAndYear(tc.input, tc.fallback);
  assert.strictEqual(res.monthIndex, tc.expectedMonth);
  assert.strictEqual(res.year, tc.expectedYear);
  console.log(`  ✓ Case ${idx + 1}: "${tc.input}" -> ${res.monthName} ${res.year}`);
});

// ----------------------------------------------------
// TEST SUITE 2: Dispatcher Labor Hours Calculation
// ----------------------------------------------------
console.log('\n--- Test Suite 2: Dispatcher Labor Hours Calculation ---');

const aug2026 = calculateMonthlyDispatchHours(7, 2026);
assert.strictEqual(aug2026.daysInMonth, 31);
assert.strictEqual(aug2026.totalHours, 819);
console.log('  ✓ August 2026 calendar math verified: exactly 819.0 hours (21 active weekdays * 39h/day)');

// ----------------------------------------------------
// TEST SUITE 3: Daily Driver Aggregation Verification
// ----------------------------------------------------
console.log('\n--- Test Suite 3: Driver Daily Record Aggregation ---');

// Simulated daily raw rows from source spreadsheet
const sampleDailyRows = [
  { driver: 'Brian Macancela', hours: 6.0, trips: 5 },
  { driver: 'Koba Svanadze', hours: 5.5, trips: 5 },
  { driver: 'Nikolay Iankov', hours: 10.0, trips: 14 },
  { driver: 'Brian Macancela', hours: 10.0, trips: 13 },
  { driver: 'Koba Svanadze', hours: 15.0, trips: 19 },
  { driver: 'Nikolay Iankov', hours: 9.5, trips: 13 },
  { driver: 'brian macancela', hours: 9.0, trips: 14 }, // Test case-insensitivity
  { driver: 'Total Fleet Sum', hours: 65.0, trips: 70 } // Test summary skip
];

const driverMap = {};
for (const row of sampleDailyRows) {
  if (isSummaryRowName(row.driver)) continue;
  const key = row.driver.toLowerCase();
  if (!driverMap[key]) {
    driverMap[key] = { driverName: row.driver, hours: 0, trips: 0, count: 0 };
  }
  driverMap[key].hours += row.hours;
  driverMap[key].trips += row.trips;
  driverMap[key].count += 1;
}

const uniqueList = Object.values(driverMap);
uniqueList.sort((a, b) => b.hours - a.hours);

console.log(`  Raw daily entries: ${sampleDailyRows.length - 1} -> Aggregated unique drivers: ${uniqueList.length}`);
assert.strictEqual(uniqueList.length, 3, 'Must aggregate down to 3 unique drivers');

const brian = uniqueList.find(d => d.driverName.toLowerCase() === 'brian macancela');
assert.strictEqual(brian.hours, 25.0, 'Brian Macancela total hours must equal 6.0 + 10.0 + 9.0 = 25.0');
assert.strictEqual(brian.trips, 32, 'Brian Macancela total trips must equal 5 + 13 + 14 = 32');
assert.strictEqual(brian.count, 3, 'Brian Macancela appeared in 3 daily entries');
console.log(`  ✓ Brian Macancela: 3 daily entries aggregated -> ${brian.hours} hrs, ${brian.trips} trips`);

const koba = uniqueList.find(d => d.driverName.toLowerCase() === 'koba svanadze');
assert.strictEqual(koba.hours, 20.5);
assert.strictEqual(koba.trips, 24);
console.log(`  ✓ Koba Svanadze: 2 daily entries aggregated -> ${koba.hours} hrs, ${koba.trips} trips`);

const nikolay = uniqueList.find(d => d.driverName.toLowerCase() === 'nikolay iankov');
assert.strictEqual(nikolay.hours, 19.5);
assert.strictEqual(nikolay.trips, 27);
console.log(`  ✓ Nikolay Iankov: 2 daily entries aggregated -> ${nikolay.hours} hrs, ${nikolay.trips} trips`);

console.log('\n🎉 ALL TEST SUITES PASSED!\n');
