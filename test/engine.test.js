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
// TEST SUITE 3: Multi-Month Master Summary Cumulative Aggregation
// ----------------------------------------------------
console.log('\n--- Test Suite 3: Multi-Month Driver Cumulative Aggregation ---');

const mockMonths = [
  {
    monthName: 'August',
    year: 2026,
    drivers: [
      { driverName: 'Brian Macancela', hours: 140.0, trips: 180 },
      { driverName: 'Koba Svanadze', hours: 120.5, trips: 150 }
    ]
  },
  {
    monthName: 'September',
    year: 2026,
    drivers: [
      { driverName: 'Brian Macancela', hours: 160.0, trips: 210 },
      { driverName: 'Nikolay Iankov', hours: 110.0, trips: 130 }
    ]
  }
];

const cumulative = aggregateCumulativeDriverData(mockMonths);
console.log(`  Processed 2 months -> Cumulative unique drivers: ${cumulative.length}`);
assert.strictEqual(cumulative.length, 3, 'Must have 3 unique cumulative drivers');

const brianCum = cumulative.find(d => d.driverName === 'Brian Macancela');
assert.strictEqual(brianCum.monthsActive, 2);
assert.strictEqual(brianCum.totalHours, 300.0);
assert.strictEqual(brianCum.totalTrips, 390);
console.log(`  ✓ Brian Macancela cumulative: 2 months active, 300.0 hrs, 390 trips`);

const kobaCum = cumulative.find(d => d.driverName === 'Koba Svanadze');
assert.strictEqual(kobaCum.monthsActive, 1);
assert.strictEqual(kobaCum.totalHours, 120.5);
console.log(`  ✓ Koba Svanadze cumulative: 1 month active, 120.5 hrs`);

console.log('\n🎉 ALL TEST SUITES PASSED!\n');
