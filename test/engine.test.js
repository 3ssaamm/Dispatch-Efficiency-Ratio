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
// TEST SUITE 2: Single-to-Full Name Normalization
// ----------------------------------------------------
console.log('\n--- Test Suite 2: Driver Name Normalization & Alias Merging ---');

assert.strictEqual(normalizeDriverName('Angel'), 'Angel Yoy');
assert.strictEqual(normalizeDriverName('angel'), 'Angel Yoy');
assert.strictEqual(normalizeDriverName('Angel Yoy'), 'Angel Yoy');

assert.strictEqual(normalizeDriverName('Brian'), 'Brian Macancela');
assert.strictEqual(normalizeDriverName('brian'), 'Brian Macancela');
assert.strictEqual(normalizeDriverName('Brian Macancela'), 'Brian Macancela');

assert.strictEqual(normalizeDriverName('Nikolay'), 'Nikolay Iankov');
assert.strictEqual(normalizeDriverName('Biaoming'), 'Biaoming Feng');
assert.strictEqual(normalizeDriverName('Oumarou'), 'Oumarou Amadou');
assert.strictEqual(normalizeDriverName('Koba'), 'Koba Svanadze');

// Dynamic alias map resolution test
const dynamicMap = { 'jake': 'Jake Sully' };
assert.strictEqual(normalizeDriverName('Jake', dynamicMap), 'Jake Sully');
console.log('  ✓ Single name aliases correctly resolved to canonical full names');

// ----------------------------------------------------
// TEST SUITE 3: Multi-Month Merging Test (User Screenshot Case)
// ----------------------------------------------------
console.log('\n--- Test Suite 3: Multi-Month Merging Simulation (User Screenshot Case) ---');

const userScreenshotMonths = [
  {
    monthName: 'MonthA',
    year: 2025,
    drivers: [
      { driverName: 'Angel', hours: 811.5, trips: 1000 },
      { driverName: 'Brian', hours: 555.0, trips: 700 },
      { driverName: 'Nikolay', hours: 362.5, trips: 450 },
      { driverName: 'Biaoming', hours: 205.0, trips: 250 },
      { driverName: 'Oumarou', hours: 182.5, trips: 220 }
    ]
  },
  {
    monthName: 'MonthB',
    year: 2026,
    drivers: [
      { driverName: 'Angel Yoy', hours: 76.5, trips: 95 },
      { driverName: 'Brian Macancela', hours: 163.5, trips: 210 },
      { driverName: 'Nikolay Iankov', hours: 108.0, trips: 135 },
      { driverName: 'Biaoming Feng', hours: 73.0, trips: 90 },
      { driverName: 'Oumarou Amadou', hours: 50.0, trips: 60 }
    ]
  }
];

const mergedResults = aggregateCumulativeDriverData(userScreenshotMonths);
console.log(`  Input: 10 driver rows across 2 months -> Merged Unique Drivers: ${mergedResults.length}`);
assert.strictEqual(mergedResults.length, 5, 'Must merge down to exactly 5 unique drivers');

const angel = mergedResults.find(d => d.driverName === 'Angel Yoy');
assert.notStrictEqual(angel, undefined);
assert.strictEqual(angel.monthsActive, 2);
assert.strictEqual(angel.totalHours, 888.0, 'Angel total hours must equal 811.5 + 76.5 = 888.0');
console.log(`  ✓ Angel + Angel Yoy -> Merged: "Angel Yoy" (2 months, 888.0 hrs)`);

const brian = mergedResults.find(d => d.driverName === 'Brian Macancela');
assert.notStrictEqual(brian, undefined);
assert.strictEqual(brian.monthsActive, 2);
assert.strictEqual(brian.totalHours, 718.5, 'Brian total hours must equal 555.0 + 163.5 = 718.5');
console.log(`  ✓ Brian + Brian Macancela -> Merged: "Brian Macancela" (2 months, 718.5 hrs)`);

const nikolay = mergedResults.find(d => d.driverName === 'Nikolay Iankov');
assert.notStrictEqual(nikolay, undefined);
assert.strictEqual(nikolay.monthsActive, 2);
assert.strictEqual(nikolay.totalHours, 470.5, 'Nikolay total hours must equal 362.5 + 108.0 = 470.5');
console.log(`  ✓ Nikolay + Nikolay Iankov -> Merged: "Nikolay Iankov" (2 months, 470.5 hrs)`);

const biaoming = mergedResults.find(d => d.driverName === 'Biaoming Feng');
assert.notStrictEqual(biaoming, undefined);
assert.strictEqual(biaoming.monthsActive, 2);
assert.strictEqual(biaoming.totalHours, 278.0, 'Biaoming total hours must equal 205.0 + 73.0 = 278.0');
console.log(`  ✓ Biaoming + Biaoming Feng -> Merged: "Biaoming Feng" (2 months, 278.0 hrs)`);

console.log('\n🎉 ALL TEST SUITES PASSED!\n');
