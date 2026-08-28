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
// TEST SUITE 2: 20-Hour Operating Window & Real-Time Concurrency Math (Neutral)
// ----------------------------------------------------
console.log('\n--- Test Suite 2: 20-Hour Window & Concurrency Math (Neutral) ---');

const operatingWindow = CONFIG.DAILY_OPERATING_WINDOW_HOURS; // 20h
const avgConcurrentDisp = CONFIG.AVG_CONCURRENT_DISPATCHERS; // 1.95

assert.strictEqual(operatingWindow, 20.0);
assert.strictEqual(avgConcurrentDisp, 1.95);

// Example test case: August 2026 with 909.0 driver hours over 31 days
const augustDriverHours = 909.0;
const augustDays = 31;

const estConcurrentCars = (augustDriverHours / augustDays) / operatingWindow; // (909 / 31) / 20 = 1.466 cars
const carsPerDispatcher = estConcurrentCars / avgConcurrentDisp; // 1.466 / 1.95 = 0.752 cars/disp

assert(estConcurrentCars > 1.4 && estConcurrentCars < 1.5);
assert(carsPerDispatcher > 0.7 && carsPerDispatcher < 0.8);

console.log(`  ✓ August 2026 Neutral Concurrency Metrics:`);
console.log(`    • Est. Concurrent Cars on Road: ${estConcurrentCars.toFixed(2)} cars`);
console.log(`    • Live Cars per On-Duty Dispatcher: ${carsPerDispatcher.toFixed(2)} cars/disp`);
console.log(`    • Operating Coverage: ${operatingWindow} hrs/day`);

// ----------------------------------------------------
// TEST SUITE 3: October 2025 Timesheet Math (with Mohanad)
// ----------------------------------------------------
console.log('\n--- Test Suite 3: October 2025 Timesheet Math (with Mohanad) ---');

// October 2025: 31 days total (23 weekdays)
const muhammadOct = calculateBaseContractHours('Muhammad', 9, 2025);
assert.strictEqual(muhammadOct.hours, 230.0);

const mariamOct = calculateBaseContractHours('Mariam', 9, 2025);
assert.strictEqual(mariamOct.hours, 220.0);

const mohanadOctWT = 30.5;
const nourOct = calculateNourFixedHours(9, 2025);
assert.strictEqual(nourOct.hours, 207.0);

const totalStandardWT = muhammadOct.hours + mariamOct.hours + mohanadOctWT + nourOct.hours; // 687.5
const totalOT = 29.0 + 34.0; // 63.0
const grandTotal = totalStandardWT + totalOT; // 750.5

assert.strictEqual(totalStandardWT, 687.5);
assert.strictEqual(totalOT, 63.0);
assert.strictEqual(grandTotal, 750.5);

console.log(`  ✓ October 2025 Verified with Mohanad: Total Regular WT: ${totalStandardWT} hrs | Total OT: ${totalOT} hrs | Grand Total: ${grandTotal} hrs`);

// ----------------------------------------------------
// TEST SUITE 4: Dispatcher Inclusion & Exclusion Rules
// ----------------------------------------------------
console.log('\n--- Test Suite 4: Dispatcher Inclusion & Exclusion Rules ---');

const activeList = CONFIG.ACTIVE_TIMESHEET_DISPATCHERS;
const excludedList = CONFIG.EXCLUDED_DISPATCHER_NAMES;

function isTimesheetDispatcher(name) {
  const lower = name.toLowerCase();
  const isExcluded = excludedList.some(ex => lower.includes(ex));
  const isActive = activeList.some(al => lower.includes(al));
  return isActive && !isExcluded;
}

assert.strictEqual(isTimesheetDispatcher('Muhammad WT'), true);
assert.strictEqual(isTimesheetDispatcher('Mariam WT'), true);
assert.strictEqual(isTimesheetDispatcher('Nourween WT'), true);
assert.strictEqual(isTimesheetDispatcher('Mohanad WT'), true);
assert.strictEqual(isTimesheetDispatcher('Fares WT'), false);
assert.strictEqual(isTimesheetDispatcher('Abdulrahman DR'), false);

console.log('  ✓ Verified: Muhammad, Mariam, Nourween, Mohanad are INCLUDED. Fares & Abdulrahman are EXCLUDED.');

console.log('\n🎉 ALL TEST SUITES PASSED!\n');
