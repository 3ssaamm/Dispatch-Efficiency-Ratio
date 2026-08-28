/**
 * @fileoverview Master Summary Sheet Builder.
 * Consolidates all monthly balance sheets with month-over-month trend tables,
 * live dynamic formulas, aggregate KPI cards, 20h operating window concurrency metrics,
 * and an intelligently merged cumulative driver leaderboard.
 */

// Node.js fallback import for testing
let _CONFIG_MS = (typeof CONFIG !== 'undefined') ? CONFIG : null;
let _isSummaryRowName = (typeof isSummaryRowName === 'function') ? isSummaryRowName : null;
let _normalizeDriverName = (typeof normalizeDriverName === 'function') ? normalizeDriverName : null;

if (typeof require !== 'undefined') {
  try {
    if (!_CONFIG_MS) _CONFIG_MS = require('./Config').CONFIG;
    if (!_isSummaryRowName || !_normalizeDriverName) {
      const ds = require('./DriveSync');
      if (!_isSummaryRowName) _isSummaryRowName = ds.isSummaryRowName;
      if (!_normalizeDriverName) _normalizeDriverName = ds.normalizeDriverName;
    }
  } catch (e) {}
}

/**
 * Builds or refreshes the '📊 Master Summary' sheet tab.
 * 
 * @param {Array<Object>} [allMonthsData] - Array of month data bundles
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function buildMasterSummarySheet(allMonthsData, targetSpreadsheet) {
  const ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  const tabName = '📊 Master Summary';

  let sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    sheet = ss.getSheetByName('Master Summary');
  }

  if (sheet) {
    sheet.clear();
    sheet.clearConditionalFormatRules();
  } else {
    sheet = ss.insertSheet(tabName, 1);
  }

  let rawMonths = allMonthsData;
  if (!rawMonths || rawMonths.length === 0) {
    rawMonths = collectDataFromExistingAnalysisSheets(ss);
  }

  // Filter out any month with zero active driver hours (e.g. July 2026)
  const months = rawMonths.filter(m => (m.totalDriverHours > 0) && (m.drivers && m.drivers.length > 0));
  const excludedZeroMonths = rawMonths.filter(m => !m.totalDriverHours || m.totalDriverHours <= 0);
  const excludedNotice = excludedZeroMonths.length > 0
    ? ` | Excluded from summary (0 driver hrs): ${excludedZeroMonths.map(em => `${em.monthName} ${em.year}`).join(', ')}`
    : '';

  months.sort((a, b) => (a.year - b.year) || (a.monthIndex - b.monthIndex));

  const monthCount = months.length;
  const startMonthRow = 15;
  const endMonthRow = startMonthRow + Math.max(monthCount - 1, 0);
  const totalMonthRow = endMonthRow + 1;

  // Build Multi-Month Driver Cumulative Aggregation with Name Merging
  const cumulativeDrivers = aggregateCumulativeDriverData(months);
  const driverCount = cumulativeDrivers.length;
  const startDriverRow = totalMonthRow + 4;
  const endDriverRow = startDriverRow + Math.max(driverCount - 1, 0);
  const totalDriverRow = endDriverRow + 1;

  const totalRowsNeeded = Math.max(totalDriverRow + 5, 45);
  const totalCols = 14;
  const matrix = Array.from({ length: totalRowsNeeded }, () => Array(totalCols).fill(''));

  // Row 1: Title Banner
  matrix[0][0] = '🚀 FLEET DISPATCH EFFICIENCY — ALL-MONTHS MASTER EXECUTIVE SUMMARY';

  // Row 2: Subtitle
  matrix[1][0] = `Generated: ${new Date().toLocaleString()} | Active Months: ${monthCount} | 20h Daily Window (4am-12am) | Benchmark: 1:4 (4 cars/disp)${excludedNotice}`;

  // Rows 3-5: Aggregate Top Labor KPI Cards
  // Card 1: TOTAL FLEET DRIVER HOURS (Cols A:C)
  matrix[2][0] = 'TOTAL FLEET DRIVER HOURS';
  matrix[3][0] = monthCount > 0 ? `=SUM(C${startMonthRow}:C${endMonthRow})` : 0;
  matrix[4][0] = 'Active Road Hours (Active Months)';

  // Card 2: STANDARD DISPATCH HOURS (Cols D:F)
  matrix[2][3] = 'TOTAL STANDARD DISPATCH';
  matrix[3][3] = monthCount > 0 ? `=SUM(D${startMonthRow}:D${endMonthRow})` : 0;
  matrix[4][3] = 'Muhammad, Mariam, Nourween, Mohanad, Nour';

  // Card 3: OVERTIME HOURS (Cols G:I)
  matrix[2][6] = 'TOTAL OVERTIME / ADJ';
  matrix[3][6] = monthCount > 0 ? `=SUM(E${startMonthRow}:E${endMonthRow})` : 0;
  matrix[4][6] = 'Total Overtime Logged';

  // Card 4: TOTAL DISPATCH HOURS (Cols J:N)
  matrix[2][9] = 'TOTAL DISPATCH LABOR HOURS';
  matrix[3][9] = monthCount > 0 ? `=SUM(F${startMonthRow}:F${endMonthRow})` : 0;
  matrix[4][9] = 'Standard + Overtime';

  // Rows 6-8: Total Labor Overhead Ratio Cards
  // Card 5: OVERALL FLEET DISPATCH RATIO (Cols A:D)
  matrix[5][0] = 'OVERALL FLEET DISPATCH RATIO';
  matrix[6][0] = `=IF($A$4>0, $J$4/$A$4, 0)`;
  matrix[7][0] = 'Total Dispatch Hrs / Road Hour';

  // Card 6: OVERALL SUPPORT MINS / ROAD HR (Cols E:H)
  matrix[5][4] = 'OVERALL SUPPORT MINS / ROAD HR';
  matrix[6][4] = `=$A$7*60`;
  matrix[7][4] = 'Minutes of Dispatch per 1 Road Hr';

  // Card 7: OVERALL DRIVER LEVERAGE RATIO (Cols I:N)
  matrix[5][8] = 'OVERALL DRIVER LEVERAGE RATIO';
  matrix[6][8] = `=IF($J$4>0, $A$4/$J$4, 0)`;
  matrix[7][8] = 'Road Hours driven per 1 Dispatch Hr';

  // Rows 9-11: 20-Hour Operating Window & Real-Time Concurrency Summary Cards
  // Card 8: AVG REAL-TIME ACTIVE CARS (Cols A:D)
  matrix[8][0] = 'AVG REAL-TIME CONCURRENT CARS';
  matrix[9][0] = monthCount > 0 ? `=AVERAGE(K${startMonthRow}:K${endMonthRow})` : 0;
  matrix[10][0] = 'Estimated Active Cars in 20h Window';

  // Card 9: AVG LIVE CARS / ON-DUTY DISPATCHER (Cols E:H)
  matrix[8][4] = 'LIVE CARS / ON-DUTY DISPATCHER';
  matrix[9][4] = monthCount > 0 ? `=AVERAGE(L${startMonthRow}:L${endMonthRow})` : 0;
  matrix[10][4] = 'Real-Time Load (~1.95 Concurrent Disp)';

  // Card 10: FLEET DESK CAPACITY UTILIZATION (Cols I:N)
  matrix[8][8] = 'DESK CAPACITY UTILIZATION';
  matrix[9][8] = monthCount > 0 ? `=AVERAGE(M${startMonthRow}:M${endMonthRow})` : 0;
  matrix[10][8] = 'Vs 1:4 Optimal Benchmark (4 cars/disp)';

  // Row 13: Section 1 Header
  matrix[12][0] = '1. MONTH-OVER-MONTH FLEET DISPATCH & REAL-TIME CONCURRENCY COMPARISON';

  // Row 14: Month Table Column Headers
  matrix[13][0] = 'Month & Year';
  matrix[13][1] = 'Active Drivers';
  matrix[13][2] = 'Driver Hours';
  matrix[13][3] = 'Standard Disp';
  matrix[13][4] = 'Overtime';
  matrix[13][5] = 'Total Disp Hrs';
  matrix[13][6] = 'Dispatch Ratio';
  matrix[13][7] = 'Support Mins/Hr';
  matrix[13][8] = 'Driver Leverage';
  matrix[13][9] = 'Completed Trips';
  matrix[13][10] = 'Est Live Cars';
  matrix[13][11] = 'Cars / Disp';
  matrix[13][12] = 'Capacity %';
  matrix[13][13] = 'Trips / Disp Hr';

  // Month Table Rows
  if (monthCount > 0) {
    for (let i = 0; i < monthCount; i++) {
      const m = months[i];
      const r = startMonthRow + i;
      const rowIdx = r - 1;
      const daysInMonth = new Date(m.year, m.monthIndex + 1, 0).getDate();
      
      const tabTargetName = `${m.monthName} ${m.year}`;
      const targetSheet = ss.getSheetByName(tabTargetName) 
        || ss.getSheetByName(`${m.monthName} ${m.year} - Analysis`)
        || ss.getSheetByName(`${m.monthName} - Analysis`)
        || ss.getSheetByName(m.monthName);

      matrix[rowIdx][0] = `${m.monthName} ${m.year}`;
      matrix[rowIdx][1] = m.driverCount || (m.drivers ? m.drivers.length : 0);

      if (targetSheet) {
        const actualTab = targetSheet.getName();
        matrix[rowIdx][2] = `='${actualTab}'!A4`;
        matrix[rowIdx][3] = `='${actualTab}'!C4`;
        matrix[rowIdx][4] = `='${actualTab}'!E4`;
        matrix[rowIdx][5] = `='${actualTab}'!F4`;
        matrix[rowIdx][6] = `='${actualTab}'!A7`;
        matrix[rowIdx][7] = `='${actualTab}'!C7`;
        matrix[rowIdx][8] = `='${actualTab}'!E7`;
      } else {
        matrix[rowIdx][2] = m.totalDriverHours || 0;
        matrix[rowIdx][3] = m.standardDispatchHours || 0;
        matrix[rowIdx][4] = m.overtimeHours || 0;
        matrix[rowIdx][5] = `=D${r}+E${r}`;
        matrix[rowIdx][6] = `=IF(C${r}>0, F${r}/C${r}, 0)`;
        matrix[rowIdx][7] = `=G${r}*60`;
        matrix[rowIdx][8] = `=IF(F${r}>0, C${r}/F${r}, 0)`;
      }

      matrix[rowIdx][9] = m.totalTrips || 0;
      matrix[rowIdx][10] = `=IF(C${r}>0, (C${r}/${daysInMonth})/20, 0)`; // Est. Live Cars
      matrix[rowIdx][11] = `=IF(K${r}>0, K${r}/1.95, 0)`;              // Cars / On-Duty Disp
      matrix[rowIdx][12] = `=IF(L${r}>0, L${r}/4.0, 0)`;               // Capacity %
      matrix[rowIdx][13] = `=IF(F${r}>0, J${r}/F${r}, 0)`;              // Trips / Disp Hr
    }

    // Total Month Row
    const totMIdx = totalMonthRow - 1;
    matrix[totMIdx][0] = 'TOTAL / FLEET AVERAGE';
    matrix[totMIdx][1] = `=AVERAGE(B${startMonthRow}:B${endMonthRow})`;
    matrix[totMIdx][2] = `=SUM(C${startMonthRow}:C${endMonthRow})`;
    matrix[totMIdx][3] = `=SUM(D${startMonthRow}:D${endMonthRow})`;
    matrix[totMIdx][4] = `=SUM(E${startMonthRow}:E${endMonthRow})`;
    matrix[totMIdx][5] = `=SUM(F${startMonthRow}:F${endMonthRow})`;
    matrix[totMIdx][6] = `=IF(C${totalMonthRow}>0, F${totalMonthRow}/C${totalMonthRow}, 0)`;
    matrix[totMIdx][7] = `=G${totalMonthRow}*60`;
    matrix[totMIdx][8] = `=IF(F${totalMonthRow}>0, C${totalMonthRow}/F${totalMonthRow}, 0)`;
    matrix[totMIdx][9] = `=SUM(J${startMonthRow}:J${endMonthRow})`;
    matrix[totMIdx][10] = `=AVERAGE(K${startMonthRow}:K${endMonthRow})`;
    matrix[totMIdx][11] = `=AVERAGE(L${startMonthRow}:L${endMonthRow})`;
    matrix[totMIdx][12] = `=AVERAGE(M${startMonthRow}:M${endMonthRow})`;
    matrix[totMIdx][13] = `=IF(F${totalMonthRow}>0, J${totalMonthRow}/F${totalMonthRow}, 0)`;
  }

  // Section 2: Cumulative Driver Leaderboard
  const sec2TitleRow = totalMonthRow + 2;
  const sec2HeaderRow = sec2TitleRow + 1;
  const sec2TitleIdx = sec2TitleRow - 1;
  const sec2HeaderIdx = sec2HeaderRow - 1;

  matrix[sec2TitleIdx][0] = '2. CUMULATIVE DRIVER PERFORMANCE LEADERBOARD (ACTIVE MONTHS)';

  matrix[sec2HeaderIdx][0] = 'Driver Name';
  matrix[sec2HeaderIdx][1] = 'Months Active';
  matrix[sec2HeaderIdx][2] = 'Cumulative Hours';
  matrix[sec2HeaderIdx][3] = '% Fleet Share';
  matrix[sec2HeaderIdx][4] = 'Allocated Disp Hrs';
  matrix[sec2HeaderIdx][5] = 'Completed Trips';
  matrix[sec2HeaderIdx][6] = 'Trips / Driver Hr';
  matrix[sec2HeaderIdx][7] = 'Trips / Disp Hr';
  matrix[sec2HeaderIdx][8] = 'Avg Monthly Hrs';
  matrix[sec2HeaderIdx][9] = 'Avg Monthly Trips';
  matrix[sec2HeaderIdx][10] = 'Status';
  matrix[sec2HeaderIdx][11] = 'Notes';
  matrix[sec2HeaderIdx][12] = '';
  matrix[sec2HeaderIdx][13] = '';

  if (driverCount > 0) {
    for (let j = 0; j < driverCount; j++) {
      const cd = cumulativeDrivers[j];
      const dr = startDriverRow + j;
      const drIdx = dr - 1;

      matrix[drIdx][0] = cd.driverName;
      matrix[drIdx][1] = cd.monthsActive;
      matrix[drIdx][2] = cd.totalHours;
      matrix[drIdx][3] = `=IF($A$4>0, C${dr}/$A$4, 0)`;
      matrix[drIdx][4] = `=C${dr}*$A$7`;
      matrix[drIdx][5] = cd.totalTrips;
      matrix[drIdx][6] = `=IF(C${dr}>0, F${dr}/C${dr}, 0)`;
      matrix[drIdx][7] = `=IF(E${dr}>0, F${dr}/E${dr}, 0)`;
      matrix[drIdx][8] = `=IF(B${dr}>0, C${dr}/B${dr}, 0)`;
      matrix[drIdx][9] = `=IF(B${dr}>0, F${dr}/B${dr}, 0)`;
      matrix[drIdx][10] = cd.monthsActive >= Math.max(monthCount - 1, 1) ? '⭐ Core Driver' : 'Active';
      matrix[drIdx][11] = `${cd.monthsActive} of ${monthCount} months`;
    }

    // Cumulative Driver Total Row
    const totDIdx = totalDriverRow - 1;
    matrix[totDIdx][0] = 'TOTAL FLEET CUMULATIVE';
    matrix[totDIdx][1] = monthCount;
    matrix[totDIdx][2] = `=SUM(C${startDriverRow}:C${endDriverRow})`;
    matrix[totDIdx][3] = `=SUM(D${startDriverRow}:D${endDriverRow})`;
    matrix[totDIdx][4] = `=SUM(E${startDriverRow}:E${endDriverRow})`;
    matrix[totDIdx][5] = `=SUM(F${startDriverRow}:F${endDriverRow})`;
    matrix[totDIdx][6] = `=IF(C${totalDriverRow}>0, F${totalDriverRow}/C${totalDriverRow}, 0)`;
    matrix[totDIdx][7] = `=IF(E${totalDriverRow}>0, F${totalDriverRow}/E${totalDriverRow}, 0)`;
    matrix[totDIdx][8] = `=IF(B${totalDriverRow}>0, C${totalDriverRow}/B${totalDriverRow}, 0)`;
    matrix[totDIdx][9] = `=IF(B${totalDriverRow}>0, F${totalDriverRow}/B${totalDriverRow}, 0)`;
    matrix[totDIdx][10] = 'Fleet Total';
    matrix[totDIdx][11] = '';
  }

  sheet.getRange(1, 1, totalRowsNeeded, totalCols).setValues(matrix);

  formatMasterSummarySheet(sheet, {
    monthCount: monthCount,
    startMonthRow: startMonthRow,
    endMonthRow: endMonthRow,
    totalMonthRow: totalMonthRow,
    driverCount: driverCount,
    sec2TitleRow: sec2TitleRow,
    sec2HeaderRow: sec2HeaderRow,
    startDriverRow: startDriverRow,
    endDriverRow: endDriverRow,
    totalDriverRow: totalDriverRow
  });

  return sheet;
}

/**
 * Aggregates all driver data across multiple active months into cumulative totals.
 */
function aggregateCumulativeDriverData(months) {
  const isSummaryFn = (typeof isSummaryRowName === 'function') ? isSummaryRowName : (_isSummaryRowName || (() => false));
  const normDriverFn = (typeof normalizeDriverName === 'function') ? normalizeDriverName : _normalizeDriverName;

  const globalFullNameMap = {};
  for (const m of months) {
    const drivers = m.drivers || [];
    for (const d of drivers) {
      const name = String(d.driverName || '').trim();
      const parts = name.split(/\s+/);
      if (parts.length >= 2 && !isSummaryFn(name)) {
        const first = parts[0].toLowerCase();
        if (!globalFullNameMap[first]) {
          globalFullNameMap[first] = name;
        }
      }
    }
  }

  const driverMap = {};

  for (const m of months) {
    if (!m.totalDriverHours || m.totalDriverHours <= 0) continue;

    const drivers = m.drivers || [];
    for (const d of drivers) {
      const rawName = String(d.driverName || '').trim();
      if (!rawName || isSummaryFn(rawName)) continue;

      const canonical = typeof normDriverFn === 'function' 
        ? normDriverFn(rawName, globalFullNameMap)
        : (globalFullNameMap[rawName.toLowerCase()] || rawName);

      const key = canonical.toLowerCase().trim();

      if (!driverMap[key]) {
        driverMap[key] = {
          driverName: canonical,
          monthsActive: 0,
          totalHours: 0,
          totalTrips: 0,
          totalCredit: 0,
          totalCash: 0,
          totalBalance: 0,
          totalNoShow: 0
        };
      }
      driverMap[key].monthsActive += 1;
      driverMap[key].totalHours += d.hours;
      driverMap[key].totalTrips += d.trips;
      driverMap[key].totalCredit += (d.credit || 0);
      driverMap[key].totalCash += (d.cash || 0);
      driverMap[key].totalBalance += (d.balance || 0);
      driverMap[key].totalNoShow += (d.noShow || 0);
    }
  }

  const list = Object.values(driverMap);
  for (const item of list) {
    item.totalHours = Math.round(item.totalHours * 10) / 10;
  }

  list.sort((a, b) => b.totalHours - a.totalHours || a.driverName.localeCompare(b.driverName));
  return list;
}

/**
 * Discovers existing monthly tabs from the active workbook.
 */
function collectDataFromExistingAnalysisSheets(ss) {
  const sheets = ss.getSheets();
  const list = [];
  const systemNames = ['master summary', 'read me', 'user guide', 'settings', 'execution log', 'dispatcher hours audit', 'sheet1'];

  for (const s of sheets) {
    const name = s.getName();
    const lowerName = name.toLowerCase();

    const isSystem = systemNames.some(sys => lowerName.includes(sys));
    if (isSystem) continue;

    const isMonth = MONTH_NAMES.some(m => lowerName.includes(m.toLowerCase()));
    if (!isMonth) continue;

    const monthPart = name.replace('- Analysis', '').trim();
    const parsed = parseMonthAndYear(monthPart);
    const schedule = calculateMonthlyDispatchHours(parsed.monthIndex, parsed.year);

    const lastRow = s.getLastRow();
    const drivers = [];
    let totalDriverHours = 0;
    let totalTrips = 0;

    // Data rows start at row 14 (or row 11 in older templates)
    const dataStart = 14;
    if (lastRow >= dataStart) {
      const vals = s.getRange(dataStart, 1, lastRow - dataStart, 7).getValues();
      for (const row of vals) {
        const dName = String(row[0] || '').trim();
        if (!dName || isSummaryRowName(dName)) continue;
        const hrs = parseNumericValue(row[1]);
        const trips = parseNumericValue(row[4]);
        drivers.push({ driverName: dName, hours: hrs, trips: trips });
        totalDriverHours += hrs;
        totalTrips += trips;
      }
    }

    list.push({
      monthName: parsed.monthName,
      monthIndex: parsed.monthIndex,
      year: parsed.year,
      fileName: `${parsed.monthName} - Drivers Daily Balance`,
      totalDriverHours: totalDriverHours,
      totalTrips: totalTrips,
      standardDispatchHours: schedule.totalHours,
      driverCount: drivers.length,
      drivers: drivers
    });
  }

  return list;
}

/**
 * Styling and formatting engine for Master Summary sheet.
 */
function formatMasterSummarySheet(sheet, cfg) {
  const activeCfg = (typeof CONFIG !== 'undefined') ? CONFIG : _CONFIG_MS;
  const theme = (activeCfg && activeCfg.THEME) ? activeCfg.THEME : {
    headerBg: '#1e293b',
    headerColor: '#ffffff',
    zebraBg: '#f1f5f9'
  };

  sheet.getRange('A1:N1').merge()
    .setBackground(theme.headerBg)
    .setFontColor(theme.headerColor)
    .setFontSize(13)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 38);

  sheet.getRange('A2:N2').merge()
    .setBackground('#334155')
    .setFontColor('#cbd5e1')
    .setFontSize(9)
    .setFontStyle('italic')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(2, 22);

  // Row 1 KPI Cards
  sheet.getRange('A3:C3').merge().setValue('TOTAL FLEET DRIVER HOURS');
  sheet.getRange('A4:C4').merge();
  sheet.getRange('A5:C5').merge().setValue('Active Months Combined');
  formatKpiCard(sheet, 'A3:C5', 'A4:C4', '#,##0.0 "hrs"', '#1e3a8a');

  sheet.getRange('D3:F3').merge().setValue('TOTAL STANDARD DISPATCH');
  sheet.getRange('D4:F4').merge();
  sheet.getRange('D5:F5').merge().setValue('Muhammad, Mariam, Nourween, Mohanad, Nour');
  formatKpiCard(sheet, 'D3:F5', 'D4:F4', '#,##0.0 "hrs"', '#1e293b');

  sheet.getRange('G3:I3').merge().setValue('TOTAL OVERTIME / ADJ');
  sheet.getRange('G4:I4').merge();
  sheet.getRange('G5:I5').merge().setValue('Total Overtime Logged');
  formatKpiCard(sheet, 'G3:I5', 'G4:I4', '#,##0.0 "hrs"', '#b45309');

  sheet.getRange('J3:N3').merge().setValue('TOTAL DISPATCH LABOR HOURS');
  sheet.getRange('J4:N4').merge();
  sheet.getRange('J5:N5').merge().setValue('Standard + Overtime');
  formatKpiCard(sheet, 'J3:N5', 'J4:N4', '#,##0.0 "hrs"', '#0f172a');

  // Row 2 KPI Cards
  sheet.getRange('A6:D6').merge().setValue('OVERALL FLEET DISPATCH RATIO');
  sheet.getRange('A7:D7').merge();
  sheet.getRange('A8:D8').merge().setValue('Cumulative Dispatch Overhead Ratio');
  formatKpiCard(sheet, 'A6:D8', 'A7:D7', '0.000', '#0369a1');

  sheet.getRange('E6:H6').merge().setValue('OVERALL SUPPORT MINS / ROAD HR');
  sheet.getRange('E7:H7').merge();
  sheet.getRange('E8:H8').merge().setValue('Dispatch Mins per 1 Road Driving Hr');
  formatKpiCard(sheet, 'E6:H8', 'E7:H7', '0.0 "mins"', '#4338ca');

  sheet.getRange('I6:N6').merge().setValue('OVERALL DRIVER LEVERAGE RATIO');
  sheet.getRange('I7:N7').merge();
  sheet.getRange('I8:N8').merge().setValue('Road Hrs driven per 1 Dispatch Hr');
  formatKpiCard(sheet, 'I6:N8', 'I7:N7', '0.00 "x"', '#15803d');

  // Row 3 KPI Cards (20h Concurrency)
  sheet.getRange('A9:D9').merge().setValue('AVG REAL-TIME CONCURRENT CARS');
  sheet.getRange('A10:D10').merge();
  sheet.getRange('A11:D11').merge().setValue('Estimated Active Cars in 20h Window');
  formatKpiCard(sheet, 'A9:D11', 'A10:D10', '0.0 "cars"', '#0d9488');

  sheet.getRange('E9:H9').merge().setValue('LIVE CARS / ON-DUTY DISPATCHER');
  sheet.getRange('E10:H10').merge();
  sheet.getRange('E11:H11').merge().setValue('Real-Time Load (~1.95 Concurrent Disp)');
  formatKpiCard(sheet, 'E9:H11', 'E10:H10', '0.00 "cars/disp"', '#0284c7');

  sheet.getRange('I9:N9').merge().setValue('DESK CAPACITY UTILIZATION');
  sheet.getRange('I10:N10').merge();
  sheet.getRange('I11:N11').merge().setValue('Vs 1:4 Optimal Benchmark (4 cars/disp)');
  formatKpiCard(sheet, 'I9:N11', 'I10:N10', '0.0%', '#7c3aed');

  sheet.setRowHeight(3, 20);
  sheet.setRowHeight(4, 30);
  sheet.setRowHeight(5, 18);
  sheet.setRowHeight(6, 20);
  sheet.setRowHeight(7, 30);
  sheet.setRowHeight(8, 18);
  sheet.setRowHeight(9, 20);
  sheet.setRowHeight(10, 30);
  sheet.setRowHeight(11, 18);

  sheet.getRange('A13:N13').merge()
    .setBackground('#0f172a')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('left');
  sheet.setRowHeight(13, 26);

  sheet.getRange('A14:N14')
    .setBackground('#334155')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(9)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
  sheet.setRowHeight(14, 28);

  if (cfg.monthCount > 0) {
    const monthDataRange = sheet.getRange(cfg.startMonthRow, 1, cfg.monthCount, 14);
    monthDataRange.setFontSize(9).setVerticalAlignment('middle');

    sheet.getRange(cfg.startMonthRow, 1, cfg.monthCount, 1).setHorizontalAlignment('left');
    sheet.getRange(cfg.startMonthRow, 2, cfg.monthCount, 13).setHorizontalAlignment('right');

    sheet.getRange(cfg.startMonthRow, 2, cfg.monthCount, 1).setNumberFormat('#,##0');
    sheet.getRange(cfg.startMonthRow, 3, cfg.monthCount, 4).setNumberFormat('#,##0.0');
    sheet.getRange(cfg.startMonthRow, 7, cfg.monthCount, 1).setNumberFormat('0.000');
    sheet.getRange(cfg.startMonthRow, 8, cfg.monthCount, 1).setNumberFormat('0.0 "mins"');
    sheet.getRange(cfg.startMonthRow, 9, cfg.monthCount, 1).setNumberFormat('0.00 "x"');
    sheet.getRange(cfg.startMonthRow, 10, cfg.monthCount, 1).setNumberFormat('#,##0');
    sheet.getRange(cfg.startMonthRow, 11, cfg.monthCount, 1).setNumberFormat('0.0');
    sheet.getRange(cfg.startMonthRow, 12, cfg.monthCount, 1).setNumberFormat('0.00');
    sheet.getRange(cfg.startMonthRow, 13, cfg.monthCount, 1).setNumberFormat('0.0%');
    sheet.getRange(cfg.startMonthRow, 14, cfg.monthCount, 1).setNumberFormat('0.00');

    for (let r = cfg.startMonthRow; r <= cfg.endMonthRow; r++) {
      sheet.getRange(r, 1, 1, 14).setBackground(r % 2 === 0 ? theme.zebraBg : '#ffffff');
      sheet.setRowHeight(r, 22);
    }

    monthDataRange.setBorder(true, true, true, true, true, true, '#e2e8f0', SpreadsheetApp.BorderStyle.SOLID);

    const totMRange = sheet.getRange(cfg.totalMonthRow, 1, 1, 14);
    totMRange.setBackground('#e2e8f0').setFontWeight('bold').setFontSize(9).setVerticalAlignment('middle');
    sheet.getRange(cfg.totalMonthRow, 1).setHorizontalAlignment('left');
    sheet.getRange(cfg.totalMonthRow, 2, 1, 13).setHorizontalAlignment('right');

    sheet.getRange(cfg.totalMonthRow, 2).setNumberFormat('#,##0.0');
    sheet.getRange(cfg.totalMonthRow, 3, 1, 4).setNumberFormat('#,##0.0');
    sheet.getRange(cfg.totalMonthRow, 7).setNumberFormat('0.000');
    sheet.getRange(cfg.totalMonthRow, 8).setNumberFormat('0.0 "mins"');
    sheet.getRange(cfg.totalMonthRow, 9).setNumberFormat('0.00 "x"');
    sheet.getRange(cfg.totalMonthRow, 10).setNumberFormat('#,##0');
    sheet.getRange(cfg.totalMonthRow, 11).setNumberFormat('0.0');
    sheet.getRange(cfg.totalMonthRow, 12).setNumberFormat('0.00');
    sheet.getRange(cfg.totalMonthRow, 13).setNumberFormat('0.0%');
    sheet.getRange(cfg.totalMonthRow, 14).setNumberFormat('0.00');

    totMRange.setBorder(true, true, true, true, false, false, '#475569', SpreadsheetApp.BorderStyle.SOLID);
    sheet.setRowHeight(cfg.totalMonthRow, 26);
  }

  sheet.getRange(cfg.sec2TitleRow, 1, 1, 14).merge()
    .setBackground('#0f172a')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('left');
  sheet.setRowHeight(cfg.sec2TitleRow, 26);

  sheet.getRange(cfg.sec2HeaderRow, 1, 1, 14)
    .setBackground('#334155')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(9)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
  sheet.setRowHeight(cfg.sec2HeaderRow, 28);

  if (cfg.driverCount > 0) {
    const driverDataRange = sheet.getRange(cfg.startDriverRow, 1, cfg.driverCount, 14);
    driverDataRange.setFontSize(9).setVerticalAlignment('middle');

    sheet.getRange(cfg.startDriverRow, 1, cfg.driverCount, 1).setHorizontalAlignment('left');
    sheet.getRange(cfg.startDriverRow, 2, cfg.driverCount, 9).setHorizontalAlignment('right');
    sheet.getRange(cfg.startDriverRow, 11, cfg.driverCount, 2).setHorizontalAlignment('center');

    sheet.getRange(cfg.startDriverRow, 2, cfg.driverCount, 1).setNumberFormat('#,##0');
    sheet.getRange(cfg.startDriverRow, 3, cfg.driverCount, 1).setNumberFormat('#,##0.0');
    sheet.getRange(cfg.startDriverRow, 4, cfg.driverCount, 1).setNumberFormat('0.0%');
    sheet.getRange(cfg.startDriverRow, 5, cfg.driverCount, 1).setNumberFormat('#,##0.0');
    sheet.getRange(cfg.startDriverRow, 6, cfg.driverCount, 1).setNumberFormat('#,##0');
    sheet.getRange(cfg.startDriverRow, 7, cfg.driverCount, 4).setNumberFormat('0.00');

    for (let dr = cfg.startDriverRow; dr <= cfg.endDriverRow; dr++) {
      sheet.getRange(dr, 1, 1, 14).setBackground(dr % 2 === 0 ? theme.zebraBg : '#ffffff');
      sheet.setRowHeight(dr, 22);
    }

    driverDataRange.setBorder(true, true, true, true, true, true, '#e2e8f0', SpreadsheetApp.BorderStyle.SOLID);

    const totDRange = sheet.getRange(cfg.totalDriverRow, 1, 1, 14);
    totDRange.setBackground('#e2e8f0').setFontWeight('bold').setFontSize(9).setVerticalAlignment('middle');
    sheet.getRange(cfg.totalDriverRow, 1).setHorizontalAlignment('left');
    sheet.getRange(cfg.totalDriverRow, 2, 1, 9).setHorizontalAlignment('right');

    sheet.getRange(cfg.totalDriverRow, 2).setNumberFormat('#,##0');
    sheet.getRange(cfg.totalDriverRow, 3).setNumberFormat('#,##0.0');
    sheet.getRange(cfg.totalDriverRow, 4).setNumberFormat('0.0%');
    sheet.getRange(cfg.totalDriverRow, 5).setNumberFormat('#,##0.0');
    sheet.getRange(cfg.totalDriverRow, 6).setNumberFormat('#,##0');
    sheet.getRange(cfg.totalDriverRow, 7, 1, 4).setNumberFormat('0.00');

    totDRange.setBorder(true, true, true, true, false, false, '#475569', SpreadsheetApp.BorderStyle.SOLID);
    sheet.setRowHeight(cfg.totalDriverRow, 26);
  }

  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 90);
  sheet.setColumnWidth(3, 110);
  sheet.setColumnWidth(4, 110);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 110);
  sheet.setColumnWidth(7, 110);
  sheet.setColumnWidth(8, 110);
  sheet.setColumnWidth(9, 110);
  sheet.setColumnWidth(10, 110);
  sheet.setColumnWidth(11, 100);
  sheet.setColumnWidth(12, 100);
  sheet.setColumnWidth(13, 100);
  sheet.setColumnWidth(14, 100);

  sheet.setFrozenRows(14);
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    buildMasterSummarySheet,
    aggregateCumulativeDriverData,
    collectDataFromExistingAnalysisSheets
  };
}
