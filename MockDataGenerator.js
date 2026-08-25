/**
 * @fileoverview Mock Data Generator for Fleet Dispatch Efficiency Engine.
 * Creates realistic sample monthly balance spreadsheets in Google Drive for testing and demonstrations.
 */

/**
 * Creates a sample monthly balance spreadsheet in the target Drive folder.
 * 
 * @param {string} [monthName='August']
 * @param {number} [year=2026]
 * @returns {GoogleAppsScript.Drive.File} Created Google Drive File
 */
function createSampleMonthlySpreadsheet(monthName, year) {
  const mName = monthName || 'August';
  const y = year || 2026;
  const fileName = `${mName} ${y} - Drivers Daily Balance`;

  const folderId = getEffectiveDriveFolderId();
  let folder;
  try {
    folder = DriveApp.getFolderById(folderId);
  } catch (e) {
    throw new Error(`Cannot access Drive folder (${folderId}) to create sample file: ${e.message}`);
  }

  // Create new spreadsheet
  const newSs = SpreadsheetApp.create(fileName);
  const targetSheetName = (typeof CONFIG !== 'undefined' && CONFIG.SOURCE_SHEET_NAME) ? CONFIG.SOURCE_SHEET_NAME : 'Summary';

  let sheet = newSs.getActiveSheet();
  sheet.setName(targetSheetName);

  // Sample driver records
  const sampleDrivers = [
    { name: 'Marcus Vance (101)', hours: 168.5, trips: 224, credit: 3200, cash: 450, balance: 3650, noShow: 3 },
    { name: 'Elena Rostova (102)', hours: 155.0, trips: 210, credit: 2950, cash: 620, balance: 3570, noShow: 1 },
    { name: 'David Kim (103)', hours: 172.0, trips: 245, credit: 3410, cash: 510, balance: 3920, noShow: 2 },
    { name: 'Fatima Al-Mansoor (104)', hours: 144.5, trips: 198, credit: 2800, cash: 390, balance: 3190, noShow: 4 },
    { name: 'Carlos Mendez (105)', hours: 160.0, trips: 215, credit: 3100, cash: 480, balance: 3580, noShow: 2 },
    { name: 'Aisha Diallo (106)', hours: 138.0, trips: 180, credit: 2600, cash: 310, balance: 2910, noShow: 1 },
    { name: 'James O\'Connor (107)', hours: 180.5, trips: 260, credit: 3750, cash: 700, balance: 4450, noShow: 5 },
    { name: 'Priya Sharma (108)', hours: 152.0, trips: 205, credit: 2900, cash: 420, balance: 3320, noShow: 0 },
    { name: 'Lucas Silva (109)', hours: 140.0, trips: 190, credit: 2750, cash: 360, balance: 3110, noShow: 2 },
    { name: 'Tariq Hassan (110)', hours: 165.0, trips: 230, credit: 3300, cash: 540, balance: 3840, noShow: 3 },
    { name: 'Nathalie Dupont (111)', hours: 125.5, trips: 165, credit: 2400, cash: 290, balance: 2690, noShow: 1 },
    { name: 'Viktor Ivanov (112)', hours: 158.0, trips: 218, credit: 3150, cash: 490, balance: 3640, noShow: 2 },
    { name: 'Grace Hopper (113)', hours: 148.0, trips: 200, credit: 2850, cash: 410, balance: 3260, noShow: 1 },
    { name: 'Kenji Sato (114)', hours: 135.5, trips: 175, credit: 2500, cash: 350, balance: 2850, noShow: 0 },
    { name: 'Alan Turing (115)', hours: 0.0, trips: 0, credit: 0, cash: 0, balance: 0, noShow: 0 } // Zero hours test case
  ];

  // Header Row
  const headers = ['Driver Name', 'Active Hours', 'Completed Trips', 'Credit Card ($)', 'Cash ($)', 'Net Balance ($)', 'No Shows'];
  
  const rows = [headers];
  let sumHours = 0;
  let sumTrips = 0;
  let sumCredit = 0;
  let sumCash = 0;
  let sumBal = 0;
  let sumNoShow = 0;

  for (const d of sampleDrivers) {
    rows.push([d.name, d.hours, d.trips, d.credit, d.cash, d.balance, d.noShow]);
    sumHours += d.hours;
    sumTrips += d.trips;
    sumCredit += d.credit;
    sumCash += d.cash;
    sumBal += d.balance;
    sumNoShow += d.noShow;
  }

  // Add Summary Total Row (to test that our parser safely ignores totals)
  rows.push(['Total Fleet Sum', sumHours, sumTrips, sumCredit, sumCash, sumBal, sumNoShow]);

  sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);

  // Format header
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#1e293b')
    .setFontColor('#ffffff')
    .setFontWeight('bold');

  // Format numbers
  sheet.getRange(2, 2, rows.length - 1, 1).setNumberFormat('#,##0.0');
  sheet.getRange(2, 3, rows.length - 1, 1).setNumberFormat('#,##0');
  sheet.getRange(2, 4, rows.length - 1, 3).setNumberFormat('$#,##0.00');

  // Move file into the target Google Drive folder
  const driveFile = DriveApp.getFileById(newSs.getId());
  driveFile.moveTo(folder);

  return driveFile;
}

// Node.js module export for testing
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    createSampleMonthlySpreadsheet
  };
}
