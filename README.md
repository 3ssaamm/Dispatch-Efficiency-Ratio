# 🚗 Google Sheets Fleet Dispatch Efficiency Engine

An automated analytics and labor efficiency system powered by Google Apps Script and Google Sheets. It connects to your designated Google Drive balance folder (including `2025` and `2026` subfolders), extracts driver operational data from the `Summary` tab of each monthly balance workbook, and computes calendar-accurate dispatch labor ratios and per-driver support metrics.

---

## 🎯 Key Features & Capabilities

1. **Direct Google Drive Integration**:
   - Built-in root folder: [`Google Drive Fleet Balance Folder`](https://drive.google.com/drive/folders/1iDd2ME2b6coX8GO6Dkom5u9jldkRSwVg?usp=sharing) (`1iDd2ME2b6coX8GO6Dkom5u9jldkRSwVg`).
   - Automatically traverses the root folder and year subfolders (e.g., `2025`, `2026`).
   - Finds all monthly spreadsheets matching `[Month] - Drivers Daily Balance` (e.g. `August - Drivers Daily Balance`, `August 2026 - Drivers Daily Balance`).

2. **Calendar-Accurate Dispatch Labor Logic**:
   - Evaluates the **exact calendar days** of each month (including leap years).
   - 4-Dispatcher schedule:
     - **Dispatcher 1:** 10 hrs/day (Mon–Fri | Sat & Sun OFF)
     - **Dispatcher 2:** 10 hrs/day (Mon–Fri | Sat & Sun OFF)
     - **Dispatcher 3:** 10 hrs/day (Mon–Fri | Sat & Sun OFF)
     - **Dispatcher 4:** 9 hrs/day (Mon–Fri | Sat & Sun OFF)
     - **Daily Fleet Labor:** Monday–Friday = **39.0 hrs/day** | Saturday & Sunday = **0.0 hrs/day**.

3. **Dynamic Source Tab Header Detection**:
   - Dynamically searches column positions in the `Summary` tab for `Driver`, `Hours`, `Trips`, `Credit`, `Cash`, `Balance`, and `No Show`.
   - Immune to shifted column orders, extra title rows, and renamed aliases.
   - Gracefully handles missing hours or zero values by logging a warning and defaulting hours to `0` without breaking the sync.

4. **Live Dynamic Formula Linkage**:
   - The generated `[Month] - Analysis` tab uses live formulas.
   - **Interactive Overtime Adjustment:** Editing cell **`E4` (Overtime / Adj Hours)** immediately updates the entire sheet in real-time (Total Dispatch Hours, Fleet Dispatch Ratio, Support Minutes/Road Hour, Driver Leverage Ratio, and every individual driver's Allocated Dispatch Hours and Trips/Dispatch Hour).

5. **Custom UI Menu & Execution Logging**:
   - Custom menu **`🚗 Fleet Dispatch Tools`** directly inside Google Sheets.
   - Timestamped auditing in the `Execution Log` sheet.

---

## 📊 Core Metric Definitions & Formulas

| Metric | Formula | Description |
| :--- | :--- | :--- |
| **Total Fleet Driver Hours** | `=SUM(Driver Hours)` | Total active road hours logged by all drivers in the month. |
| **Standard Dispatch Hours** | Calendar Logic | Exact monthly base labor hours (Sum of all weekday 39.0-hr shifts). |
| **Overtime / Adjustment** | User Input (Cell `E4`) | Additional or adjusted dispatch hours (Default: `0`). |
| **Total Dispatch Hours** | `= Standard + Overtime` | Full dispatch labor deployed for the month. |
| **Fleet Dispatch Ratio** | `= Total Dispatch Hrs / Total Driver Hrs` | Overhead ratio representing dispatch hours required per road hour. |
| **Support Mins / Road Hr** | `= Fleet Dispatch Ratio * 60` | Number of dispatch minutes spent supporting 1 hour of driving. |
| **Driver Leverage Ratio** | `= Total Driver Hrs / Total Dispatch Hrs` | Number of active road hours generated per 1 dispatch hour. |
| **% Share of Fleet Hours** | `= Driver Hours / Total Fleet Hours` | Driver's individual contribution to total fleet driving time. |
| **Allocated Dispatch Hours** | `= Driver Hours * Fleet Dispatch Ratio` | Fair-share dispatch labor hours attributed to that specific driver. |
| **Trips / Driver Hour** | `= Completed Trips / Driver Hours` | Driver's trip completion pace per active driving hour. |
| **Trips / Dispatch Hour** | `= Completed Trips / Allocated Disp Hours` | Trip efficiency yield per unit of dispatch labor allocated. |

---

## 🚀 Quick Setup Instructions (3 Steps)

### Step 1: Open Your Target Google Spreadsheet
1. Create a new Google Spreadsheet (or open an existing one).
2. Name it e.g. `Fleet Dispatch Efficiency Dashboard`.

### Step 2: Add the Apps Script Code
1. In the top menu, click **Extensions** $\rightarrow$ **Apps Script**.
2. Delete any boilerplate code in `Code.gs`.
3. Open [`Consolidated_Code.gs`](file:///c:/Users/midoe/Documents/Business/Freelancing/AYOAH%20Go/Balance/Ayoah-Scripts/Dispatch-Efficiency-Ratio/Consolidated_Code.gs) from this repository, copy its entire contents, and paste it into the editor.
   *(Alternatively, you can keep the modular files: `Config.js`, `DispatcherSchedule.js`, `DriveSync.js`, `AnalysisBuilder.js`, `SettingsManager.js`, `Menu.js`, `Utils.js`).*
4. Click the **💾 Save project** icon (or press `Ctrl+S`).

### Step 3: Refresh and Run
1. Switch back to your Google Sheet tab and **refresh the browser** (`F5` or `Ctrl+R`).
2. A new menu **`🚗 Fleet Dispatch Tools`** will appear in the top toolbar!
3. Click **`🚗 Fleet Dispatch Tools`** $\rightarrow$ **`🔄 Sync All Monthly Files from Drive`** (or **`📅 Sync Selected Month`**).
4. Google will prompt you to authorize permissions on the first run (Click *Continue* $\rightarrow$ *Advanced* $\rightarrow$ *Go to Fleet Dispatch Tools* $\rightarrow$ *Allow*).
5. The script will automatically scan the Google Drive folder, process all monthly files (including `2025` and `2026` subfolders), and generate a dedicated `[Month] - Analysis` tab for each month!

---

## 🖥️ Custom Menu Options

```
🚗 Fleet Dispatch Tools
├── 🔄 Sync All Monthly Files from Drive
├── 📅 Sync Selected Month
├──────────────────────────────────────
├── ⚡ Recalculate All Dispatch Ratios
├──────────────────────────────────────
├── ⚙️ Open / Reset Settings Tab
├── 🧪 Generate Sample Data in Drive (Demo)
└── 📖 User Guide & Metric Formulas
```

- **🔄 Sync All Monthly Files from Drive:** Scans the folder and subfolders, processes all monthly sheets, and creates/refreshes the analysis tabs.
- **📅 Sync Selected Month:** Prompts for a month (e.g. `August 2026` or `August`) and syncs just that workbook.
- **⚡ Recalculate All Dispatch Ratios:** Updates calendar dispatch baselines across all existing tabs.
- **⚙️ Open / Reset Settings Tab:** Initializes an interactive configuration sheet to review shift schedules and folder settings.
- **🧪 Generate Sample Data in Drive (Demo):** Creates a sample `August 2026 - Drivers Daily Balance` sheet directly in your Drive folder with realistic driver data to test immediately.

---

## 📁 Repository Structure

```
Dispatch-Efficiency-Ratio/
├── Consolidated_Code.gs    # Single-file copy-paste bundle for Google Apps Script
├── Config.js               # Global configuration, shifts, aliases, theme
├── DispatcherSchedule.js   # Calendar engine & day-of-week labor math
├── DriveSync.js            # Drive scanner & Summary tab dynamic parser
├── AnalysisBuilder.js      # Sheet generator, KPI card builder, table styling
├── SettingsManager.js      # Settings sheet initializer & reader
├── Menu.js                 # UI Menu handlers and orchestrators
├── Utils.js                # Execution logger and UI helpers
├── MockDataGenerator.js    # Demo spreadsheet generator
├── appsscript.json         # Apps Script project manifest
├── test/
│   └── engine.test.js      # Automated unit test suite (Node.js)
└── README.md               # User & Technical Documentation
```

---

## 🧪 Automated Testing & Verification

Run the test suite locally:
```bash
node test/engine.test.js
```
**Test Results:**
- ✅ **Test Suite 1:** Month & Year Parsing across varied strings (`August`, `August 2026`, `2025 - January`, etc.)
- ✅ **Test Suite 2:** Calendar labor schedule math (August 2026: 21 weekdays $\times 39.0\text{h} = 819.0\text{ hrs}$; Leap years)
- ✅ **Test Suite 3:** Dynamic header detection with shifted and transposed columns
- ✅ **Test Suite 4:** Numeric parsing with currency signs, commas, and edge cases
- ✅ **Test Suite 5:** Summary/Total row detection and filtering
- ✅ **Test Suite 6:** Live ratio math & metric formulas
