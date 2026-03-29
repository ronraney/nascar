/**
 * ============================================================
 *  NASCAR DFS 2026 — Module 4: DASHBOARD
 * ============================================================
 *  Active groups: DOM, PD, LEVERAGE, UNDER.
 *  CORE dissolved — falls into Fill/ungrouped.
 * ============================================================
 */


/* -------------------------------------------------------
 *  1. Menu & Entry Points
 * ------------------------------------------------------- */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('NASCAR DFS')
    .addItem('▶  Refresh Dashboard',     'runGPPAnalysis')
    .addItem('📊  Generate Pools',       'renderPools')
    .addSeparator()
    .addItem('🏁  Race Control',         'showSidebar')
    .addItem('🧰  Lineup Builder',       'showLineupBuilder')
    .addItem('⚡  Optimizer',            'openOptimizer')
    .addSeparator()
    .addItem('🔍  Data Diagnostics',    'runDiagnostics')
    .addItem('Setup Sheets',             'setupAllSheets')
    .addToUi();
}


/* -------------------------------------------------------
 *  2. Master Analysis Run
 * ------------------------------------------------------- */

function runGPPAnalysis() {
  const t0 = new Date();

  const data    = loadAllData();
  const drivers = runAnalysis(data);
  const rc      = data.raceContext;

  renderGPPTable(drivers, rc);

  const domT = rc.targetDoms;
  const pdT  = rc.targetPD;
  const domStr = (domT && domT.min !== undefined)
    ? domT.min + "-" + domT.max
    : "" + (domT || 0);
  const pdStr = (pdT && pdT.min !== undefined)
    ? pdT.min + "-" + pdT.max
    : "" + (pdT || 0);

  const elapsed = ((new Date() - t0) / 1000).toFixed(1);
  SpreadsheetApp.getUi().alert(
    "Dashboard Refreshed (" + elapsed + "s)\n" +
    drivers.length + " drivers | " + rc.trackType +
    " | Dom: " + domStr + " | PD: " + pdStr
  );
}


/* -------------------------------------------------------
 *  3. Master Cash Run
 * ------------------------------------------------------- */

function runCashAnalysis() {
  const data    = loadAllData();
  const drivers = runAnalysis(data);
  const lineup  = buildCashLineup(drivers);

  renderCashLineup(lineup, data.raceContext);

  SpreadsheetApp.getUi().alert(
    "Cash Lineup Built — " + lineup.length + " drivers, $" +
    lineup.reduce((s, d) => s + d.salary, 0).toLocaleString() + " salary"
  );
}


/* -------------------------------------------------------
 *  4. GPP Driver Table
 * ------------------------------------------------------- */

function renderGPPTable(drivers, rc) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName("Dashboard") || ss.insertSheet("Dashboard");
  const HDR  = DASH_COLS.GPP_HEADERS;
  const hRow = DASH_COLS.GPP_HEADER_ROW;
  const dRow = DASH_COLS.GPP_DATA_START;
  const TC   = DASH_COLS.TOTAL_COLS;

  dash.clear();
  dash.getRange(1, 1, dash.getMaxRows(), TC).clearFormat().clearDataValidations();

  dash.getRange(hRow, 1, 1, TC).setValues([HDR])
    .setBackground("#1a1a2e").setFontColor("white").setFontWeight("bold")
    .setFontSize(9).setHorizontalAlignment("center");

  dash.setFrozenRows(hRow);
  drivers.sort((a, b) => b.adjProj - a.adjProj);

  const targetDomsMax = (rc.targetDoms && rc.targetDoms.max !== undefined)
    ? rc.targetDoms.max : (rc.targetDoms || 0);

  const rows = drivers.map(d => [
    false,
    d.name,
    d.salary,
    d.startPos,
    d.ownPct / 100,
    d.group,
    round2(d.proj),
    round2(d.adjProj),
    round2(d.floor),
    round2(d.ceiling),
    round2(d.dkStd),
    round2(d.domPts),
    d.domRank,
    round2(d.pdProj),
    round2(d.edge),
    round2(d.value),
    round2(d.cashScore),
    round2(d.trackHistScore),
    round2(d.histAvgStartFinishDiff),
    d.notes.join(" | ")
  ]);

  if (rows.length === 0) return;

  const dataRange = dash.getRange(dRow, 1, rows.length, TC);
  dataRange.setValues(rows);
  dataRange.setHorizontalAlignment("center");
  dataRange.setFontSize(9);

  dash.getRange(dRow, DASH_COLS.COL_CHECK,  rows.length).insertCheckboxes();
  dash.getRange(dRow, DASH_COLS.COL_START,  rows.length).setNumberFormat("0");
  dash.getRange(dRow, DASH_COLS.COL_OWN,    rows.length).setNumberFormat("0.0%");
  dash.getRange(dRow, DASH_COLS.COL_SALARY, rows.length).setNumberFormat("$#,##0");

  // Active group colors — CORE removed
  const GROUP_COLORS = {
    "DOM":      { bg: "#ff6b6b", fg: "white" },
    "PD":       { bg: "#4ecdc4", fg: "white" },
    "LEVERAGE": { bg: "#ffe66d", fg: "#333"  },
    "UNDER":    { bg: "#999999", fg: "white" },
    "CASHCORE": { bg: "#6c63ff", fg: "white" }
  };

  for (let i = 0; i < rows.length; i++) {
    const group  = rows[i][5];
    const colors = GROUP_COLORS[group];
    if (colors) {
      dash.getRange(dRow + i, DASH_COLS.COL_GROUP)
        .setBackground(colors.bg).setFontColor(colors.fg).setFontWeight("bold");
    }
  }

  for (let i = 0; i < rows.length; i++) {
    if (rows[i][12] <= targetDomsMax && targetDomsMax > 0) {
      dash.getRange(dRow + i, DASH_COLS.COL_DOMRANK)
        .setBackground("#ffe0e0").setFontWeight("bold");
    }
  }

  for (let i = 0; i < rows.length; i++) {
    const diff = rows[i][DASH_COLS.COL_AVGDIFF - 1];
    let bg = null;
    if (diff >= 8)       bg = "#E8F5E9";
    else if (diff >= 4)  bg = "#F1F8E9";
    else if (diff <= -4) bg = "#FFEBEE";
    if (bg) dash.getRange(dRow + i, DASH_COLS.COL_AVGDIFF).setBackground(bg);
  }

  dash.setColumnWidth(1, 30);
  for (let c = 2; c <= TC; c++) dash.autoResizeColumn(c);
}


/* -------------------------------------------------------
 *  5. Cash Lineup Renderer
 * ------------------------------------------------------- */

function renderCashLineup(lineup, rc) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Cash_Lineup");
  if (!sheet) sheet = ss.insertSheet("Cash_Lineup");
  sheet.clear();

  sheet.getRange(1, 1, 1, 7).merge()
    .setValue("CASH LINEUP — " + rc.raceName + " (" + rc.trackType + ")")
    .setBackground("#1a1a2e").setFontColor("white").setFontWeight("bold")
    .setFontSize(12).setHorizontalAlignment("center");

  const headers = DASH_COLS.CASH_HEADERS;
  sheet.getRange(3, 1, 1, headers.length).setValues([headers])
    .setBackground("#1a73e8").setFontColor("white").setFontWeight("bold")
    .setHorizontalAlignment("center");

  let totalSal = 0, totalFloor = 0, totalProj = 0, totalCash = 0;
  const rows = lineup.map((d, i) => {
    totalSal   += d.salary;
    totalFloor += d.floor;
    totalProj  += d.adjProj;
    totalCash  += d.cashScore;
    return [i + 1, d.name, d.salary, round2(d.floor), round2(d.adjProj), d.ownPct / 100, round2(d.cashScore)];
  });

  if (rows.length > 0) {
    sheet.getRange(4, 1, rows.length, headers.length).setValues(rows).setHorizontalAlignment("center");
    const totRow = 4 + rows.length;
    sheet.getRange(totRow, 1, 1, headers.length)
      .setValues([["TOTAL", "", totalSal, round2(totalFloor), round2(totalProj), "", round2(totalCash)]])
      .setFontWeight("bold").setBackground("#f0f0f0").setHorizontalAlignment("center");
    sheet.getRange(totRow + 1, 1, 1, 3)
      .setValues([["Salary Remaining", "", CASH_SALARY_CAP - totalSal]]).setFontWeight("bold");
    sheet.getRange(4, 3, rows.length + 1).setNumberFormat("$#,##0");
    sheet.getRange(4, 6, rows.length).setNumberFormat("0.0%");
  }

  for (let c = 1; c <= headers.length; c++) sheet.autoResizeColumn(c);
  sheet.setFrozenRows(3);
}


/* -------------------------------------------------------
 *  6. Setup All Sheets
 * ------------------------------------------------------- */

function setupAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ["Dashboard", "Cash_Lineup", "Lineups", "Pools", "Model_Config"].forEach(name => {
    if (!ss.getSheetByName(name)) ss.insertSheet(name);
  });

  const lineups = ss.getSheetByName("Lineups");
  if (lineups.getRange(1, 1).getValue() !== "D1") {
    lineups.getRange(1, 1, 1, 8).setValues([[
      "D1", "D2", "D3", "D4", "D5", "D6", "Total Sal", "Saved At"
    ]]).setBackground("#1a1a2e").setFontColor("white").setFontWeight("bold");
    lineups.setFrozenRows(1);
  }

  SpreadsheetApp.getUi().alert(
    "Sheets ready.\n\n" +
    "1. Paste DraftKings CSV into DK_Salaries\n" +
    "2. Paste iFantasyRace data into Playability\n" +
    "3. NASCAR DFS › ▶ Refresh Dashboard\n" +
    "4. NASCAR DFS › 🧰 Lineup Builder"
  );
}


/* -------------------------------------------------------
 *  7. Sidebar Launchers
 * ------------------------------------------------------- */

function showSidebar() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const config  = ss.getSheetByName("Model_Config") || ss.insertSheet("Model_Config");
  const raceEnv = ss.getSheetByName("Race_Environment");
  const html    = HtmlService.createTemplateFromFile('Sidebar');

  if (raceEnv) {
    const raceData = raceEnv.getRange(2, 3, raceEnv.getLastRow() - 1, 1).getValues();
    html.raceList = raceData.flat();
  } else {
    html.raceList = ["No Race Data"];
  }

  html.currentRace = config.getRange("B1").getValue() || "Select Race";
  html.type        = config.getRange("B2").getValue() || "N/A";
  html.laps        = config.getRange("B3").getValue() || "0";
  html.strat       = config.getRange("B4").getValue() || "N/A";

  const trackType = config.getRange("B2").getValue() || "";
  const w = getWeights(trackType);
  html.weightStr = trackType
    ? "dom:" + (w.dom * 100) + "% pd:" + (w.pd * 100) + "% spd:" + (w.speed * 100)
      + "% skl:" + (w.skill * 100) + "% hst:" + ((w.history || 0) * 100) + "%"
    : "";

  const domInfo = calcDomPointsAvailable(config.getRange("B3").getValue() || 0, trackType);
  html.domAvail = domInfo.adjusted || 0;

  const domT = getTargetDominators(trackType);
  const pdT  = getTargetPD(trackType);
  html.targetDoms = domT.min === domT.max ? "" + domT.max : domT.min + "-" + domT.max;
  html.targetPD   = pdT.min === pdT.max   ? "" + pdT.max  : pdT.min  + "-" + pdT.max;

  SpreadsheetApp.getUi().showSidebar(
    html.evaluate().setTitle('Race Control').setWidth(300)
  );
}

function updateRaceSettings(raceName) {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const data   = ss.getSheetByName("Race_Environment").getDataRange().getValues();
  const config = ss.getSheetByName("Model_Config");

  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === raceName) {
      const type  = data[i][4];
      const laps  = data[i][5];
      const strat = data[i][6];

      config.getRange("A1:B4").setValues([
        ["Race", data[i][2]], ["Type", type], ["Laps", laps], ["Strategy", strat]
      ]);

      const w       = getWeights(type);
      const wStr    = "dom:" + (w.dom * 100) + "% pd:" + (w.pd * 100)
                    + "% spd:" + (w.speed * 100) + "% skl:" + (w.skill * 100)
                    + "% hst:" + ((w.history || 0) * 100) + "%";
      const domInfo = calcDomPointsAvailable(laps, type);
      const domT    = getTargetDominators(type);
      const pdT     = getTargetPD(type);

      return {
        type:       type,
        laps:       laps,
        strat:      strat,
        weights:    wStr,
        domAvail:   domInfo.adjusted,
        targetDoms: domT.min === domT.max ? "" + domT.max : domT.min + "-" + domT.max,
        targetPD:   pdT.min === pdT.max   ? "" + pdT.max  : pdT.min  + "-" + pdT.max
      };
    }
  }
  return null;
}

function getGroupSummary() {
  const drivers = readDashboardDrivers();
  // Active groups — CORE removed
  const groups  = { DOM: [], PD: [], CASHCORE: [], LEVERAGE: [], UNDER: [] };
  let ungrouped = 0;

  for (const d of drivers) {
    if (d.group && groups[d.group] !== undefined) {
      groups[d.group].push(d.name);
    } else {
      ungrouped++;
    }
  }

  return { groups, ungrouped, total: drivers.length };
}


/* -------------------------------------------------------
 *  8. Helper
 * ------------------------------------------------------- */

function round2(v) {
  return Math.round((v || 0) * 100) / 100;
}


/* -------------------------------------------------------
 *  9. Lineup Builder — Sidebar Launcher
 * ------------------------------------------------------- */

function showLineupBuilder() {
  const html      = HtmlService.createTemplateFromFile('LineupBuilder');
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const config    = ss.getSheetByName("Model_Config");
  const trackType = config ? config.getRange("B2").getValue() || "Intermediate" : "Intermediate";

  html.trackType = trackType;
  html.salaryCap = CASH_SALARY_CAP;

  SpreadsheetApp.getUi().showSidebar(
    html.evaluate().setTitle('Lineup Builder').setWidth(350)
  );
}


/* -------------------------------------------------------
 *  10. Lineup Builder — Server-Side API
 * ------------------------------------------------------- */

function readDashboardDrivers() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Dashboard");
  if (!sheet) return [];

  const dRow = DASH_COLS.GPP_DATA_START;
  const TC   = DASH_COLS.TOTAL_COLS;
  const last = sheet.getLastRow();
  if (last < dRow) return [];

  const data = sheet.getRange(dRow, 1, last - dRow + 1, TC).getValues();

  return data
    .map((row, i) => ({
      rowIdx:   i + dRow,
      checked:  row[DASH_COLS.COL_CHECK    - 1] === true,
      name:     String(row[DASH_COLS.COL_DRIVER  - 1] || ""),
      salary:   parseFloat(row[DASH_COLS.COL_SALARY - 1]) || 0,
      startPos: parseFloat(row[DASH_COLS.COL_START  - 1]) || 0,
      ownPct:   parseFloat(row[DASH_COLS.COL_OWN    - 1]) || 0,
      group:    String(row[DASH_COLS.COL_GROUP  - 1] || ""),
      proj:     parseFloat(row[DASH_COLS.COL_PROJ   - 1]) || 0,
      adjProj:  parseFloat(row[DASH_COLS.COL_ADJPROJ - 1]) || 0,
      floor:    parseFloat(row[DASH_COLS.COL_FLOOR  - 1]) || 0,
      ceiling:  parseFloat(row[DASH_COLS.COL_CEIL   - 1]) || 0,
      dkStd:    parseFloat(row[DASH_COLS.COL_STD    - 1]) || 0,
      domPts:   parseFloat(row[DASH_COLS.COL_DOMPTS - 1]) || 0,
      domRank:  parseFloat(row[DASH_COLS.COL_DOMRANK - 1]) || 0,
      pdProj:   parseFloat(row[DASH_COLS.COL_PD     - 1]) || 0,
      edge:     parseFloat(row[DASH_COLS.COL_EDGE   - 1]) || 0,
      value:    parseFloat(row[DASH_COLS.COL_VALUE  - 1]) || 0,
      cashScore:           parseFloat(row[DASH_COLS.COL_CASHSCORE - 1]) || 0,
      trackHistScore:      parseFloat(row[DASH_COLS.COL_TRACKHIST  - 1]) || 0,
      avgStartFinishDiff:  parseFloat(row[DASH_COLS.COL_AVGDIFF   - 1]) || 0
    }))
    .filter(d => d.name.trim() !== "");
}

function checkDriver(rowIdx, checked) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Dashboard");
  if (!sheet) return { ok: false };
  sheet.getRange(rowIdx, DASH_COLS.COL_CHECK).setValue(checked !== false);
  return { ok: true };
}

function clearAllChecks() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Dashboard");
  if (!sheet) return { ok: false };
  const dRow = DASH_COLS.GPP_DATA_START;
  const last = sheet.getLastRow();
  if (last >= dRow) {
    sheet.getRange(dRow, DASH_COLS.COL_CHECK, last - dRow + 1).setValue(false);
  }
  return { ok: true };
}

function getSalaryInfo() {
  const drivers = readDashboardDrivers();
  const checked = drivers.filter(d => d.checked);
  const total   = checked.reduce((s, d) => s + d.salary, 0);
  return {
    selected:  checked.length,
    total:     total,
    remaining: CASH_SALARY_CAP - total,
    overcap:   total > CASH_SALARY_CAP || checked.length > 6,
    drivers:   checked.map(d => ({ name: d.name, salary: d.salary, rowIdx: d.rowIdx }))
  };
}

function saveLineupByRows(rowIdxs) {
  const drivers = readDashboardDrivers();
  const idxSet  = new Set(rowIdxs || []);
  const lineup  = drivers.filter(d => idxSet.has(d.rowIdx));
  if (lineup.length !== 6) return { ok: false, msg: "Need exactly 6 drivers, got " + lineup.length };
  const totalSal = lineup.reduce((s, d) => s + d.salary, 0);
  if (totalSal > CASH_SALARY_CAP) return { ok: false, msg: "Over salary cap: $" + totalSal.toLocaleString() };
  return writeLineup(lineup);
}

function saveLineup() {
  const drivers = readDashboardDrivers();
  const lineup  = drivers.filter(d => d.checked);
  if (lineup.length !== 6) return { ok: false, msg: "Need exactly 6 checked drivers, got " + lineup.length };
  const totalSal = lineup.reduce((s, d) => s + d.salary, 0);
  if (totalSal > CASH_SALARY_CAP) return { ok: false, msg: "Over salary cap: $" + totalSal.toLocaleString() };
  return writeLineup(lineup);
}

function writeLineup(lineup) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Lineups");
  if (!sheet) sheet = ss.insertSheet("Lineups");

  if (sheet.getRange(1, 1).getValue() !== "D1") {
    sheet.getRange(1, 1, 1, 8).setValues([[
      "D1", "D2", "D3", "D4", "D5", "D6", "Total Sal", "Saved At"
    ]]).setBackground("#1a1a2e").setFontColor("white").setFontWeight("bold");
    sheet.setFrozenRows(1);
  }

  const totalSal = lineup.reduce((s, d) => s + d.salary, 0);
  const names    = lineup.map(d => d.name);
  const now      = new Date().toLocaleString();

  const colA  = sheet.getRange(1, 1, sheet.getMaxRows(), 1).getValues();
  let nextRow = 2;
  for (let i = colA.length - 1; i >= 1; i--) {
    if (String(colA[i][0]).trim() !== "") { nextRow = i + 2; break; }
  }

  sheet.getRange(nextRow, 1, 1, 8).setValues([[
    names[0] || "", names[1] || "", names[2] || "",
    names[3] || "", names[4] || "", names[5] || "",
    totalSal, now
  ]]);

  updateExposureReport(sheet);
  clearAllChecks();

  return { ok: true, msg: "Lineup #" + (nextRow - 1) + " saved — $" + totalSal.toLocaleString() };
}

function updateExposureReport(sheet) {
  if (!sheet) return;
  const colA = sheet.getRange(2, 1, sheet.getMaxRows() - 1, 1).getValues();
  let lineupCount = 0;
  for (let i = 0; i < colA.length; i++) {
    if (String(colA[i][0]).trim() !== "") lineupCount++;
    else break;
  }

  // Always clear first — ensures stale data is removed even when all lineups deleted
  const expCol = 10;
  sheet.getRange(1, expCol, sheet.getMaxRows(), 2).clearContent().clearFormat();
  sheet.getRange(1, expCol, 1, 2).setValues([["Driver", "Exposure"]])
    .setBackground("#1a1a2e").setFontColor("white").setFontWeight("bold");

  if (lineupCount === 0) return;

  const data = sheet.getRange(2, 1, lineupCount, 6).getValues();
  const counts = {};
  for (const row of data) {
    for (const cell of row) {
      const name = String(cell || "").trim();
      if (name) counts[name] = (counts[name] || 0) + 1;
    }
  }

  const sorted = Object.entries(counts)
    .map(([name, count]) => [name, Math.round((count / lineupCount) * 100) + "%"])
    .sort((a, b) => parseInt(b[1]) - parseInt(a[1]));

  if (sorted.length > 0) {
    sheet.getRange(2, expCol, sorted.length, 2).setValues(sorted);
  }
}

function clearLineups() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Lineups");
  if (!sheet) return { ok: false };
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, sheet.getMaxColumns()).clearContent();
  sheet.getRange(1, 10, sheet.getMaxRows(), 2).clearContent().clearFormat();
  sheet.getRange(1, 10, 1, 2).setValues([["Driver", "Exposure"]])
    .setBackground("#1a1a2e").setFontColor("white").setFontWeight("bold");
  return { ok: true, msg: "All lineups cleared." };
}


/* -------------------------------------------------------
 *  11. Lineup Builder — Get Wizard Steps
 * ------------------------------------------------------- */

function getWizardSteps(contestType) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const config    = ss.getSheetByName("Model_Config");
  const trackType = config ? config.getRange("B2").getValue() || "Intermediate" : "Intermediate";
  const targetDoms = getTargetDominators(trackType);
  const targetPD   = getTargetPD(trackType);
  const isCash     = contestType === "Cash";

  var steps = [];

  // Cash Core step — always present for cash lineups
  if (isCash) {
    steps.push({
      id: "CASHCORE", min: 2, max: 2,
      label: "Cash Core",
      hdr: "Pick 2 Cash Core plays",
      why: "Highest floor value with proven track history. These are your anchors regardless of role."
    });
  }

  if (targetDoms.max > 0) {
    // Cash: always pick exactly 1 DOM. GPP: use track-type target range.
    var domMin = isCash ? 1 : targetDoms.min;
    var domMax = isCash ? 1 : targetDoms.max;
    var domLabel = domMin === domMax ? "" + domMax : domMin + "-" + domMax;
    steps.push({
      id: "DOM", min: domMin, max: domMax,
      label: "Dominators",
      hdr: "Pick " + domLabel + " Dominator" + (domMax > 1 ? "s" : ""),
      why: isCash
        ? "One dominator for laps-led and fastest-lap floor. Take the safest option."
        : "Drivers who lead laps and earn fastest-lap bonus. Look for low-owned dominators for GPP leverage."
    });
  }

  if (targetPD.max > 0) {
    // Cash: always pick exactly 1 PD. GPP: use track-type target range.
    var pdMin = isCash ? 1 : targetPD.min;
    var pdMax = isCash ? 1 : targetPD.max;
    var pdLabel = pdMin === pdMax ? "" + pdMax : pdMin + "-" + pdMax;
    steps.push({
      id: "PD", min: pdMin, max: pdMax,
      label: "Place Differential",
      hdr: "Pick " + pdLabel + " PD play" + (pdMax > 1 ? "s" : ""),
      why: isCash
        ? "One proven place-gainer. Track history required — avoid drivers with no historical data."
        : "Back starters with upside. Low ownership + big PD projection = GPP leverage."
    });
  }

  if (isCash) {
    steps.push({
      id: "PUNT", min: 0, max: 1,
      label: "Punt",
      hdr: "Pick 0-1 Punt",
      why: "Cheapest viable driver with a real floor. Salary relief without a zero."
    });
  } else {
    steps.push({
      id: "DART", min: 0, max: 1,
      label: "Dart",
      hdr: "Pick 0-1 Dart",
      why: "Cheap driver with the highest ceiling per dollar. Low ownership, high upside."
    });
  }

  steps.push({
    id: "FILL", min: 0, max: 9,
    label: "Fill",
    hdr: "Fill remaining slot(s)",
    why: isCash
      ? "Best available driver that fits salary. Prioritize floor and projection."
      : "Best available driver that fits salary. Prioritize ceiling and edge."
  });

  return { steps: steps, trackType: trackType, targetDoms: targetDoms, targetPD: targetPD };
}


/* -------------------------------------------------------
 *  12. Lineup Builder — Start Wizard
 * ------------------------------------------------------- */

function startWizard(contestType) {
  clearAllChecks();
  var result = getWizardSteps(contestType);
  return {
    ok: true,
    steps: result.steps,
    trackType: result.trackType,
    targetDoms: result.targetDoms,
    targetPD: result.targetPD,
    msg: result.steps.length + " steps for " + result.trackType
  };
}


/* -------------------------------------------------------
 *  13. Lineup Builder — Step Pool Server
 * ------------------------------------------------------- */

function getStepPool(stepId, contestType, usedNames, salaryRemaining) {
  const drivers = readDashboardDrivers();
  const used    = new Set(usedNames || []);
  const isCash  = contestType === "Cash";

  const slotsUsed  = used.size;
  const slotsAfter = Math.max(0, (6 - slotsUsed) - 1);
  const salFloor   = slotsAfter > 0 ? 5000 : 0;

  function canAfford(d) {
    if (used.has(d.name)) return false;
    if (d.salary > salaryRemaining) return false;
    if (slotsAfter > 0 && (salaryRemaining - d.salary) / slotsAfter < salFloor) return false;
    return true;
  }

  function toRow(d, keyLabel, keyVal) {
    return {
      rowIdx: d.rowIdx, name: d.name, salary: d.salary,
      group: d.group, ownPct: d.ownPct, adjProj: d.adjProj,
      floor: d.floor, ceiling: d.ceiling, edge: d.edge,
      startPos: d.startPos, cashScore: d.cashScore, trackHistScore: d.trackHistScore,
      domPts: d.domPts, domRank: d.domRank, pdProj: d.pdProj,
      keyLabel: keyLabel, keyVal: keyVal
    };
  }

  if (stepId === "CASHCORE") {
    var pool = drivers.filter(d => canAfford(d) && d.group === "CASHCORE");

    // Recompute cashCoreGrade from dashboard values — floor, adjProj, value, trackHistScore
    // Mirrors assignCashCoreGroup in Analysis.gs
    var floorArr   = pool.map(function(d){ return d.floor; });
    var projArr    = pool.map(function(d){ return d.adjProj; });
    var valueArr   = pool.map(function(d){ return d.salary > 0 ? d.adjProj / (d.salary / 1000) : 0; });
    var trackArr   = pool.map(function(d){ return d.trackHistScore; });

    function normPool(val, arr) {
      var mx = Math.max.apply(null, arr), mn = Math.min.apply(null, arr);
      return mx === mn ? 50 : ((val - mn) / (mx - mn)) * 100;
    }

    pool.forEach(function(d) {
      var rawVal  = d.salary > 0 ? d.adjProj / (d.salary / 1000) : 0;
      d._ccGrade  = Math.round(
          normPool(d.floor,           floorArr) * 0.35
        + normPool(d.adjProj,         projArr)  * 0.25
        + normPool(rawVal,            valueArr) * 0.20
        + normPool(d.trackHistScore,  trackArr) * 0.20
      );
    });

    pool.sort(function(a, b){ return b._ccGrade - a._ccGrade; });

    return pool.slice(0, 8).map(function(d) {
      var floorPerK = d.salary > 0 ? (d.floor / (d.salary / 1000)).toFixed(2) : "0";
      var val = "Grade " + d._ccGrade
              + " | Flr " + d.floor.toFixed(1)
              + " ($" + floorPerK + "/K)"
              + " | Hist " + (d.trackHistScore || 0).toFixed(0)
              + " | " + fmt$(d.salary);
      return toRow(d, "Cash Grade", val);
    });
  }

  if (stepId === "DOM") {
    var pool = drivers.filter(d => canAfford(d) && d.domPts > 0 && d.startPos <= 20);
    if (isCash) pool.sort((a, b) => (b.domPts + b.floor) - (a.domPts + a.floor));
    else pool.sort((a, b) => (b.domPts + b.ceiling + b.edge) - (a.domPts + a.ceiling + a.edge));
    return pool.slice(0, 10).map(d => {
      var label = isCash ? "Dom + Floor" : "Dom + Ceil";
      var val   = "Dom " + d.domPts.toFixed(1) + " | P" + d.startPos
        + (isCash ? " | Flr " + d.floor.toFixed(1) : " | Ceil " + d.ceiling.toFixed(1))
        + " | Own " + (d.ownPct * 100).toFixed(0) + "%";
      return toRow(d, label, val);
    });
  }

  if (stepId === "PD") {
    var pool = drivers.filter(d => canAfford(d) && d.group === "PD");
    if (isCash) {
      // Cash: deprioritize drivers with no historical avg finish data (histAvgFinish = 25 default)
      // No track history = unreliable floor estimate = poor cash play
      pool.sort((a, b) => {
        var aScore = (a.pdProj + a.floor) * (a.histAvgFinish !== 25 ? 1.0 : 0.5);
        var bScore = (b.pdProj + b.floor) * (b.histAvgFinish !== 25 ? 1.0 : 0.5);
        return bScore - aScore;
      });
    } else pool.sort((a, b) => (b.pdProj + b.ceiling + b.edge) - (a.pdProj + a.ceiling + a.edge));
    return pool.slice(0, 10).map(d => {
      var label = isCash ? "PD + Floor" : "PD + Upside";
      var val   = "PD " + d.pdProj.toFixed(1) + " | P" + d.startPos
        + (isCash ? " | Flr " + d.floor.toFixed(1) : " | Ceil " + d.ceiling.toFixed(1))
        + " | Own " + (d.ownPct * 100).toFixed(0) + "%";
      return toRow(d, label, val);
    });
  }

  if (stepId === "PUNT") {
    var pool = drivers.filter(d => canAfford(d) && d.salary <= 8000 && d.adjProj > 0);
    pool.sort((a, b) => {
      var aVal = (a.adjProj + a.floor) / (a.salary / 1000);
      var bVal = (b.adjProj + b.floor) / (b.salary / 1000);
      return bVal - aVal;
    });
    return pool.slice(0, 8).map(d => {
      var val = ((d.adjProj + d.floor) / (d.salary / 1000)).toFixed(2);
      return toRow(d, "Punt Value", val + " pts/$K | Flr " + d.floor.toFixed(1) + " | " + fmt$(d.salary));
    });
  }

  if (stepId === "DART") {
    var pool = drivers.filter(d => canAfford(d) && d.salary <= 8500 && d.ceiling > 0);
    pool.sort((a, b) => {
      var aVal = a.ceiling / (a.salary / 1000);
      var bVal = b.ceiling / (b.salary / 1000);
      return bVal - aVal;
    });
    return pool.slice(0, 8).map(d => {
      var val = (d.ceiling / (d.salary / 1000)).toFixed(2);
      return toRow(d, "Dart Value", val + " ceil/$K | Ceil " + d.ceiling.toFixed(1) + " | Own " + (d.ownPct * 100).toFixed(0) + "%");
    });
  }

  // FILL
  var pool = drivers.filter(d => canAfford(d) && (d.adjProj > 0 || d.ceiling > 0));
  if (isCash) pool.sort((a, b) => (b.adjProj + b.floor) - (a.adjProj + a.floor));
  else pool.sort((a, b) => (b.ceiling + b.edge) - (a.ceiling + a.edge));

  return pool.slice(0, 10).map(d => {
    var label = isCash ? "Proj + Floor" : "Ceil + Edge";
    var val   = isCash
      ? "Proj " + d.adjProj.toFixed(1) + " | Flr " + d.floor.toFixed(1)
      : "Ceil " + d.ceiling.toFixed(1) + " | Edge " + d.edge.toFixed(1);
    return toRow(d, label, val + " | " + fmt$(d.salary));
  });
}

function fmt$(n) { return "$" + (n || 0).toLocaleString(); }


/* -------------------------------------------------------
 *  14. Pools Sheet Generator
 * ------------------------------------------------------- */

function renderPools() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Pools");
  if (!sheet) sheet = ss.insertSheet("Pools");
  sheet.clear();

  const config    = ss.getSheetByName("Model_Config");
  const trackType = config ? config.getRange("B2").getValue() || "Intermediate" : "Intermediate";
  const raceName  = config ? config.getRange("B1").getValue() || "" : "";
  const fullSal   = CASH_SALARY_CAP;

  var row = 1;

  function writeTitle(text) {
    sheet.getRange(row, 1, 1, 7).merge().setValue(text)
      .setBackground("#1a1a2e").setFontColor("#fff").setFontWeight("bold").setFontSize(12);
    row++;
  }
  function writeSubtitle(text) {
    sheet.getRange(row, 1, 1, 7).merge().setValue(text)
      .setBackground("#f1f3f4").setFontColor("#546E7A").setFontSize(10).setFontStyle("italic");
    row++;
  }
  function writeStepHeader(label) {
    sheet.getRange(row, 1, 1, 7).merge().setValue(label)
      .setBackground("#283593").setFontColor("#fff").setFontWeight("bold").setFontSize(10);
    row++;
  }
  function writePoolHeaders(headers) {
    sheet.getRange(row, 1, 1, headers.length).setValues([headers])
      .setBackground("#e0e0e0").setFontWeight("bold").setFontSize(9);
    row++;
  }
  function writePoolRow(values) {
    sheet.getRange(row, 1, 1, values.length).setValues([values]).setFontSize(9);
    row++;
  }
  function writeBlank() { row++; }
  function getPool(stepId, ct) { return getStepPool(stepId, ct, [], fullSal); }

  var cashSteps = getWizardSteps("Cash");
  writeTitle("CASH LINEUP — " + raceName + " (" + trackType + ")");
  writeBlank();
  for (var i = 0; i < cashSteps.steps.length; i++) {
    var step = cashSteps.steps[i];
    var rangeLabel = step.min === step.max ? "" + step.max : step.min + "-" + step.max;
    writeStepHeader(step.label + "  (pick " + rangeLabel + ")");
    writeSubtitle(step.why);
    var pool = getPool(step.id, "Cash");
    if (pool.length > 0) {
      writePoolHeaders(["Driver", "Group", "Salary", "Start", "Own%", "Key Metric", ""]);
      pool.forEach(function(d) {
        writePoolRow([d.name, d.group || "", d.salary, "P" + d.startPos,
          (d.ownPct * 100).toFixed(1) + "%", d.keyLabel + ": " + d.keyVal, ""]);
      });
    } else {
      writeSubtitle("No eligible drivers for this step.");
    }
    writeBlank();
  }

  writeBlank();

  var gppSteps = getWizardSteps("GPP");
  writeTitle("GPP LINEUP — " + raceName + " (" + trackType + ")");
  writeBlank();
  for (var i = 0; i < gppSteps.steps.length; i++) {
    var step = gppSteps.steps[i];
    var rangeLabel = step.min === step.max ? "" + step.max : step.min + "-" + step.max;
    writeStepHeader(step.label + "  (pick " + rangeLabel + ")");
    writeSubtitle(step.why);
    var pool = getPool(step.id, "GPP");
    if (pool.length > 0) {
      writePoolHeaders(["Driver", "Group", "Salary", "Start", "Own%", "Key Metric", ""]);
      pool.forEach(function(d) {
        writePoolRow([d.name, d.group || "", d.salary, "P" + d.startPos,
          (d.ownPct * 100).toFixed(1) + "%", d.keyLabel + ": " + d.keyVal, ""]);
      });
    } else {
      writeSubtitle("No eligible drivers for this step.");
    }
    writeBlank();
  }

  sheet.getRange(1, 3, row, 1).setNumberFormat("$#,##0");
  for (var c = 1; c <= 7; c++) sheet.autoResizeColumn(c);
  sheet.setFrozenRows(0);
  SpreadsheetApp.getUi().alert("Pools sheet updated for " + trackType + ".");
}


/* -------------------------------------------------------
 *  15. Data Diagnostics
 * ------------------------------------------------------- */


/* -------------------------------------------------------
 *  Refresh Exposure Report
 *
 *  Called by Race Control sidebar refresh button.
 *  Recalculates the J-K exposure report from whatever
 *  lineups are currently in the Lineups sheet.
 * ------------------------------------------------------- */

function refreshExposureReport() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Lineups");
  if (!sheet) return { ok: false };
  updateExposureReport(sheet);
  return { ok: true };
}
function runDiagnostics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const data    = loadAllData();
  const drivers = data.drivers;
  const rc      = data.raceContext;

  let sheet = ss.getSheetByName("Diagnostics");
  if (!sheet) sheet = ss.insertSheet("Diagnostics");
  sheet.clear();

  const rows = [];

  function head(text) {
    rows.push(["", "", "", "", ""]);
    rows.push(["▶ " + text, "", "", "", ""]);
  }
  function row() { rows.push(Array.from(arguments)); }

  head("RACE CONTEXT");
  row("Race", rc.raceName);
  row("Track", rc.trackType);
  row("Laps", rc.laps);
  row("TargetDoms", JSON.stringify(rc.targetDoms));
  row("TargetPD", JSON.stringify(rc.targetPD));

  head("DRIVER COUNTS");
  row("Total drivers from DK_Salaries", drivers.length);

  head("SOURCE COVERAGE");
  row("Source / Field", "Matched", "Total", "Coverage %", "Note");

  var checks = [
    { label: "Playability — proj > 0",         fn: function(d){ return d.proj > 0; },              note: "ifrProjMid present" },
    { label: "Playability — floor > 0",        fn: function(d){ return d.floor > 0; },             note: "ifrProjLow present" },
    { label: "Playability — ceiling > 0",      fn: function(d){ return d.ceiling > 0; },           note: "ifrProjHigh present" },
    { label: "Playability — startPos ≠ 20",    fn: function(d){ return d.startPos !== 20; },       note: "non-default start pos" },
    { label: "Playability — domScore > 0",     fn: function(d){ return d.domScore > 0; },          note: "dominator rating present" },
    { label: "Playability — finProjLow > 0",   fn: function(d){ return d.finProjLow > 0; },        note: "finish projection present" },
    { label: "Playability — ifrRank > 0",      fn: function(d){ return d.ifrRank > 0; },           note: "fantasy rank present" },
    { label: "Practice — pracBestTime > 0",    fn: function(d){ return d.pracBestTime > 0; },      note: "practice speed present" },
    { label: "Qualifying — qualSpeed > 0",     fn: function(d){ return d.qualSpeed > 0; },         note: "qualifying speed present" },
    { label: "Qualifying — manufacturer",      fn: function(d){ return d.manufacturer !== ""; },   note: "car make present" },
    { label: "Loop — histPctLapsLed > 0",      fn: function(d){ return d.histPctLapsLed > 0; },    note: "laps led % present" },
    { label: "Loop — histFastLaps > 0",        fn: function(d){ return d.histFastLaps > 0; },      note: "fast laps present" },
    { label: "Loop — histRating > 0",          fn: function(d){ return d.histRating > 0; },        note: "driver rating present" },
    { label: "Avg Finish — ≠ default 25",      fn: function(d){ return d.histAvgFinish !== 25; },  note: "histAvgFinish present" },
    { label: "Ratings — ≠ default 20",         fn: function(d){ return d.histSkillRank !== 20; },  note: "histSkillRank present" },
    { label: "Green Speed — > 0",              fn: function(d){ return d.histGreenSpeed > 0; },    note: "green flag speed present" },
    { label: "Total Speed — consistency ≠ 50", fn: function(d){ return d.speedConsistency !== 50; }, note: "segment consistency present" },
    { label: "DK_Salaries — dkAvgFPPG > 0",       fn: function(d){ return d.dkAvgFPPG > 0; },               note: "DK historical FPPG present" },
    { label: "Avg Start — histAvgStart > 0",        fn: function(d){ return d.histAvgStart > 0; },             note: "avg start position present" },
    { label: "Avg Start — siteRaces > 0",           fn: function(d){ return d.siteRaces > 0; },               note: "site race count present" },
    { label: "Avg Start — histAvgStartFinishDiff",  fn: function(d){ return d.histAvgStartFinishDiff !== 0; }, note: "start/finish diff present" }
  ];

  for (var i = 0; i < checks.length; i++) {
    var c = checks[i];
    var matched = drivers.filter(c.fn).length;
    var pct = drivers.length > 0 ? Math.round(matched / drivers.length * 100) + "%" : "—";
    row(c.label, matched, drivers.length, pct, c.note);
  }

  head("MISSING / DEFAULT PLAYABILITY DATA (proj = 0 or floor = 0)");
  row("Driver", "Salary", "startPos", "proj", "floor");
  var missingPlay = drivers.filter(function(d){ return d.proj === 0 || d.floor === 0; });
  if (missingPlay.length === 0) {
    row("✅ All drivers have Playability projection data");
  } else {
    for (var i = 0; i < missingPlay.length; i++) {
      var d = missingPlay[i];
      row(d.name, d.salary, d.startPos, d.proj, d.floor);
    }
  }

  head("DRIVERS WITH startPos = 20");
  row("Driver", "Salary", "proj", "floor", "Note");
  var defaultStart = drivers.filter(function(d){ return d.startPos === 20; });
  if (defaultStart.length === 0) {
    row("✅ No drivers at default start position");
  } else {
    for (var i = 0; i < defaultStart.length; i++) {
      var d = defaultStart[i];
      row(d.name, d.salary, d.proj, d.floor,
          d.proj > 0 ? "Playability matched — P20 is real" : "Playability NOT matched");
    }
  }

  head("MISSING LOOP DATA (histPctLapsLed = 0 AND histFastLaps = 0)");
  row("Driver", "histPctLapsLed", "histFastLaps", "histRating", "histAvgPos");
  var missingLoop = drivers.filter(function(d){ return d.histPctLapsLed === 0 && d.histFastLaps === 0; });
  if (missingLoop.length === 0) {
    row("✅ All drivers have loop data");
  } else {
    for (var i = 0; i < missingLoop.length; i++) {
      var d = missingLoop[i];
      row(d.name, d.histPctLapsLed, d.histFastLaps, d.histRating, d.histAvgPos);
    }
  }

  head("MISSING AVG FINISH (histAvgFinish = 25 default)");
  row("Driver", "histAvgFinish");
  var missingFinish = drivers.filter(function(d){ return d.histAvgFinish === 25; });
  if (missingFinish.length === 0) {
    row("✅ All drivers have avg finish data");
  } else {
    for (var i = 0; i < missingFinish.length; i++) {
      row(missingFinish[i].name, missingFinish[i].histAvgFinish);
    }
  }

  head("LOW SITE RACE COUNT (siteRaces < 3 — reduced confidence)");
  row("Driver", "siteRaces", "histAvgStartFinishDiff", "note");
  var lowSiteRaces = drivers.filter(function(d){ return d.siteRaces > 0 && d.siteRaces < 3; });
  var noSiteHistory = drivers.filter(function(d){ return d.siteRaces === 0; });
  if (lowSiteRaces.length === 0 && noSiteHistory.length === 0) {
    row("✅ All drivers have 3+ site races");
  } else {
    for (var i = 0; i < lowSiteRaces.length; i++) {
      var d = lowSiteRaces[i];
      row(d.name, d.siteRaces, d.histAvgStartFinishDiff, "low confidence (0.70 factor)");
    }
    for (var i = 0; i < noSiteHistory.length; i++) {
      var d = noSiteHistory[i];
      row(d.name, 0, "—", "no site history (fallback pdProj)");
    }
  }

  head("SPOT CHECK — FIRST 5 DRIVERS");
  var spotFields = ["name","salary","startPos","proj","floor","ceiling","dkStd",
    "domScore","finProjLow","finProjHigh","ifrRank",
    "histPctLapsLed","histFastLaps","histAvgFinish","histSkillRank",
    "pracBestTime","qualSpeed","manufacturer","histGreenSpeed","speedConsistency","dkAvgFPPG","ownPct"];
  row.apply(null, spotFields);
  for (var i = 0; i < Math.min(5, drivers.length); i++) {
    var d = drivers[i];
    row.apply(null, spotFields.map(function(f){ return d[f] !== undefined ? d[f] : "—"; }));
  }

  if (rows.length === 0) return;

  var writeRows = rows.map(function(r){
    while (r.length < 5) r.push("");
    return r.slice(0, 5);
  });

  sheet.getRange(1, 1, writeRows.length, 5).setValues(writeRows);

  for (var i = 0; i < rows.length; i++) {
    var cell = String(rows[i][0]);
    if (cell.indexOf("▶") === 0) {
      sheet.getRange(i + 1, 1, 1, 5).setBackground("#1a1a2e").setFontColor("white").setFontWeight("bold");
    } else if (cell === "Source / Field" || cell === "Driver") {
      sheet.getRange(i + 1, 1, 1, 5).setBackground("#e0e0e0").setFontWeight("bold");
    }
    var pctCell = String(rows[i][3]);
    if (pctCell.indexOf("%") > 0) {
      var pct = parseInt(pctCell);
      if (pct === 0)     sheet.getRange(i + 1, 1, 1, 5).setBackground("#ffebee");
      else if (pct < 50) sheet.getRange(i + 1, 1, 1, 5).setBackground("#fff3cd");
    }
  }

  for (var c = 1; c <= 5; c++) sheet.autoResizeColumn(c);

  SpreadsheetApp.getUi().alert(
    "Diagnostics complete — " + drivers.length + " drivers.\n" +
    "See the Diagnostics sheet.\n\n" +
    "Yellow rows = < 50% coverage.\nRed rows = 0% coverage."
  );
}