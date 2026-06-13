/**
 * ============================================================
 *  NASCAR DFS 2026 — Module 6: POST RACE
 * ============================================================
 *  Reads post-race results from the Finish sheet and pre-race
 *  predictions from Dashboard, then writes computed analytics
 *  to the current race's row in Race_Environment (cols H+).
 *
 *  Triggered by "Log Race Outcome" in the Race Control sidebar.
 *  Depends on: Config.js (cleanName, DASH_COLS)
 * ============================================================
 */


/* -------------------------------------------------------
 *  Sheet name constants (mirroring usage elsewhere)
 * ------------------------------------------------------- */

const PR_SHEETS = {
  FINISH:       "Finish",
  DASHBOARD:    "Dashboard",
  RACE_ENV:     "Race_Environment",
  MODEL_CONFIG: "Model_Config"
};

// Column headers added to Race_Environment starting at col H
const PR_HEADERS = [
  "P1 Held", "P2 Held", "P3 Held",
  "Top 5 Held", "Top 10 Held", "Top 12 Held",
  "Laps Led Driver", "Laps Led Pct",
  "DOM Laps Led Count", "DOM Total Laps Led",
  "PD Hit Rate", "PD Hits", "PD Total",
  "Best PD Driver", "Best PD Gain", "Tagged PD Avg S/F",
  "Big Movers 10+",
  "Band 1-5 Avg Finish", "Band 6-15 Avg Finish",
  "Band 16-30 Avg Finish", "Band 31+ Avg Finish",
  "Top 8 Chevy", "Top 8 Ford", "Top 8 Toyota",
  "Top 12 Chevy", "Top 12 Ford", "Top 12 Toyota",
  "Notes"
];

const PR_START_COL = 8; // cols A-G are pre-race; new cols start at H


/* -------------------------------------------------------
 *  1. Entry Point
 * ------------------------------------------------------- */

function logRaceOutcome() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const config = ss.getSheetByName(PR_SHEETS.MODEL_CONFIG);

  const currentRace = config ? config.getRange("B1").getValue() : "";
  const raceLaps    = config ? (parseFloat(config.getRange("B3").getValue()) || 0) : 0;

  if (!currentRace) {
    SpreadsheetApp.getUi().alert("No race selected in Model_Config. Load a race first.");
    return;
  }

  const finishDrivers = prReadFinishSheet(ss);
  if (!finishDrivers) {
    SpreadsheetApp.getUi().alert("Finish sheet not found or has no data.\n\nPaste race results into a sheet named \"Finish\" with columns: Pos, St, Driver, Laps, Led.");
    return;
  }

  const raceEnv = ss.getSheetByName(PR_SHEETS.RACE_ENV);
  if (!raceEnv) {
    SpreadsheetApp.getUi().alert("Race_Environment sheet not found.");
    return;
  }

  // Find the row matching the current race name (col C = index 2, 0-based)
  const envData  = raceEnv.getDataRange().getValues();
  let targetRow  = -1;
  for (let i = 1; i < envData.length; i++) {
    if (String(envData[i][2]).trim() === String(currentRace).trim()) {
      targetRow = i + 1; // 1-indexed sheet row
      break;
    }
  }

  if (targetRow < 0) {
    SpreadsheetApp.getUi().alert("Race \"" + currentRace + "\" not found in Race_Environment.\nAdd it to the schedule first.");
    return;
  }

  prEnsureHeaders(raceEnv);

  const dashDrivers = prReadDashboard(ss);
  const m           = prComputeMetrics(finishDrivers, dashDrivers, raceLaps);

  const values = [
    m.p1Held, m.p2Held, m.p3Held,
    m.top5Held, m.top10Held, m.top12Held,
    m.lapsLedDriver, m.lapsLedPct,
    m.domLapsLedCount, m.domTotalLapsLed,
    m.pdHitRate, m.pdHits, m.pdTotal,
    m.bestPdDriver, m.bestPdGain, m.taggedPdAvgSF,
    m.bigMovers10Plus,
    m.band1to5Avg, m.band6to15Avg, m.band16to30Avg, m.band31plusAvg,
    m.top8Chevy, m.top8Ford, m.top8Toyota,
    m.top12Chevy, m.top12Ford, m.top12Toyota,
    ""  // Notes — manually filled
  ];

  raceEnv.getRange(targetRow, PR_START_COL, 1, values.length).setValues([values]);

  SpreadsheetApp.getUi().alert("Race outcome logged for: " + currentRace);
}


/* -------------------------------------------------------
 *  2. Ensure Post-Race Headers in Race_Environment Row 1
 *
 *  Always overwrites so schema changes are picked up immediately.
 * ------------------------------------------------------- */

function prEnsureHeaders(raceEnv) {
  raceEnv.getRange(1, PR_START_COL, 1, PR_HEADERS.length)
    .setValues([PR_HEADERS])
    .setFontWeight("bold");
}


/* -------------------------------------------------------
 *  3. Finish Sheet Reader
 *
 *  Expected columns: Pos, St, Driver, Laps, Led
 *  Returns array of { pos, start, name, key, laps, led }
 *  or null if sheet is missing / empty.
 * ------------------------------------------------------- */

function prReadFinishSheet(ss) {
  const sheet = ss.getSheetByName(PR_SHEETS.FINISH);
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;

  const h = data[0].map(v => v ? v.toString().toLowerCase().trim() : "");

  const idx = {
    pos:    h.findIndex(col => col === "pos" || col === "fin" || col === "finish"),
    start:  h.findIndex(col => col === "st"  || col === "start" || col === "grid"),
    driver: h.findIndex(col => col === "driver" || col.indexOf("driver") >= 0),
    laps:   h.findIndex(col => col === "laps"),
    led:    h.findIndex(col => col === "led")
  };

  if (idx.pos < 0 || idx.driver < 0) {
    Logger.log("PostRace: Finish sheet missing Pos or Driver column. Headers: " + h.join(", "));
    return null;
  }

  const drivers = [];
  for (let i = 1; i < data.length; i++) {
    const rawName = data[i][idx.driver] ? data[i][idx.driver].toString().trim() : "";
    if (!rawName) continue;

    const pos = parseInt(data[i][idx.pos]) || 0;
    if (pos <= 0) continue;

    drivers.push({
      pos:   pos,
      start: idx.start >= 0 ? (parseInt(data[i][idx.start]) || 0) : 0,
      name:  rawName,
      key:   cleanName(rawName),
      laps:  idx.laps >= 0 ? (parseInt(data[i][idx.laps])  || 0) : 0,
      led:   idx.led  >= 0 ? (parseInt(data[i][idx.led])   || 0) : 0
    });
  }

  return drivers.length > 0 ? drivers : null;
}


/* -------------------------------------------------------
 *  4. Dashboard Reader (post-race subset)
 *
 *  Returns array of { name, key, startPos, group, notes }
 * ------------------------------------------------------- */

function prReadDashboard(ss) {
  const sheet = ss.getSheetByName(PR_SHEETS.DASHBOARD);
  if (!sheet) return [];

  const dRow = DASH_COLS.GPP_DATA_START;
  const last = sheet.getLastRow();
  if (last < dRow) return [];

  const data = sheet.getRange(dRow, 1, last - dRow + 1, DASH_COLS.TOTAL_COLS).getValues();

  return data
    .map(row => ({
      name:     String(row[DASH_COLS.COL_DRIVER - 1] || ""),
      key:      cleanName(String(row[DASH_COLS.COL_DRIVER - 1] || "")),
      startPos: parseFloat(row[DASH_COLS.COL_START  - 1]) || 0,
      group:    String(row[DASH_COLS.COL_GROUP  - 1] || ""),
      notes:    String(row[DASH_COLS.COL_NOTES  - 1] || "")
    }))
    .filter(d => d.name.trim() !== "");
}


/* -------------------------------------------------------
 *  5. Metrics Calculator
 * ------------------------------------------------------- */

function prComputeMetrics(finishDrivers, dashDrivers, raceLaps) {
  const finishByKey = {};
  for (const d of finishDrivers) finishByKey[d.key] = d;

  const dashByKey = {};
  for (const d of dashDrivers) dashByKey[d.key] = d;

  const sorted = finishDrivers.slice().sort((a, b) => a.pos - b.pos);

  // ---- P1/P2/P3 Held ----
  function pHeld(startPos) {
    const starter = finishDrivers.find(d => d.start === startPos);
    if (!starter) return "N/A";
    return starter.pos <= startPos ? "Y" : "N";
  }

  // ---- Top N Held (integer count) ----
  function topNHeld(n) {
    const starters = finishDrivers.filter(d => d.start >= 1 && d.start <= n);
    if (starters.length === 0) return "";
    return starters.filter(d => d.pos <= n).length;
  }

  // ---- Laps Led Driver + Pct ----
  let lapsLedDriver = "";
  let lapsLedPct    = "";
  const withLaps = finishDrivers.filter(d => d.led > 0);
  if (withLaps.length > 0) {
    const leader = withLaps.reduce((best, d) => d.led > best.led ? d : best, withLaps[0]);
    const total  = raceLaps > 0
      ? raceLaps
      : finishDrivers.reduce((s, d) => s + d.led, 0);
    lapsLedDriver = prShortName(leader.name);
    lapsLedPct    = total > 0 ? Math.round((leader.led / total) * 100) / 100 : "";
  }

  // ---- DOM: count who led + total laps led ----
  const domDrivers = dashDrivers.filter(d => d.group.includes("DOM"));
  let domLapsLedCount  = 0;
  let domTotalLapsLed  = 0;
  for (const dom of domDrivers) {
    const result = finishByKey[dom.key];
    if (!result) continue;
    if (result.led > 0) domLapsLedCount++;
    domTotalLapsLed += result.led;
  }

  // ---- PD metrics ----
  // Hit = PD driver gained positions (finished better than they started)
  const pdDrivers = dashDrivers.filter(d => d.group.includes("PD"));
  let pdHits       = 0;
  let bestPdDriver = "";
  let bestPdGain   = 0;
  let pdSFSum      = 0;
  let pdSFCount    = 0;

  for (const pd of pdDrivers) {
    const result = finishByKey[pd.key];
    if (!result || pd.startPos <= 0) continue;
    const gain = pd.startPos - result.pos; // positive = gained positions
    pdSFSum += gain;
    pdSFCount++;
    if (gain > 0) pdHits++;
    if (gain > bestPdGain) {
      bestPdGain   = gain;
      bestPdDriver = prShortName(pd.name);
    }
  }

  const pdTotal      = pdDrivers.length;
  const pdHitRate    = pdTotal > 0 ? Math.round((pdHits / pdTotal) * 100) / 100 : "";
  const taggedPdAvgSF = pdSFCount > 0 ? Math.round((pdSFSum / pdSFCount) * 10) / 10 : "";

  // ---- Big Movers 10+ ----
  const bigMovers10Plus = finishDrivers.filter(d => d.start > 0 && (d.start - d.pos) >= 10).length;

  // ---- Band Avg Finish ----
  function bandAvg(minStart, maxStart) {
    const band = finishDrivers.filter(d => d.start >= minStart && d.start <= maxStart);
    if (band.length === 0) return "";
    return Math.round(band.reduce((s, d) => s + d.pos, 0) / band.length * 10) / 10;
  }

  // ---- Manufacturer counts by finish group ----
  function parseMfr(notes) {
    const m = notes.match(/MFR:([^\s|]+)/i);
    if (!m) return null;
    const raw = m[1].toLowerCase();
    if (raw.indexOf("chevy") >= 0 || raw.indexOf("chevrolet") >= 0) return "Chevy";
    if (raw.indexOf("ford") >= 0)                                    return "Ford";
    if (raw.indexOf("toyota") >= 0)                                  return "Toyota";
    return null;
  }

  function mfrCounts(n) {
    const counts = { Chevy: 0, Ford: 0, Toyota: 0 };
    for (const d of sorted.slice(0, n)) {
      const dash = dashByKey[d.key];
      if (!dash) continue;
      const mfr = parseMfr(dash.notes);
      if (mfr && mfr in counts) counts[mfr]++;
    }
    return counts;
  }

  const mfr8  = mfrCounts(8);
  const mfr12 = mfrCounts(12);

  return {
    p1Held:           pHeld(1),
    p2Held:           pHeld(2),
    p3Held:           pHeld(3),
    top5Held:         topNHeld(5),
    top10Held:        topNHeld(10),
    top12Held:        topNHeld(12),
    lapsLedDriver:    lapsLedDriver,
    lapsLedPct:       lapsLedPct,
    domLapsLedCount:  domLapsLedCount,
    domTotalLapsLed:  domTotalLapsLed,
    pdHitRate:        pdHitRate,
    pdHits:           pdHits,
    pdTotal:          pdTotal,
    bestPdDriver:     bestPdDriver,
    bestPdGain:       bestPdGain > 0 ? bestPdGain : "",
    taggedPdAvgSF:    taggedPdAvgSF,
    bigMovers10Plus:  bigMovers10Plus,
    band1to5Avg:      bandAvg(1, 5),
    band6to15Avg:     bandAvg(6, 15),
    band16to30Avg:    bandAvg(16, 30),
    band31plusAvg:    bandAvg(31, 99),
    top8Chevy:        mfr8.Chevy,
    top8Ford:         mfr8.Ford,
    top8Toyota:       mfr8.Toyota,
    top12Chevy:       mfr12.Chevy,
    top12Ford:        mfr12.Ford,
    top12Toyota:      mfr12.Toyota
  };
}


/* -------------------------------------------------------
 *  6. Utility: Short Driver Name
 *
 *  "Kyle Larson"      → "Larson"
 *  "Martin Truex Jr." → "Truex Jr."
 *  "SVG"              → "SVG"
 * ------------------------------------------------------- */

function prShortName(fullName) {
  if (!fullName) return "";
  const parts = fullName.trim().split(/\s+/);
  if (parts.length === 1) return parts[0];
  return parts.slice(1).join(" ");
}
