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

// New column headers added to Race_Environment starting at col H (index 8)
const PR_HEADERS = [
  "P1 Held", "P2 Held", "P3 Held",
  "Top 5 Held", "Top 10 Held", "Top 12 Held",
  "Laps Led Leader", "DOM Hit Rate",
  "Band 1-5 Avg Finish", "Band 6-15 Avg Finish",
  "Band 16-30 Avg Finish", "Band 31+ Avg Finish",
  "Top 8 Mfr", "Top 12 Mfr", "Notes"
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

  prEnsureHeaders(raceEnv, envData[0]);

  const dashDrivers = prReadDashboard(ss);
  const metrics     = prComputeMetrics(finishDrivers, dashDrivers, raceLaps);

  const values = [
    metrics.p1Held,
    metrics.p2Held,
    metrics.p3Held,
    metrics.top5Held,
    metrics.top10Held,
    metrics.top12Held,
    metrics.lapsLedLeader,
    metrics.domHitRate,
    metrics.band1to5Avg,
    metrics.band6to15Avg,
    metrics.band16to30Avg,
    metrics.band31plusAvg,
    metrics.top8Mfr,
    metrics.top12Mfr,
    ""  // Notes — manually filled
  ];

  raceEnv.getRange(targetRow, PR_START_COL, 1, values.length).setValues([values]);

  SpreadsheetApp.getUi().alert("Race outcome logged for: " + currentRace);
}


/* -------------------------------------------------------
 *  2. Ensure Post-Race Headers in Race_Environment Row 1
 * ------------------------------------------------------- */

function prEnsureHeaders(raceEnv, headerRow) {
  // Only write if col H (index 7) is empty
  if (headerRow[PR_START_COL - 1]) return;
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
  // Lookup maps
  const finishByKey = {};
  for (const d of finishDrivers) finishByKey[d.key] = d;

  const dashByKey = {};
  for (const d of dashDrivers) dashByKey[d.key] = d;

  // Finish sorted ascending by position
  const sorted = finishDrivers.slice().sort((a, b) => a.pos - b.pos);

  // ---- P1/P2/P3 Held ----
  // "Held" = pole/front-row starter finished at or better than their start
  function pHeld(startPos) {
    const starter = finishDrivers.find(d => d.start === startPos);
    if (!starter) return "N/A";
    return starter.pos <= startPos ? "Y" : "N";
  }

  // ---- Top N Held ----
  // Count how many of the top-N starters finished inside the top N
  function topNHeld(n) {
    const starters = finishDrivers.filter(d => d.start >= 1 && d.start <= n);
    const held     = starters.filter(d => d.pos <= n).length;
    return held + "/" + n;
  }

  // ---- Laps Led Leader ----
  let lapsLedLeader = "—";
  const withLaps = finishDrivers.filter(d => d.led > 0);
  if (withLaps.length > 0) {
    const leader = withLaps.reduce((best, d) => d.led > best.led ? d : best, withLaps[0]);
    const total  = raceLaps > 0
      ? raceLaps
      : finishDrivers.reduce((s, d) => s + d.led, 0);
    const pct = total > 0 ? Math.round((leader.led / total) * 100) : 0;
    lapsLedLeader = prShortName(leader.name) + " - " + pct + "%";
  }

  // ---- DOM Hit Rate ----
  // DOM drivers who finished better (lower pos) than their start position
  const domDrivers = dashDrivers.filter(d => d.group === "DOM");
  let domHits = 0;
  for (const dom of domDrivers) {
    const result = finishByKey[dom.key];
    if (result && dom.startPos > 0 && result.pos < dom.startPos) domHits++;
  }
  const domHitRate = domDrivers.length > 0
    ? domHits + "/" + domDrivers.length
    : "—";

  // ---- Band Avg Finish ----
  function bandAvg(minStart, maxStart) {
    const band = finishDrivers.filter(d => d.start >= minStart && d.start <= maxStart);
    if (band.length === 0) return "—";
    const avg = band.reduce((s, d) => s + d.pos, 0) / band.length;
    return Math.round(avg * 10) / 10;
  }

  // ---- Manufacturer Breakdown ----
  function parseMfr(notes) {
    const m = notes.match(/MFR:([^\s|]+)/i);
    if (!m) return null;
    const raw = m[1].toLowerCase();
    if (raw.indexOf("chevy") >= 0 || raw.indexOf("chevrolet") >= 0) return "Chevy";
    if (raw.indexOf("ford") >= 0)                                    return "Ford";
    if (raw.indexOf("toyota") >= 0)                                  return "Toyota";
    return m[1]; // pass through unknown makes as-is
  }

  function topNMfr(n) {
    const counts = {};
    for (const d of sorted.slice(0, n)) {
      const dash = dashByKey[d.key];
      if (!dash) continue;
      const mfr = parseMfr(dash.notes);
      if (!mfr) continue;
      counts[mfr] = (counts[mfr] || 0) + 1;
    }
    const parts = Object.entries(counts)
      .sort((a, b) => b[1] - a[1])
      .map(([mfr, cnt]) => mfr + " " + cnt);
    return parts.length > 0 ? parts.join(" / ") : "—";
  }

  return {
    p1Held:        pHeld(1),
    p2Held:        pHeld(2),
    p3Held:        pHeld(3),
    top5Held:      topNHeld(5),
    top10Held:     topNHeld(10),
    top12Held:     topNHeld(12),
    lapsLedLeader: lapsLedLeader,
    domHitRate:    domHitRate,
    band1to5Avg:   bandAvg(1, 5),
    band6to15Avg:  bandAvg(6, 15),
    band16to30Avg: bandAvg(16, 30),
    band31plusAvg: bandAvg(31, 99),
    top8Mfr:       topNMfr(8),
    top12Mfr:      topNMfr(12)
  };
}


/* -------------------------------------------------------
 *  6. Utility: Short Driver Name
 *
 *  "Kyle Larson" → "Larson"
 *  "Martin Truex Jr." → "Truex Jr."
 *  "SVG" → "SVG" (already abbreviated)
 * ------------------------------------------------------- */

function prShortName(fullName) {
  if (!fullName) return "";
  const parts = fullName.trim().split(/\s+/);
  if (parts.length === 1) return parts[0];
  return parts.slice(1).join(" ");
}
