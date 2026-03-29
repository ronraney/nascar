/**
 * ============================================================
 *  NASCAR DFS 2026 — Optimizer
 * ============================================================
 *  Generates up to 150 GPP lineups using adjProj (or ceiling)
 *  as the projection base. Bell-curve randomization creates
 *  lineup diversity.
 *
 *  No forced group slots — pure score-driven selection.
 *  Exposure caps are based on wizard pool membership:
 *
 *    DOM   (d.group === "DOM")              → 65% default
 *    PD    (d.group === "PD")               → 50% default
 *    PUNT  (!d.group, salary ≤ $8,500)      → 35% default
 *    FILL  (everything else)                → 20% default
 *
 *  User can adjust each pool cap in the sidebar.
 *  Global max exposure slider is the ceiling over all pools.
 *
 *  Reads from Dashboard sheet. Appends to Lineups sheet.
 *  Writes Name (ID) format for DraftKings CSV upload.
 * ============================================================
 */

const OPT_SALARY_CAP  = 50000;
const OPT_ROSTER_SIZE = 6;
const OPT_MIN_SALARY  = 4500;


/* -------------------------------------------------------
 *  1. Sidebar Launcher
 * ------------------------------------------------------- */

function openOptimizer() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const config = ss.getSheetByName("Model_Config");
  const html   = HtmlService.createTemplateFromFile("Optimizer");

  if (config) {
    const trackType = config.getRange("B2").getValue() || "";
    const domT = trackType ? getTargetDominators(trackType) : { min: 0, max: 0 };
    const pdT  = trackType ? getTargetPD(trackType)         : { min: 0, max: 0 };
    html.raceName  = config.getRange("B1").getValue() || "No Race Selected";
    html.trackType = trackType || "Unknown";
    html.laps      = config.getRange("B3").getValue() || 0;
    html.domRange  = domT.min === domT.max ? "" + domT.max : domT.min + "-" + domT.max;
    html.pdRange   = pdT.min  === pdT.max  ? "" + pdT.max  : pdT.min  + "-" + pdT.max;
  } else {
    html.raceName  = "No Race Selected";
    html.trackType = "Unknown";
    html.laps      = 0;
    html.domRange  = "0";
    html.pdRange   = "0";
  }

  SpreadsheetApp.getUi().showSidebar(
    html.evaluate().setTitle("⚡ Optimizer").setWidth(320)
  );
}


/* -------------------------------------------------------
 *  2. Pool-Based Exposure Cap
 *
 *  Mirrors the wizard pool structure:
 *    DOM   → group === "DOM"
 *    PD    → group === "PD"
 *    PUNT  → ungrouped, salary ≤ $8,500 (covers both punt/dart)
 *    FILL  → everything else
 *
 *  User-supplied poolCaps override defaults.
 *  Global maxExposure is the ceiling over all pools.
 * ------------------------------------------------------- */

function getPoolCap(d, poolCaps, totalLineups, globalMax) {
  const defaults = {
    DOM:  0.65,
    PD:   0.50,
    PUNT: 0.35,
    FILL: 0.15   // hard cap — backmarkers capped at ~3/20 lineups
  };

  let poolKey;
  if (d.group === "DOM")     poolKey = "DOM";
  else if (d.group === "PD") poolKey = "PD";
  // PUNT: ungrouped value plays in realistic punt range ($6,000-$8,500)
  // Below $6,000 = backmarker territory → FILL cap regardless of salary
  else if (!d.group && d.salary >= 6000 && d.salary <= 8500) poolKey = "PUNT";
  else poolKey = "FILL";

  const capPct = (poolCaps && poolCaps[poolKey] !== undefined)
    ? poolCaps[poolKey]
    : defaults[poolKey];

  const effective = Math.min(capPct, globalMax || 1.0);
  return Math.max(1, Math.ceil(effective * totalLineups));
}


/* -------------------------------------------------------
 *  3. Bell-Curve Randomizer (Box-Muller transform)
 * ------------------------------------------------------- */

function gaussianRandom(mean, stdDev) {
  let u = 0, v = 0;
  while (u === 0) u = Math.random();
  while (v === 0) v = Math.random();
  const z = Math.sqrt(-2.0 * Math.log(u)) * Math.cos(2.0 * Math.PI * v);
  return mean + z * stdDev;
}

function applyNoise(drivers, optimizeBy, noiseLevel) {
  return drivers.map(d => {
    const base = optimizeBy === "ceiling"
      ? (d.ceiling || d.adjProj)
      : d.adjProj;
    const noise = noiseLevel > 0 ? gaussianRandom(1.0, noiseLevel) : 1.0;
    const mult  = Math.max(0.5, Math.min(2.0, noise));
    return Object.assign({}, d, { _score: base * mult });
  });
}


/* -------------------------------------------------------
 *  4. Lineup Builder (Single Attempt)
 *
 *  Pure score-sort with pool-based exposure caps.
 *  No forced group slots.
 * ------------------------------------------------------- */

function buildLineup(drivers, settings, exposureCounts, totalLineups, loosenFactor) {
  const lf        = loosenFactor || 1.0;
  const globalMax = settings.maxExposure || 0.65;
  const poolCaps  = settings.poolCaps || {};
  const lineup    = [];
  let   salaryUsed = 0;

  function isAtCap(d) {
    if (totalLineups <= 1) return false;
    // LoosenFactor only relaxes DOM and PD caps when the optimizer is struggling.
    // PUNT and FILL caps are hard — backmarkers should never slip in via loosening.
    const poolKey = d.group === "DOM" ? "DOM" :
                    d.group === "PD"  ? "PD"  : "FIXED";
    const appliedLF = (poolKey === "FIXED") ? 1.0 : lf;
    const cap = getPoolCap(d, poolCaps, totalLineups, globalMax) * appliedLF;
    return (exposureCounts[d.name] || 0) >= Math.ceil(cap);
  }

  function salaryFits(d) {
    const slotsLeft = OPT_ROSTER_SIZE - lineup.length;
    const remaining = OPT_SALARY_CAP - salaryUsed;
    if (d.salary > remaining) return false;
    if (slotsLeft > 1 && (remaining - d.salary) < (slotsLeft - 1) * OPT_MIN_SALARY) return false;
    return true;
  }

  // Sort by score descending, fill 6 slots
  const sorted = drivers.slice().sort((a, b) => b._score - a._score);

  for (const d of sorted) {
    if (lineup.length >= OPT_ROSTER_SIZE) break;
    if (isAtCap(d)) continue;
    if (!salaryFits(d)) continue;
    lineup.push(d);
    salaryUsed += d.salary;
  }

  return lineup.length === OPT_ROSTER_SIZE ? lineup : null;
}


/* -------------------------------------------------------
 *  5. Lineup Uniqueness Check
 * ------------------------------------------------------- */

function lineupKey(lineup) {
  return lineup.map(d => d.name).sort().join("|");
}


/* -------------------------------------------------------
 *  6. Main Optimizer
 *
 *  settings object:
 *    count            {number}   lineups to generate (1-150)
 *    optimizeBy       {string}   'adjProj' | 'ceiling'
 *    noiseLevel       {number}   0.0-0.50
 *    maxExposure      {number}   0.10-1.0 global ceiling
 *    poolCaps         {object}   { DOM, PD, PUNT, FILL } as fractions
 *    minUnique        {number}   min drivers that must differ vs every other lineup (0-5)
 *    minExposureNames {string[]} drivers requiring min exposure
 *    minExposureCount {number}   min lineups for those drivers
 * ------------------------------------------------------- */

function generateLineups(settings) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const allDrivers = readDashboardDrivers().filter(d => d.salary > 0 && d.adjProj > 0);
  if (allDrivers.length === 0) {
    return { ok: false, msg: "No driver data found. Run Refresh Dashboard first." };
  }

  const config = ss.getSheetByName("Model_Config");
  if (!config) return { ok: false, msg: "Model_Config sheet missing." };

  const count       = Math.min(Math.max(parseInt(settings.count) || 20, 1), 150);
  const noiseLevel  = parseFloat(settings.noiseLevel)  || 0.15;
  const optimizeBy  = settings.optimizeBy || "adjProj";
  const maxExposure = parseFloat(settings.maxExposure) || 0.65;
  const poolCaps    = settings.poolCaps || {};
  const minUnique   = Math.min(Math.max(parseInt(settings.minUnique) || 0, 0), 5);
  const minNames    = settings.minExposureNames || [];
  const minCount    = parseInt(settings.minExposureCount) || 0;

  const lineups        = [];
  const seenKeys       = new Set();
  const exposureCounts = {};
  let   attempts       = 0;
  const maxAttempts    = count * 300;

  while (lineups.length < count && attempts < maxAttempts) {
    attempts++;

    const loosenFactor = attempts > maxAttempts * 0.60 ? 1.75 :
                         attempts > maxAttempts * 0.30 ? 1.35 : 1.0;

    const noisyDrivers = applyNoise(allDrivers, optimizeBy, noiseLevel);
    const lineup = buildLineup(noisyDrivers, settings, exposureCounts, count, loosenFactor);
    if (!lineup) continue;

    const key = lineupKey(lineup);
    if (seenKeys.has(key)) continue;

    // Min unique check — new lineup must differ from every existing lineup
    // by at least minUnique drivers
    if (minUnique > 0) {
      const newNames = new Set(lineup.map(d => d.name));
      const tooSimilar = lineups.some(function(existing) {
        var overlap = existing.filter(function(d) { return newNames.has(d.name); }).length;
        return (OPT_ROSTER_SIZE - overlap) < minUnique;
      });
      if (tooSimilar) continue;
    }

    seenKeys.add(key);
    lineups.push(lineup);
    lineup.forEach(d => {
      exposureCounts[d.name] = (exposureCounts[d.name] || 0) + 1;
    });
  }

  if (lineups.length === 0) {
    return { ok: false, msg: "Could not generate any valid lineups. Try loosening exposure caps or running Refresh Dashboard." };
  }

  // Write to Lineups sheet
  const trackType  = config.getRange("B2").getValue() || "Unknown";
  const raceContext = { raceName: config.getRange("B1").getValue() || "Unknown", trackType };
  writeOptimizerLineups(ss, lineups, raceContext);

  // Build result message
  const expLines = Object.entries(exposureCounts)
    .sort((a, b) => b[1] - a[1])
    .map(([name, n]) => name + ": " + n + "/" + lineups.length
         + " (" + Math.round(n / lineups.length * 100) + "%)")
    .join("\n");

  let msg = "✅ Generated " + lineups.length + " lineup" + (lineups.length > 1 ? "s" : "")
          + " | " + attempts + " attempts\n"
          + "Track: " + trackType
          + (minUnique > 0 ? " | Min unique: " + minUnique : "")
          + "\n\nExposure:\n" + expLines;

  if (lineups.length < count) {
    msg = "⚠ Only " + lineups.length + "/" + count + " unique lineups found.\n\n" + msg;
  }

  if (minNames.length > 0 && minCount > 0) {
    const warnings = [];
    minNames.forEach(name => {
      const actual = exposureCounts[name] || 0;
      if (actual < minCount) warnings.push(name + ": " + actual + "/" + minCount);
    });
    if (warnings.length > 0) msg += "\n\n⚠ Min Exposure Not Met:\n" + warnings.join("\n");
  }

  return { ok: true, msg, count: lineups.length };
}


/* -------------------------------------------------------
 *  7. DK Name+ID Lookup Builder
 * ------------------------------------------------------- */

function buildDkIdMap(ss) {
  const sheet = ss.getSheetByName("DK_Salaries");
  if (!sheet) return {};

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return {};

  const h = data[0].map(v => v ? v.toString().toLowerCase().trim() : "");
  let nameIdx = h.findIndex(col => col === "name");
  if (nameIdx < 0) nameIdx = h.findIndex(col => col.indexOf("name") >= 0);
  const idIdx = h.findIndex(col => col === "id" || col === "dfs id");
  if (nameIdx < 0) return {};

  const map = {};
  for (let i = 1; i < data.length; i++) {
    const raw = data[i][nameIdx] ? data[i][nameIdx].toString().trim() : "";
    if (!raw) continue;
    const id  = idIdx >= 0 ? data[i][idIdx].toString().trim() : "";
    map[cleanName(raw)] = id ? raw + " (" + id + ")" : raw;
  }
  return map;
}


/* -------------------------------------------------------
 *  8. Write Lineups to Lineups Sheet
 * ------------------------------------------------------- */

function writeOptimizerLineups(ss, lineups, raceContext) {
  let sheet = ss.getSheetByName("Lineups");
  if (!sheet) sheet = ss.insertSheet("Lineups");
  if (sheet.getRange(1, 8).getValue() !== "Total Proj") {
    sheet.getRange(1, 1, 1, 9)
      .setValues([["D1", "D2", "D3", "D4", "D5", "D6", "Total Sal", "Total Proj", "Saved At"]])
      .setBackground("#1a1a2e").setFontColor("white").setFontWeight("bold")
      .setFontSize(9).setHorizontalAlignment("center");
  }

  const dkIdMap = buildDkIdMap(ss);
  function dkLabel(d) {
    if (!d || !d.name) return "";
    return dkIdMap[cleanName(d.name)] || d.name;
  }

  const colAValues = sheet.getRange(1, 1, sheet.getMaxRows(), 1).getValues();
  let nextRow = 1;
  for (let i = 0; i < colAValues.length; i++) {
    if (colAValues[i][0] !== "") nextRow = i + 2;
  }

  const timestamp = new Date().toLocaleString();
  const rows = lineups.map(lineup => {
    const sorted    = lineup.slice().sort((a, b) => b.salary - a.salary);
    const totalSal  = sorted.reduce((s, d) => s + d.salary, 0);
    const totalProj = Math.round(sorted.reduce((s, d) => s + (d.adjProj || 0), 0) * 10) / 10;
    return [
      dkLabel(sorted[0]), dkLabel(sorted[1]), dkLabel(sorted[2]),
      dkLabel(sorted[3]), dkLabel(sorted[4]), dkLabel(sorted[5]),
      totalSal, totalProj, timestamp
    ];
  });

  if (rows.length === 0) return;

  const writeRange = sheet.getRange(nextRow, 1, rows.length, 9);
  writeRange.setValues(rows).setFontSize(9).setHorizontalAlignment("center");
  sheet.getRange(nextRow, 7, rows.length).setNumberFormat("$#,##0");
  sheet.getRange(nextRow, 8, rows.length).setNumberFormat("0.0");

  for (let i = 0; i < lineups.length; i++) {
    const sorted = lineups[i].slice().sort((a, b) => b.salary - a.salary);
    for (let j = 0; j < sorted.length; j++) {
      const cell = sheet.getRange(nextRow + i, j + 1);
      if (sorted[j].group === "DOM")     cell.setBackground("#ffe0e0").setFontWeight("bold");
      else if (sorted[j].group === "PD") cell.setBackground("#e0f7f7").setFontWeight("bold");
    }
  }

  updateLineupsExposure(ss, sheet);
}


/* -------------------------------------------------------
 *  9. Lineups Exposure Report (Cols J-K)
 * ------------------------------------------------------- */

function updateLineupsExposure(ss, sheet) {
  if (!sheet) sheet = ss.getSheetByName("Lineups");
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  const counts = {};
  for (const row of data) {
    for (let c = 0; c < 6; c++) {
      const val = row[c] ? row[c].toString().trim() : "";
      if (!val || val === "D1" || val === "D2") continue;
      counts[val] = (counts[val] || 0) + 1;
    }
  }

  const totalLineups = data.filter(r => r[0] && r[0].toString().trim()).length;
  if (totalLineups === 0) return;

  const sorted = Object.entries(counts).sort((a, b) => b[1] - a[1]);

  sheet.getRange(1, 10, sheet.getMaxRows(), 2).clearContent().clearFormat();
  sheet.getRange(1, 10, 1, 2)
    .setValues([["Driver", "Exp %"]])
    .setBackground("#1a73e8").setFontColor("white").setFontWeight("bold").setFontSize(9);

  if (sorted.length === 0) return;

  const expRows = sorted.map(([name, n]) => [name, Math.round(n / totalLineups * 100) + "%"]);
  sheet.getRange(2, 10, expRows.length, 2).setValues(expRows).setFontSize(9).setHorizontalAlignment("center");
  sheet.getRange(2, 10, expRows.length, 1).setHorizontalAlignment("left");

  for (let i = 0; i < sorted.length; i++) {
    const pct = sorted[i][1] / totalLineups;
    const bg  = pct > 0.65 ? "#FFEBEE" :
                pct >= 0.50 ? "#FFF8E1" :
                pct >= 0.25 ? "#E8F5E9" :
                              "#F3E5F5";
    sheet.getRange(i + 2, 11).setBackground(bg);
  }

  sheet.autoResizeColumn(10);
}