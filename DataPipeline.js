/**
 * ============================================================
 *  NASCAR DFS 2026 — Module 2: DATA PIPELINE
 * ============================================================
 *  Reads every data sheet and produces a single unified array
 *  of driver objects for the Analysis module to process.
 *
 *  Data sources loaded:
 *    - DK_Salaries      (DraftKings CSV: roster, salary, AvgFPPG, Own%)
 *    - Playability      (iFantasyRace: start pos, proj range, dominator, rank)
 *    - Driver_Names     (alias map: resolves name format differences)
 *    - Data_Loop        (historical race stats — secondary)
 *    - Data_Avg_Finish  (multi-season avg finish, recency-weighted)
 *    - Data_Ratings     (multi-season skill ranks, recency-weighted)
 *    - practice_1       (current-week best lap times / speed)
 *    - qualifying_1     (qualifying speed + manufacturer)
 *    - Data_Green_Speed (green flag speed — historical race pace at track)
 *    - Data_Total_Speed (segment speed ranks — consistency signal)
 *
 *  Depends on: Config.gs (cleanName, getRecencyWeight, DK constants)
 * ============================================================
 */


/* -------------------------------------------------------
 *  1. Master Data Loader
 *
 *  Call this once per analysis run. Returns:
 *    { drivers: [...], raceContext: {...} }
 *
 *  DK_Salaries is the roster source of record.
 *  Every driver in DK_Salaries gets an object.
 *  All other sources are merged in by canonical key.
 * ------------------------------------------------------- */

function loadAllData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ---- Load race context from Model_Config ----
  const config = ss.getSheetByName("Model_Config");
  if (!config) throw new Error("Missing Model_Config sheet.");
  const raceContext = loadRaceContext(ss, config);

  // ---- Load name alias map first (used by all subsequent loaders) ----
  const nameMap = loadNameMap(ss);

  // ---- Load roster + salary from DraftKings CSV ----
  const dkDrivers = loadDKSalaries(ss, nameMap);
  if (dkDrivers.length === 0) {
    throw new Error("DK_Salaries sheet is empty or missing. Paste DraftKings CSV before running.");
  }

  // ---- Load iFantasyRace Playability (projection authority) ----
  const playMap = loadPlayability(ss, nameMap);

  // ---- Load all supplementary sources into lookup maps ----
  const loopMap       = loadLoopData(ss, nameMap);
  const avgFinishMap  = loadAvgFinish(ss, nameMap);
  const avgStartMap   = loadAvgStart(ss, nameMap);
  const ratingsMap    = loadRatings(ss, nameMap);
  const practiceMap   = loadPractice(ss, nameMap);
  const qualifyingMap = loadQualifying(ss, nameMap);
  const greenSpeedMap = loadGreenSpeed(ss, nameMap);
  const totalSpeedMap = loadTotalSpeed(ss, nameMap);

  // ---- Merge everything into unified driver objects ----
  const drivers = [];

  for (const dk of dkDrivers) {
    const key = dk.key;

    const play = playMap[key]       || {};
    const loop = loopMap[key]       || {};
    const hist = avgFinishMap[key]  || {};
    const avgS = avgStartMap[key]   || {};
    const skill = ratingsMap[key]   || {};
    const prac = practiceMap[key]   || {};
    const qual = qualifyingMap[key] || {};
    const gSpd = greenSpeedMap[key] || {};
    const tSpd = totalSpeedMap[key] || {};

    // --- iFantasyRace projection range ---
    // ifrProjMid is the anchor projection (replaces SaberSim SS Proj)
    // ifrProjLow is the floor, ifrProjHigh is the ceiling
    // Fall back to dkAvgFPPG if iFantasyRace data is missing
    const ifrMid  = play.ifrProjMid  || dk.dkAvgFPPG || 0;
    const ifrLow  = play.ifrProjLow  || (ifrMid * 0.75) || 0;
    const ifrHigh = play.ifrProjHigh || (ifrMid * 1.35) || 0;

    drivers.push({
      // --- Identity ---
      name:         dk.name,
      key:          key,
      salary:       dk.salary,
      dkId:         dk.dkId,

      // --- DraftKings historical average (season FPPG — weak baseline) ---
      dkAvgFPPG:    dk.dkAvgFPPG,

      // --- Ownership (manual entry in DK_Salaries Own% column) ---
      // 0 means not entered — triggers ownership-degraded mode in Analysis
      ownPct:       dk.ownPct,

      // --- iFantasyRace projections (primary projection source) ---
      proj:         ifrMid,           // anchor projection (was ssProj)
      floor:        ifrLow,           // low end of scoring range
      ceiling:      ifrHigh,          // high end of scoring range
      dkStd:        (ifrHigh - ifrMid) / 2,  // spread approximation

      // --- iFantasyRace supplementary fields ---
      ifrRank:      play.ifrRank     || 0,    // iFantasyRace Fantasy Rank
      domScore:     play.domScore    || 0,    // dominator rating (High=4..Low=1)
      domLabel:     play.domLabel    || "",   // "High", "Medium-High", etc.
      finProjLow:   play.finProjLow  || 0,    // finish projection low end
      finProjHigh:  play.finProjHigh || 0,    // finish projection high end

      // --- Starting position (from iFantasyRace Playability) ---
      // Falls back to mid-field if Playability data is missing
      startPos:     play.startPos    || 20,

      // --- Historical Loop Data (backward-looking, secondary) ---
      // histPctLapsLed is a percentage (0-100). Analysis.gs converts to
      // projected lap count via: histPctLapsLed / 100 * raceContext.laps
      histPctLapsLed: loop.pctLapsLed || 0,
      histFastLaps:   loop.fastLaps   || 0,
      histTop15Pct:   loop.top15Pct   || 0,
      histRating:     loop.rating     || 0,
      histAvgPos:     loop.avgPos     || 0,

      // --- Historical Avg Finish (recency-weighted) ---
      histAvgFinish: hist.weightedAvg || 25,

      // --- Historical Avg Start + Site Stats (from Data_Avg_Start) ---
      histAvgStart:           avgS.avgStart            || 0,
      histAvgStartFinishDiff: avgS.avgStartFinishDiff  || 0,
      siteLapsLed:            avgS.siteLapsLed         || 0,
      siteRaces:              avgS.siteRaces           || 0,
      siteAvgRating:          avgS.siteAvgRating       || 0,

      // --- Historical Skill Rating (recency-weighted) ---
      histSkillRank: skill.weightedAvg || 20,

      // --- Practice Speed (current week) ---
      pracBestTime:  prac.bestTime  || 0,
      pracSpeedRank: 0,               // computed after all drivers loaded

      // --- Green Flag Speed (historical race pace at this track) ---
      histGreenSpeed:   gSpd.greenFlagSpeed || 0,
      histGreenRank:    0,            // computed after all drivers loaded

      // --- Total Speed / Consistency (segment variance at this track) ---
      histAvgSpeedRank: tSpd.avgSpeedRank  || 0,
      speedConsistency: tSpd.consistency   || 50,  // 0-100, higher = more consistent

      // --- Composite Speed Score (blended, filled by post-processing) ---
      speedComposite:   0,            // 1 = fastest, higher number = slower

      // --- Qualifying / Manufacturer ---
      qualSpeed:    qual.speed       || 0,
      manufacturer: qual.car         || "",

      // --- Computed fields (filled by Analysis module) ---
      domPts:       0,
      domRank:      0,
      pdProj:       0,
      adjProj:      0,
      edge:         0,
      value:        0,
      group:        "",
      minExp:         0,
      maxExp:         0,
      cashScore:      0,
      trackHistScore: 0,
      cashCoreGrade:  0,
      notes:          []
    });
  }

  // ---- Post-processing: speed ranks and composite ----
  computePracticeRanks(drivers);
  computeGreenSpeedRanks(drivers);
  computeSpeedComposite(drivers);

  // ---- Add notes from available data ----
  for (const d of drivers) {
    if (d.manufacturer)  d.notes.push("MFR:" + d.manufacturer);
    if (d.ifrRank > 0)   d.notes.push("iFR:#" + d.ifrRank);
    if (d.domLabel)      d.notes.push("DOM:" + d.domLabel);
  }

  // ---- Log any DK drivers with no Playability data (actionable warning) ----
  const missingPlay = drivers.filter(d => !playMap[d.key]);
  if (missingPlay.length > 0) {
    Logger.log("WARNING: " + missingPlay.length + " DK drivers have no Playability data: "
      + missingPlay.map(d => d.name).join(", "));
  }

  return { drivers, raceContext };
}


/* -------------------------------------------------------
 *  2. Race Context Loader
 * ------------------------------------------------------- */

function loadRaceContext(ss, config) {
  const raceName  = config.getRange("B1").getValue() || "No Race Selected";
  const trackType = config.getRange("B2").getValue() || "Intermediate";
  const laps      = config.getRange("B3").getValue() || 267;
  const strategy  = config.getRange("B4").getValue() || "";

  const weights    = getWeights(trackType);
  const targetDoms = getTargetDominators(trackType);
  const targetPD   = getTargetPD(trackType);
  const domAvail   = calcDomPointsAvailable(laps, trackType);

  return { raceName, trackType, laps, strategy, weights, targetDoms, targetPD, domAvail };
}


/* -------------------------------------------------------
 *  3. Name Alias Map Loader
 *
 *  Reads the Driver_Names sheet and builds a reverse lookup:
 *    cleanName(anyAlias) → canonicalKey
 *
 *  Sheet columns: Canonical Key | DK Name | iFantasyRace Name | ...
 *
 *  All loaders pass names through resolveKey() which applies
 *  this map before falling back to raw cleanName().
 *  This handles format differences across sources without
 *  requiring changes to cleanName() itself.
 * ------------------------------------------------------- */

function loadNameMap(ss) {
  const sheet = ss.getSheetByName("Driver_Names");
  if (!sheet) {
    Logger.log("NOTE: Driver_Names sheet not found — using cleanName() only.");
    return {};
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return {};

  const map = {};

  // Row 0 is the header. Columns: [0] Canonical Key, [1] DK Name, [2] iFantasyRace Name
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const canonical = row[0] ? row[0].toString().trim() : "";
    if (!canonical) continue;

    // Register every non-empty alias column against this canonical key
    // Skip the header row and the "Alias Needed" / "Notes" columns
    for (let c = 1; c <= 2; c++) {
      const alias = row[c] ? row[c].toString().trim() : "";
      if (!alias || alias.startsWith("—") || alias.startsWith("-")) continue;
      const aliasKey = cleanName(alias);
      if (aliasKey) {
        map[aliasKey] = canonical;
      }
    }
  }

  Logger.log("Name map loaded: " + Object.keys(map).length + " alias entries.");
  return map;
}


/* -------------------------------------------------------
 *  Utility: resolveKey
 *
 *  Given a raw name string and the nameMap, returns the
 *  canonical key to use for cross-sheet matching.
 *  Falls back to cleanName() if no alias entry exists.
 * ------------------------------------------------------- */

function resolveKey(rawName, nameMap) {
  if (!rawName) return "";
  const cleaned = cleanName(rawName);
  return (nameMap && nameMap[cleaned]) ? nameMap[cleaned] : cleaned;
}


/* -------------------------------------------------------
 *  4. DK_Salaries Loader (DraftKings CSV)
 *
 *  Replaces loadWeeklyLive. Reads the DraftKings player
 *  export pasted into the DK_Salaries sheet.
 *
 *  Required DK columns: Name, Salary
 *  Optional DK column:  AvgPointsPerGame
 *  Manual column:       Own% (added by user before running)
 *
 *  Uses fuzzy header matching to tolerate column order
 *  changes and whitespace differences.
 *
 *  Returns an array of { name, key, salary, dkId, dkAvgFPPG, ownPct }
 * ------------------------------------------------------- */

function loadDKSalaries(ss, nameMap) {
  const sheet = ss.getSheetByName("DK_Salaries");
  if (!sheet) {
    Logger.log("WARNING: DK_Salaries sheet not found.");
    return [];
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const h = data[0].map(v => v ? v.toString().toLowerCase().trim() : "");

  // Fuzzy column finder — finds the first column whose header contains the target string
  function findCol(target) {
    return h.findIndex(col => col.indexOf(target) >= 0);
  }

  const idx = {
    name:   findCol("name"),          // "Name" (not "Name + ID")
    salary: findCol("salary"),
    dkId:   findCol("id"),            // "ID" column
    fppg:   findCol("avgpointspergame"),
    own:    findCol("own")            // manual "Own%" column, if present
  };

  // "Name + ID" header also contains "name" — prefer the standalone Name column
  // If both match, pick the one whose header is exactly "name"
  if (idx.name >= 0) {
    const exactName = h.findIndex(col => col === "name");
    if (exactName >= 0) idx.name = exactName;
  }

  if (idx.name < 0 || idx.salary < 0) {
    Logger.log("ERROR: DK_Salaries missing required Name or Salary column. Headers found: " + h.join(", "));
    return [];
  }

  const drivers = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rawName = row[idx.name] ? row[idx.name].toString().trim() : "";
    if (!rawName) continue;

    // Skip header-like rows that may have been pasted in
    if (rawName.toLowerCase() === "name") continue;

    const salary = parseFloat(row[idx.salary]) || 0;
    if (salary === 0) continue;  // no salary = not a real player row

    let ownPct = 0;
    if (idx.own >= 0) {
      let rawOwn = parseFloat(row[idx.own]) || 0;
      // Normalize: if entered as decimal (0.15) convert to percent (15)
      if (rawOwn > 0 && rawOwn <= 1) rawOwn *= 100;
      ownPct = rawOwn;
    }

    drivers.push({
      name:       rawName,
      key:        resolveKey(rawName, nameMap),
      salary:     salary,
      dkId:       idx.dkId >= 0 ? (row[idx.dkId] || "") : "",
      dkAvgFPPG:  idx.fppg >= 0 ? (parseFloat(row[idx.fppg]) || 0) : 0,
      ownPct:     ownPct
    });
  }

  Logger.log("DK_Salaries loaded: " + drivers.length + " drivers.");
  return drivers;
}


/* -------------------------------------------------------
 *  5. Playability Loader (iFantasyRace)
 *
 *  Primary projection source. Replaces the LineStar
 *  projection fields entirely.
 *
 *  Expected iFantasyRace Playability columns:
 *    Name | Start | DK Price | Finish Projection |
 *    DK Scoring Projection | Dominator | Fantasy Rank
 *
 *  Range fields ("55 - 90") are parsed to low/mid/high.
 *  Dominator rating text is converted to a 1-4 numeric score.
 *
 *  Returns a map: canonicalKey → {
 *    startPos, ifrProjLow, ifrProjMid, ifrProjHigh,
 *    finProjLow, finProjHigh, domScore, domLabel, ifrRank
 *  }
 * ------------------------------------------------------- */

function loadPlayability(ss, nameMap) {
  const sheet = ss.getSheetByName("Playability");
  if (!sheet) {
    Logger.log("WARNING: Playability sheet not found — all drivers will use fallback projections.");
    return {};
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return {};

  const h = data[0].map(v => v ? v.toString().toLowerCase().trim() : "");

  function findCol(target) {
    return h.findIndex(col => col.indexOf(target) >= 0);
  }

  const idx = {
    name:     h.findIndex(col => col.includes("name") || col.includes("driver")),
    start:    findCol("start"),
    finProj:  findCol("finish projection"),
    dkProj:   findCol("dk scoring projection"),
    dom:      findCol("dominator"),
    rank:     findCol("rank")
  };

  if (idx.name < 0) {
    Logger.log("ERROR: Playability sheet missing Name/Driver column. Headers: " + h.join(", "));
    return {};
  }

  const map = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rawName = row[idx.name] ? row[idx.name].toString().trim() : "";
    if (!rawName) continue;

    const key = resolveKey(rawName, nameMap);
    if (!key) continue;

    // --- Parse range fields ("55 - 90" → low=55, high=90, mid=72.5) ---
    const dkProjRange = idx.dkProj >= 0 ? parseRange(row[idx.dkProj]) : { low: 0, high: 0, mid: 0 };
    const finProjRange = idx.finProj >= 0 ? parseRange(row[idx.finProj]) : { low: 0, high: 0, mid: 0 };

    // --- Parse start position ---
    const startPos = idx.start >= 0 ? (parseFloat(row[idx.start]) || 20) : 20;

    // --- Parse dominator rating ---
    const domRaw   = idx.dom >= 0 ? row[idx.dom].toString().trim() : "";
    const domScore  = parseDomRating(domRaw);

    // --- Parse iFantasyRace rank ---
    const ifrRank = idx.rank >= 0 ? (parseInt(row[idx.rank]) || 0) : 0;

    map[key] = {
      startPos:    startPos,
      ifrProjLow:  dkProjRange.low,
      ifrProjMid:  dkProjRange.mid,
      ifrProjHigh: dkProjRange.high,
      finProjLow:  finProjRange.low,
      finProjHigh: finProjRange.high,
      domScore:    domScore,
      domLabel:    domRaw,
      ifrRank:     ifrRank
    };
  }

  Logger.log("Playability loaded: " + Object.keys(map).length + " drivers.");
  return map;
}


/* -------------------------------------------------------
 *  Utility: parseRange
 *
 *  Parses iFantasyRace range strings into low/mid/high.
 *  Handles: "55 - 90", "55-90", "55–90", or plain "72"
 *  Returns { low, mid, high } — all zero if unparseable.
 * ------------------------------------------------------- */

function parseRange(raw) {
  if (!raw && raw !== 0) return { low: 0, mid: 0, high: 0 };
  const str = raw.toString().trim();

  // Try to split on common separators: " - ", "-", "–", " to "
  const parts = str.split(/\s*[-–to]+\s*/).map(p => parseFloat(p.trim()));

  if (parts.length >= 2 && !isNaN(parts[0]) && !isNaN(parts[1])) {
    const low  = Math.min(parts[0], parts[1]);
    const high = Math.max(parts[0], parts[1]);
    const mid  = Math.round(((low + high) / 2) * 100) / 100;
    return { low, mid, high };
  }

  // Single value — treat as midpoint with ±20% spread
  const val = parseFloat(str);
  if (!isNaN(val) && val > 0) {
    return {
      low:  Math.round(val * 0.80 * 100) / 100,
      mid:  val,
      high: Math.round(val * 1.20 * 100) / 100
    };
  }

  return { low: 0, mid: 0, high: 0 };
}


/* -------------------------------------------------------
 *  Utility: parseDomRating
 *
 *  Converts iFantasyRace dominator text labels to numeric
 *  scores used in Notes and wizard sorting.
 *
 *  High = 4, Medium-High = 3, Medium = 2, Low = 1, blank = 0
 * ------------------------------------------------------- */

function parseDomRating(raw) {
  if (!raw) return 0;
  const s = raw.toString().toLowerCase().trim();
  if (s === "high")                                   return 4;
  if (s === "medium-high" || s === "medium high")     return 3;
  if (s === "medium" || s === "med")                  return 2;
  if (s === "low")                                    return 1;
  return 0;
}
/* -------------------------------------------------------
 *  6. Loop Data Loader (Historical Race Stats at Track)
 *
 *  Reads the most recent single race at this track.
 *  Exact column headers expected:
 *    Driver, Avg Position, # Fastest Laps,
 *    % Laps in Top 15, % Laps Led, Driver Rating
 *
 *  pctLapsLed is a percentage (0-100). Analysis.gs converts
 *  to projected lap count: (pctLapsLed / 100) * raceContext.laps
 *
 *  Returns a map: canonicalKey → {
 *    pctLapsLed, fastLaps, top15Pct, rating, avgPos
 *  }
 * ------------------------------------------------------- */

function loadLoopData(ss, nameMap) {
  const sheet = ss.getSheetByName("Data_Loop");
  if (!sheet) { Logger.log("WARNING: Data_Loop sheet not found."); return {}; }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return {};

  const h  = data[0].map(v => v ? v.toString().trim() : "");
  const hl = h.map(v => v.toLowerCase());

  const idx = {
    driver:     hl.findIndex(col => col === "driver"),
    pctLapsLed: hl.findIndex(col => col === "% laps led"),
    fastLaps:   hl.findIndex(col => col === "# fastest laps"),
    top15:      hl.findIndex(col => col === "% laps in top 15"),
    rating:     hl.findIndex(col => col === "driver rating"),
    avgPos:     hl.findIndex(col => col === "avg position")
  };

  if (idx.driver < 0) {
    Logger.log("ERROR: Data_Loop missing Driver column. Headers: " + h.join(", "));
    return {};
  }

  const map = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const key = resolveKey(row[idx.driver], nameMap);
    if (!key) continue;

    map[key] = {
      pctLapsLed: idx.pctLapsLed >= 0 ? (parseFloat(row[idx.pctLapsLed]) || 0) : 0,
      fastLaps:   idx.fastLaps   >= 0 ? (parseFloat(row[idx.fastLaps])   || 0) : 0,
      top15Pct:   idx.top15      >= 0 ? (parseFloat(row[idx.top15])      || 0) : 0,
      rating:     idx.rating     >= 0 ? (parseFloat(row[idx.rating])     || 0) : 0,
      avgPos:     idx.avgPos     >= 0 ? (parseFloat(row[idx.avgPos])     || 25) : 25
    };
  }

  Logger.log("Data_Loop loaded: " + Object.keys(map).length + " drivers.");
  return map;
}


/* -------------------------------------------------------
 *  7. Avg Finish Loader (Multi-Season, Recency-Weighted)
 *
 *  Columns: Rank, Driver, 2022 #1, 2022 #2, ..., Avg. Finish
 *  Returns a map: canonicalKey → { weightedAvg, rawAvg }
 * ------------------------------------------------------- */

function loadAvgFinish(ss, nameMap) {
  const sheet = ss.getSheetByName("Data_Avg_Finish");
  if (!sheet) { Logger.log("WARNING: Data_Avg_Finish not found."); return {}; }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const driverIdx = headers.findIndex(h => h && h.toString().toLowerCase().includes("driver"));
  const avgIdx    = headers.findIndex(h => h && h.toString().toLowerCase().includes("avg"));

  const map = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const key = resolveKey(row[driverIdx >= 0 ? driverIdx : 1], nameMap);
    if (!key) continue;

    let weightedSum = 0, weightTotal = 0, rawSum = 0, rawCount = 0;

    for (let c = 0; c < row.length; c++) {
      if (c === driverIdx || c === avgIdx) continue;
      const val = parseFloat(row[c]);
      if (!val || val <= 0) continue;
      const colHeader = headers[c] ? headers[c].toString() : "";
      if (colHeader.toLowerCase().indexOf("avg") >= 0) continue;
      const w = getRecencyWeight(colHeader);
      weightedSum += val * w;
      weightTotal += w;
      rawSum += val;
      rawCount++;
    }

    map[key] = {
      weightedAvg: weightTotal > 0 ? weightedSum / weightTotal : 25,
      rawAvg:      rawCount > 0 ? rawSum / rawCount : 25
    };
  }

  return map;
}


/* -------------------------------------------------------
 *  7b. Avg Start Loader (All-Time Track Stats)
 *
 *  Sourced from DriverAverages.com for the current track.
 *  No recency weighting — all-time stats at this venue.
 *
 *  Column structure (0-indexed):
 *    0: Rank  1: Driver  2: Avg Finish  3: Races
 *    4: Wins  5: Top 5s  6: Top 10s  7: Top 20s
 *    8: Laps Led (total)  9: Avg Start  10: Best Finish
 *    11: Low Finish  12: DNF  13: Avg Rating  14: detail
 *
 *  Returns a map: canonicalKey → {
 *    avgStart, avgStartFinishDiff, siteLapsLed, siteRaces, siteAvgRating
 *  }
 * ------------------------------------------------------- */

function loadAvgStart(ss, nameMap) {
  const sheet = ss.getSheetByName("Data_Avg_Start");
  if (!sheet) { Logger.log("WARNING: Data_Avg_Start not found."); return {}; }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return {};

  const map = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const key = resolveKey(row[1], nameMap);
    if (!key) continue;

    const avgFinish = parseFloat(row[2])  || 0;
    const races     = parseFloat(row[3])  || 0;
    const lapsLed   = parseFloat(row[8])  || 0;
    const avgStart  = parseFloat(row[9])  || 0;
    const avgRating = parseFloat(row[13]) || 0;

    if (!avgStart && !races) continue;

    map[key] = {
      avgStart:            avgStart,
      avgStartFinishDiff:  (avgStart > 0 && avgFinish > 0) ? avgStart - avgFinish : 0,
      siteLapsLed:         lapsLed,
      siteRaces:           races,
      siteAvgRating:       avgRating
    };
  }

  Logger.log("Data_Avg_Start loaded: " + Object.keys(map).length + " drivers.");
  return map;
}


/* -------------------------------------------------------
 *  8. Ratings Loader (Multi-Season, Recency-Weighted)
 *
 *  Values are RANK positions (1=best). Lower = better.
 *  Returns a map: canonicalKey → { weightedAvg }
 * ------------------------------------------------------- */

function loadRatings(ss, nameMap) {
  const sheet = ss.getSheetByName("Data_Ratings");
  if (!sheet) { Logger.log("WARNING: Data_Ratings not found."); return {}; }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const driverIdx = headers.findIndex(h => h && h.toString().toLowerCase().includes("driver"));
  const avgIdx    = headers.findIndex(h => {
    const v = h ? h.toString().toLowerCase() : "";
    return v.includes("average") || v.includes("avg");
  });

  const map = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const key = resolveKey(row[driverIdx >= 0 ? driverIdx : 1], nameMap);
    if (!key) continue;

    let weightedSum = 0, weightTotal = 0;

    for (let c = 0; c < row.length; c++) {
      if (c === driverIdx || c === avgIdx) continue;
      const val = parseFloat(row[c]);
      if (!val || val <= 0) continue;
      const colHeader = headers[c] ? headers[c].toString() : "";
      // Skip any summary column that slipped through
      const ch = colHeader.toLowerCase();
      if (ch.includes("average") || ch.includes("avg")) continue;
      const w = getRecencyWeight(colHeader);
      weightedSum += val * w;
      weightTotal += w;
    }

    map[key] = {
      weightedAvg: weightTotal > 0 ? weightedSum / weightTotal : 20
    };
  }

  return map;
}


/* -------------------------------------------------------
 *  9. Practice Data Loader
 *
 *  bestTime = practice speed in MPH (higher = faster).
 *  computePracticeRanks() sorts descending accordingly.
 *  Returns a map: canonicalKey → { bestTime }
 * ------------------------------------------------------- */

function loadPractice(ss, nameMap) {
  const sheet = ss.getSheetByName("practice_1");
  if (!sheet) { Logger.log("WARNING: practice_1 not found."); return {}; }

  const data = sheet.getDataRange().getValues();
  const h = data[0];
  const nameIdx  = h.findIndex(col => col && col.toString().toLowerCase().includes("driver"));
  const speedIdx = h.findIndex(col => col && col.toString().toLowerCase().includes("speed"));

  const map = {};
  for (let i = 1; i < data.length; i++) {
    const key = resolveKey(data[i][nameIdx >= 0 ? nameIdx : 0], nameMap);
    if (!key) continue;
    map[key] = {
      bestTime: parseFloat(data[i][speedIdx >= 0 ? speedIdx : 6]) || 0
    };
  }

  return map;
}


/* -------------------------------------------------------
 *  10. Qualifying Data Loader
 *
 *  Returns a map: canonicalKey → { speed, car }
 *  car = manufacturer (Chevrolet, Ford, Toyota)
 * ------------------------------------------------------- */

function loadQualifying(ss, nameMap) {
  const sheet = ss.getSheetByName("qualifying_1");
  if (!sheet) { Logger.log("WARNING: qualifying_1 not found."); return {}; }

  const data = sheet.getDataRange().getValues();
  const h = data[0];
  const driverIdx = h.findIndex(col => col && col.toString().toLowerCase().includes("driver"));
  const carIdx    = h.findIndex(col => col && col.toString().toLowerCase().includes("car"));
  const speedIdx  = h.findIndex(col => col && col.toString().toLowerCase().includes("speed"));

  const map = {};
  for (let i = 1; i < data.length; i++) {
    const key = resolveKey(data[i][driverIdx >= 0 ? driverIdx : 1], nameMap);
    if (!key) continue;
    map[key] = {
      speed: parseFloat(data[i][speedIdx >= 0 ? speedIdx : 5]) || 0,
      car:   data[i][carIdx >= 0 ? carIdx : 3] || ""
    };
  }

  return map;
}


/* -------------------------------------------------------
 *  11. Practice Speed Rank Calculator
 *
 *  Ranks by practice speed descending (higher MPH = rank 1).
 *  Drivers without data get 60th-percentile default rank.
 * ------------------------------------------------------- */

function computePracticeRanks(drivers) {
  const withSpeed = drivers
    .filter(d => d.pracBestTime > 0)
    .sort((a, b) => b.pracBestTime - a.pracBestTime);

  for (let i = 0; i < withSpeed.length; i++) {
    withSpeed[i].pracSpeedRank = i + 1;
  }

  const defaultRank = Math.round(drivers.length * 0.6);
  for (const d of drivers) {
    if (d.pracSpeedRank === 0) d.pracSpeedRank = defaultRank;
  }
}
/* -------------------------------------------------------
 *  12. Green Flag Speed Loader
 *
 *  Reads the most recent single race at this track.
 *  Exact column header expected: Green Flag Speed
 *
 *  Returns a map: canonicalKey → { greenFlagSpeed }
 * ------------------------------------------------------- */

function loadGreenSpeed(ss, nameMap) {
  const sheet = ss.getSheetByName("Data_Green_Speed");
  if (!sheet) { Logger.log("WARNING: Data_Green_Speed not found."); return {}; }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return {};

  const h  = data[0].map(v => v ? v.toString().trim() : "");
  const hl = h.map(v => v.toLowerCase());

  const driverIdx = hl.findIndex(col => col === "driver");
  const speedIdx  = hl.findIndex(col => col === "green flag speed");

  if (driverIdx < 0 || speedIdx < 0) {
    Logger.log("ERROR: Data_Green_Speed missing Driver or Green Flag Speed column. Headers: " + h.join(", "));
    return {};
  }

  const map = {};
  for (let i = 1; i < data.length; i++) {
    const key = resolveKey(data[i][driverIdx], nameMap);
    if (!key) continue;
    const spd = parseFloat(data[i][speedIdx]);
    if (!spd || spd <= 0) continue;
    map[key] = { greenFlagSpeed: spd };
  }

  Logger.log("Data_Green_Speed loaded: " + Object.keys(map).length + " drivers.");
  return map;
}

/* -------------------------------------------------------
 *  13. Total Speed Loader (Segment Consistency)
 *
 *  Reads the most recent single race at this track.
 *  Exact column headers expected:
 *    Driver, Avg. Speed Rank, #1 (...), #2 (...), etc.
 *
 *  Segment columns (starting with "#") are used to compute
 *  consistency — inverse of rank standard deviation across
 *  segments. A driver ranked 3rd in all segments is more
 *  reliable than one alternating 1st and 25th.
 *
 *  Returns a map: canonicalKey → { avgSpeedRank, consistency }
 * ------------------------------------------------------- */

function loadTotalSpeed(ss, nameMap) {
  const sheet = ss.getSheetByName("Data_Total_Speed");
  if (!sheet) { Logger.log("WARNING: Data_Total_Speed not found."); return {}; }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return {};

  const h  = data[0].map(v => v ? v.toString().trim() : "");
  const hl = h.map(v => v.toLowerCase());

  const driverIdx  = hl.findIndex(col => col === "driver");
  const avgRankIdx = hl.findIndex(col => col === "avg. speed rank");

  // Segment columns start with "#"
  const segIdxs = [];
  for (let c = 0; c < h.length; c++) {
    if (h[c] && h[c].toString().indexOf("#") === 0) segIdxs.push(c);
  }

  if (driverIdx < 0) {
    Logger.log("ERROR: Data_Total_Speed missing Driver column. Headers: " + h.join(", "));
    return {};
  }

  const entries = [];
  for (let i = 1; i < data.length; i++) {
    const key = resolveKey(data[i][driverIdx], nameMap);
    if (!key) continue;

    const avgRank = avgRankIdx >= 0 ? (parseFloat(data[i][avgRankIdx]) || 0) : 0;
    if (avgRank <= 0) continue;

    const segs = segIdxs.map(c => parseFloat(data[i][c])).filter(v => v && v > 0);

    let stdev = 0;
    if (segs.length >= 2) {
      const mean = segs.reduce((a, b) => a + b, 0) / segs.length;
      stdev = Math.sqrt(segs.reduce((a, v) => a + Math.pow(v - mean, 2), 0) / segs.length);
    }

    entries.push({ key, avgRank, stdev });
  }

  const stdevs   = entries.map(e => e.stdev);
  const maxStdev = stdevs.length > 0 ? Math.max(...stdevs) : 1;
  const minStdev = stdevs.length > 0 ? Math.min(...stdevs) : 0;

  const map = {};
  for (const e of entries) {
    const consistency = maxStdev > minStdev
      ? ((maxStdev - e.stdev) / (maxStdev - minStdev)) * 100
      : 50;

    map[e.key] = {
      avgSpeedRank: e.avgRank,
      consistency:  Math.round(consistency * 10) / 10
    };
  }

  Logger.log("Data_Total_Speed loaded: " + Object.keys(map).length + " drivers.");
  return map;
}


/* -------------------------------------------------------
 *  14. Green Flag Speed Rank Calculator
 *
 *  Ranks by green flag speed descending (higher MPH = rank 1).
 *  Drivers without data get 60th-percentile default rank.
 * ------------------------------------------------------- */

function computeGreenSpeedRanks(drivers) {
  const withSpeed = drivers
    .filter(d => d.histGreenSpeed > 0)
    .sort((a, b) => b.histGreenSpeed - a.histGreenSpeed);

  for (let i = 0; i < withSpeed.length; i++) {
    withSpeed[i].histGreenRank = i + 1;
  }

  const defaultRank = Math.round(drivers.length * 0.6);
  for (const d of drivers) {
    if (d.histGreenRank === 0) d.histGreenRank = defaultRank;
  }
}


/* -------------------------------------------------------
 *  15. Composite Speed Score Calculator
 *
 *  Blends practice, green flag speed, and segment
 *  consistency into a single composite rank.
 *
 *  WITH practice data (>30% of drivers have times):
 *    Practice 50% | Green Speed 30% | Consistency 20%
 *
 *  WITHOUT practice data:
 *    Green Speed 55% | Consistency 45%
 *
 *  Output: speedComposite — rank where 1 = fastest.
 * ------------------------------------------------------- */

function computeSpeedComposite(drivers) {
  const withPractice = drivers.filter(d => d.pracBestTime > 0).length;
  const hasPractice  = withPractice > (drivers.length * 0.3);
  const fieldSize    = drivers.length;

  for (const d of drivers) {
    const pracScore  = fieldSize - d.pracSpeedRank + 1;
    const greenScore = fieldSize - d.histGreenRank + 1;
    const consScore  = d.speedConsistency;  // already 0-100

    d._speedRawScore = hasPractice
      ? (pracScore * 0.50) + (greenScore * 0.30) + (consScore * 0.20)
      : (greenScore * 0.55) + (consScore * 0.45);
  }

  const sorted = drivers.slice().sort((a, b) => b._speedRawScore - a._speedRawScore);
  for (let i = 0; i < sorted.length; i++) {
    sorted[i].speedComposite = i + 1;
  }

  for (const d of drivers) delete d._speedRawScore;
}