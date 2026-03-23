/**
 * ============================================================
 *  NASCAR DFS 2026 — Module 1: CONFIG
 * ============================================================
 */


/* -------------------------------------------------------
 *  1. DraftKings Scoring Constants
 * ------------------------------------------------------- */

const DK = {
  LAPS_LED_PTS:    0.25,
  FASTEST_LAP_PTS: 0.45,
  PASS_DIFF_PTS:   0.25,
  WIN_BONUS:       3,
  FINISH_BASE:     43
};


/* -------------------------------------------------------
 *  2. Weight Profiles by Track Type
 * ------------------------------------------------------- */

function getWeights(trackType) {
  const profiles = {
    "Superspeedway":            { dom: 0.00, pd: 0.70, speed: 0.05, skill: 0.05, history: 0.20 },
    "Superspeedway (Drafting)": { dom: 0.00, pd: 0.70, speed: 0.05, skill: 0.05, history: 0.20 },
    "Intermediate":             { dom: 0.25, pd: 0.20, speed: 0.25, skill: 0.15, history: 0.15 },
    "Short Track (Flat)":       { dom: 0.35, pd: 0.10, speed: 0.25, skill: 0.15, history: 0.15 },
    "Short Track (Steep)":      { dom: 0.35, pd: 0.10, speed: 0.20, skill: 0.20, history: 0.15 },
    "Short Track (Fast)":       { dom: 0.25, pd: 0.15, speed: 0.30, skill: 0.15, history: 0.15 },
    "Short Track":              { dom: 0.30, pd: 0.15, speed: 0.25, skill: 0.15, history: 0.15 },
    "Short Track (Wear)":       { dom: 0.25, pd: 0.10, speed: 0.15, skill: 0.20, history: 0.30 },
    "Road Course":              { dom: 0.05, pd: 0.25, speed: 0.15, skill: 0.15, history: 0.40 },
    "Street Course":            { dom: 0.00, pd: 0.30, speed: 0.10, skill: 0.15, history: 0.45 },
    "Large Oval":               { dom: 0.20, pd: 0.20, speed: 0.30, skill: 0.15, history: 0.15 },
    "Large Triangle":           { dom: 0.10, pd: 0.25, speed: 0.25, skill: 0.15, history: 0.25 },
    "1-Mile Flat":              { dom: 0.25, pd: 0.15, speed: 0.30, skill: 0.15, history: 0.15 }
  };
  const w = profiles[trackType];
  if (!w) {
    Logger.log("CONFIG WARNING: Unknown track type '" + trackType + "' — using balanced defaults.");
    return { dom: 0.20, pd: 0.20, speed: 0.20, skill: 0.20, history: 0.20 };
  }
  return w;
}


/* -------------------------------------------------------
 *  3. Estimated Caution Rates by Track Type
 * ------------------------------------------------------- */

function getCautionRate(trackType) {
  const rates = {
    "Superspeedway":            0.30,
    "Superspeedway (Drafting)": 0.30,
    "Intermediate":             0.22,
    "Short Track (Flat)":       0.17,
    "Short Track (Steep)":      0.20,
    "Short Track (Fast)":       0.18,
    "Short Track":              0.18,
    "Short Track (Wear)":       0.20,
    "Road Course":              0.15,
    "Street Course":            0.18,
    "Large Oval":               0.22,
    "Large Triangle":           0.20,
    "1-Mile Flat":              0.18
  };
  return rates[trackType] || 0.20;
}


/* -------------------------------------------------------
 *  4. Target Dominator Counts by Track Type
 *  Returns { min, max }
 * ------------------------------------------------------- */

function getTargetDominators(trackType) {
  const targets = {
    "Superspeedway":            { min: 0, max: 0 },
    "Superspeedway (Drafting)": { min: 0, max: 0 },
    "Intermediate":             { min: 2, max: 3 },
    "Short Track (Flat)":       { min: 2, max: 3 },
    "Short Track (Steep)":      { min: 2, max: 3 },
    "Short Track (Fast)":       { min: 2, max: 3 },
    "Short Track":              { min: 2, max: 3 },
    "Short Track (Wear)":       { min: 2, max: 3 },
    "Road Course":              { min: 0, max: 1 },
    "Street Course":            { min: 0, max: 0 },
    "Large Oval":               { min: 2, max: 3 },
    "Large Triangle":           { min: 1, max: 2 },
    "1-Mile Flat":              { min: 2, max: 3 }
  };
  const t = targets[trackType];
  if (!t) {
    Logger.log("CONFIG WARNING: Unknown track type '" + trackType + "' for target doms.");
    return { min: 1, max: 2 };
  }
  return t;
}


/* -------------------------------------------------------
 *  4b. Target PD Value Play Counts by Track Type
 *  Returns { min, max }
 * ------------------------------------------------------- */

function getTargetPD(trackType) {
  const targets = {
    "Superspeedway":            { min: 3, max: 4 },
    "Superspeedway (Drafting)": { min: 3, max: 4 },
    "Intermediate":             { min: 1, max: 2 },
    "Short Track (Flat)":       { min: 1, max: 1 },
    "Short Track (Steep)":      { min: 1, max: 1 },
    "Short Track (Fast)":       { min: 1, max: 2 },
    "Short Track":              { min: 1, max: 1 },
    "Short Track (Wear)":       { min: 1, max: 1 },
    "Road Course":              { min: 1, max: 2 },
    "Street Course":            { min: 2, max: 3 },
    "Large Oval":               { min: 1, max: 2 },
    "Large Triangle":           { min: 1, max: 2 },
    "1-Mile Flat":              { min: 1, max: 2 }
  };
  const t = targets[trackType];
  if (!t) {
    Logger.log("CONFIG WARNING: Unknown track type '" + trackType + "' for target PD.");
    return { min: 1, max: 2 };
  }
  return t;
}


/* -------------------------------------------------------
 *  5. Dom Points Available Calculator
 * ------------------------------------------------------- */

function calcDomPointsAvailable(totalLaps, trackType) {
  const cautionRate = getCautionRate(trackType);
  const greenLaps   = Math.round(totalLaps * (1 - cautionRate));
  const raw         = totalLaps * (DK.LAPS_LED_PTS + DK.FASTEST_LAP_PTS);
  const adjusted    = greenLaps * (DK.LAPS_LED_PTS + DK.FASTEST_LAP_PTS);
  return {
    raw:         Math.round(raw * 10) / 10,
    adjusted:    Math.round(adjusted * 10) / 10,
    greenLaps:   greenLaps,
    cautionRate: cautionRate
  };
}


/* -------------------------------------------------------
 *  6. Group Assignment Thresholds
 *
 *  Active groups: DOM, PD, LEVERAGE (ownership-dependent),
 *  UNDER (ownership-dependent).
 *  CORE has been dissolved — those drivers fall into Fill.
 * ------------------------------------------------------- */

const GROUP_THRESHOLDS = {
  DOM_MAX_START_POS:  15,
  PD_MIN_START_POS:   20,
  PD_MIN_PROJ_PD:     8,
  LEVERAGE_MAX_OWN:   15,
  LEVERAGE_MIN_EDGE:  0,
  UNDER_PERCENTILE:   25
};


/* -------------------------------------------------------
 *  7. Exposure Defaults
 * ------------------------------------------------------- */

function calcMaxExposure(ownPct, group, trackType) {
  const isSS = trackType.indexOf("Superspeedway") >= 0;
  let maxExp;
  if (ownPct > 40)      maxExp = ownPct - 10;
  else if (ownPct > 20) maxExp = ownPct + 5;
  else if (ownPct > 10) maxExp = ownPct + 10;
  else                  maxExp = 25;
  if (isSS && maxExp > 40) maxExp = 40;
  return Math.round(Math.min(maxExp, 100));
}

function calcMinExposure(group) {
  switch (group) {
    case "DOM":      return 15;
    case "LEVERAGE": return 10;
    case "PD":       return 5;
    default:         return 0;
  }
}


/* -------------------------------------------------------
 *  8. Cash Game Scoring Weights
 * ------------------------------------------------------- */

const CASH_WEIGHTS = {
  floorW:     0.40,
  projW:      0.25,
  pdW:        0.10,
  stdPenalty: 0.15,
  chalkW:     0.10,
  valueW:     0.15
};

const CASH_ROSTER_SIZE = 6;
const CASH_SALARY_CAP  = 50000;


/* -------------------------------------------------------
 *  9. Recency Weights for Historical Averages
 * ------------------------------------------------------- */

const RECENCY_WEIGHTS = {
  "2025": 1.00,
  "2024": 0.75,
  "2023": 0.50,
  "2022": 0.30
};

function getRecencyWeight(columnHeader) {
  for (const year in RECENCY_WEIGHTS) {
    if (columnHeader && columnHeader.toString().indexOf(year) >= 0) {
      return RECENCY_WEIGHTS[year];
    }
  }
  return 0.25;
}


/* -------------------------------------------------------
 *  10. Shared Utility Functions
 * ------------------------------------------------------- */

function cleanName(n) {
  if (!n) return "";
  return n.toString()
    .toLowerCase()
    .replace(/no\.\s*\d+/g, "")
    .replace(/chevrolet|ford|toyota/gi, "")
    .replace(/[^a-z]/g, "")
    .trim();
}

function normalize(value, arr) {
  const max = Math.max(...arr);
  const min = Math.min(...arr);
  if (max === min) return 50;
  return ((value - min) / (max - min)) * 100;
}

function percentile(arr, p) {
  const sorted = arr.slice().sort((a, b) => a - b);
  const idx    = (p / 100) * (sorted.length - 1);
  const lower  = Math.floor(idx);
  const frac   = idx - lower;
  if (lower + 1 >= sorted.length) return sorted[lower];
  return sorted[lower] + frac * (sorted[lower + 1] - sorted[lower]);
}

function calcDomPoints(projLapsLed, projFastestLaps) {
  return (projLapsLed * DK.LAPS_LED_PTS) + (projFastestLaps * DK.FASTEST_LAP_PTS);
}


/* -------------------------------------------------------
 *  11. Dashboard Column Layout
 *
 *  Active groups: DOM, PD, LEVERAGE, UNDER.
 *  CORE dissolved — falls into Fill/ungrouped.
 * ------------------------------------------------------- */

const DASH_COLS = {
  GPP_HEADER_ROW: 1,
  GPP_DATA_START: 2,
  GPP_HEADERS: [
    "✓",
    "Driver", "Sal", "Start", "Own%", "Group",
    "Proj", "Adj Proj", "Floor", "Ceil", "Std",
    "Dom Pts", "Dom Rank", "PD Proj",
    "Edge", "Value",
    "Cash Score", "Track Hist",
    "Avg S/F Diff", "Notes"
  ],

  COL_CHECK:    1,
  COL_DRIVER:   2,
  COL_SALARY:   3,
  COL_START:    4,
  COL_OWN:      5,
  COL_GROUP:    6,
  COL_PROJ:     7,
  COL_ADJPROJ:  8,
  COL_FLOOR:    9,
  COL_CEIL:     10,
  COL_STD:      11,
  COL_DOMPTS:   12,
  COL_DOMRANK:  13,
  COL_PD:       14,
  COL_EDGE:     15,
  COL_VALUE:    16,
  COL_CASHSCORE: 17,
  COL_TRACKHIST: 18,
  COL_AVGDIFF:  19,
  COL_NOTES:    20,
  TOTAL_COLS:   20,

  CASH_HEADER_ROW: 5,
  CASH_DATA_START: 6,
  CASH_HEADERS: [
    "Slot", "Driver", "Salary", "Floor", "Proj", "Own%", "Cash Score"
  ]
};


/* -------------------------------------------------------
 *  12. Adj Projection: Maximum Adjustment Bounds
 * ------------------------------------------------------- */

const ADJ_PROJ_BOUNDS = {
  MAX_DOM_ADJ:     8,
  MAX_PD_ADJ:      6,
  MAX_SPEED_ADJ:   5,
  MAX_HISTORY_ADJ: 5
};