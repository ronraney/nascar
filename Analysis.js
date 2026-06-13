/**
 * ============================================================
 *  NASCAR DFS 2026 — Module 3: ANALYSIS
 * ============================================================
 *  Active groups: DOM, PD, LEVERAGE (ownership-dependent),
 *  UNDER (ownership-dependent).
 *  CORE dissolved — those drivers fall into Fill/ungrouped.
 * ============================================================
 */


/* -------------------------------------------------------
 *  1. Main Analysis Runner
 * ------------------------------------------------------- */

function runAnalysis(data) {
  const { drivers, raceContext } = data;
  const trackType = raceContext.trackType;

  computeRecentForm(drivers);

  computeDomPoints(drivers, raceContext);
  applyTTDomBlend(drivers);

  computePDProjection(drivers);
  applyTTPDBlend(drivers);

  computeAdjProjection(drivers, raceContext);
  applyTTAdjBlend(drivers);

  computeEdge(drivers);

  for (const d of drivers) {
    d.value = d.salary > 0
      ? Math.round((d.adjProj / (d.salary / 1000)) * 100) / 100
      : 0;
  }

  assignGroups(drivers, raceContext);

  for (const d of drivers) {
    d.maxExp = calcMaxExposure(d.ownPct, d.group, trackType);
    d.minExp = calcMinExposure(d.group);
  }

  computeCashScores(drivers);

  // ---- Step 9: Track History Score ----
  computeTrackHistScore(drivers);
  computeTTHistScore(drivers);

  // ---- Step 10: Cash Core Grade + Group Assignment ----
  // Must run after track history score and cash score are computed.
  // Reassigns top 15% of non-DOM/PD drivers to CASHCORE group.
  assignCashCoreGroup(drivers);

  return drivers;
}


/* -------------------------------------------------------
 *  1b. Recent Form Score (Current Season Standings)
 *
 *  Composite of five signals, all normalized 0-100:
 *    Win rate      (0.30)
 *    Top-5 rate    (0.25)
 *    Top-10 rate   (0.15)
 *    Laps led/race (0.20)
 *    Avg finish    (0.10, inverted — lower finish pos = better)
 *
 *  Result stored as recentFormScore (0-100, higher = better form).
 *  Drivers with no standings data receive the field median.
 * ------------------------------------------------------- */

function computeRecentForm(drivers) {
  const withData = drivers.filter(d => d.recentRaces > 0);
  if (withData.length === 0) return;

  const winPctArr   = withData.map(d => d.recentWins    / d.recentRaces);
  const top5PctArr  = withData.map(d => d.recentTop5    / d.recentRaces);
  const top10PctArr = withData.map(d => d.recentTop10   / d.recentRaces);
  const lledRaceArr = withData.map(d => d.recentLapsLed / d.recentRaces);
  const finInvArr   = withData.map(d => -d.recentAvgFinish);

  for (const d of withData) {
    const winNorm   = normalize(d.recentWins    / d.recentRaces, winPctArr);
    const top5Norm  = normalize(d.recentTop5    / d.recentRaces, top5PctArr);
    const top10Norm = normalize(d.recentTop10   / d.recentRaces, top10PctArr);
    const lledNorm  = normalize(d.recentLapsLed / d.recentRaces, lledRaceArr);
    const finNorm   = normalize(-d.recentAvgFinish,              finInvArr);

    d.recentFormScore = Math.round(
      (winNorm   * 0.30)
    + (top5Norm  * 0.25)
    + (top10Norm * 0.15)
    + (lledNorm  * 0.20)
    + (finNorm   * 0.10)
    );
  }

  const scores      = withData.map(d => d.recentFormScore);
  const medianScore = Math.round(percentile(scores, 50));
  for (const d of drivers) {
    if (d.recentRaces === 0) d.recentFormScore = medianScore;
  }
}


/* -------------------------------------------------------
 *  2. Dominator Points Calculation
 *
 *  Six normalized signals weighted and combined.
 *  Reliability-adjusted by site race count.
 *  Speed signal priority: qualSpeed → pracBestTime → 50 (neutral).
 * ------------------------------------------------------- */

function computeDomPoints(drivers, raceContext) {
  const hasQualSpeed  = drivers.some(d => d.qualSpeed    > 0);
  const hasPracSpeed  = drivers.some(d => d.pracBestTime > 0);

  const speedSignalArr = hasQualSpeed
    ? drivers.map(d => d.qualSpeed)
    : hasPracSpeed
      ? drivers.map(d => d.pracBestTime)
      : null;

  const histPctArr    = drivers.map(d => d.histPctLapsLed);
  const projArr       = drivers.map(d => d.proj);
  const histRatingArr = drivers.map(d => d.histRating);
  const siteLapsArr   = drivers.map(d => d.siteLapsLed);
  const recentLLArr   = drivers.map(d => d.recentRaces > 0 ? d.recentLapsLed / d.recentRaces : 0);

  // startPos penalty: logarithmic — P1 gets highest raw penalty value (best)
  // rawPenalty = 1 / log(startPos + e - 1), then normalize across field
  const startPenaltyArr = drivers.map(d => 1.0 / Math.log(d.startPos + Math.E - 1));

  for (const d of drivers) {
    const speedVal         = hasQualSpeed ? d.qualSpeed : d.pracBestTime;
    const histPctNorm      = normalize(d.histPctLapsLed, histPctArr);
    const startPenaltyNorm = normalize(1.0 / Math.log(d.startPos + Math.E - 1), startPenaltyArr);
    const qualSpeedNorm    = speedSignalArr ? normalize(speedVal, speedSignalArr) : 50;
    const projNorm         = normalize(d.proj,           projArr);
    const histRatingNorm   = normalize(d.histRating,     histRatingArr);
    const siteLapsNorm     = normalize(d.siteLapsLed,    siteLapsArr);
    const recentLLPerRace  = d.recentRaces > 0 ? d.recentLapsLed / d.recentRaces : 0;
    const recentLLNorm     = normalize(recentLLPerRace, recentLLArr);

    const raw = (histPctNorm      * 0.25)
              + (startPenaltyNorm * 0.25)
              + (qualSpeedNorm    * 0.20)
              + (projNorm         * 0.15)
              + (histRatingNorm   * 0.05)
              + (recentLLNorm     * 0.15)
              + (siteLapsNorm     * 0.05);

    let factor;
    if (d.siteRaces >= 5)      factor = 1.00;
    else if (d.siteRaces >= 3) factor = 0.85;
    else if (d.siteRaces >= 1) factor = 0.70;
    else                       factor = 0.50;

    d.domPts = Math.round(raw * factor * 100) / 100;
  }

  const sorted = drivers.slice().sort((a, b) => b.domPts - a.domPts);
  for (let i = 0; i < sorted.length; i++) {
    sorted[i].domRank = i + 1;
  }
}


/* -------------------------------------------------------
 *  3. Place Differential Projection
 *
 *  Built from our own data — no iFantasyRace finish proj.
 *  Reliability-adjusted by site race count.
 *  Drivers with no site history fall back to startPos - 20.
 * ------------------------------------------------------- */

function computePDProjection(drivers) {
  const fieldSize        = drivers.length;
  const projArr          = drivers.map(d => d.proj);
  const startArr         = drivers.map(d => d.startPos);
  const fieldMedianStart = percentile(startArr, 50);

  // Arrays for normalizing quality signals to 0-10 position-equivalent scale
  const ratingArr    = drivers.map(d => d.siteAvgRating);
  const finishInvArr = drivers.map(d => 40 - d.histAvgFinish);
  const formArr      = drivers.map(d => d.recentFormScore);

  for (const d of drivers) {
    if (d.siteRaces === 0) {
      const formContrib0 = (normalize(d.recentFormScore, formArr) / 100) * 10;
      const raw0 = (d.startPos - 20) * 0.70 + formContrib0 * 0.30;
      d.pdProj = Math.round(raw0 * 100) / 100;
      continue;
    }

    // iFantasyRace implied position gain
    const projNorm      = normalize(d.proj, projArr);
    const impliedFinish = 1 + (1 - projNorm / 100) * (fieldSize - 1);
    const projContrib   = d.startPos - impliedFinish;

    // Starting position vs field median (further back = more room to gain)
    const startContrib  = (d.startPos - fieldMedianStart) / 2;

    // Quality signals normalized 0-100 then scaled to 0-10 (position-equivalent)
    const ratingContrib = (normalize(d.siteAvgRating,      ratingArr)    / 100) * 10;
    const finishContrib = (normalize(40 - d.histAvgFinish, finishInvArr) / 100) * 10;
    const formContrib   = (normalize(d.recentFormScore,    formArr)      / 100) * 10;

    const raw = (d.histAvgStartFinishDiff * 0.30)
              + (projContrib              * 0.10)
              + (formContrib              * 0.20)
              + (ratingContrib            * 0.15)
              + (finishContrib            * 0.15)
              + (startContrib             * 0.10);

    let factor;
    if (d.siteRaces >= 5)      factor = 1.00;
    else if (d.siteRaces >= 3) factor = 0.85;
    else                       factor = 0.70;

    d.pdProj = Math.round(raw * factor * 100) / 100;
  }
}


/* -------------------------------------------------------
 *  4. Adjusted Projection
 *
 *  Starts with d.proj (iFantasyRace midpoint) and nudges
 *  via four weighted signals normalized against the field.
 * ------------------------------------------------------- */

function computeAdjProjection(drivers, raceContext) {
  const w = raceContext.weights;

  const domArr  = drivers.map(d => d.domPts);
  const pdArr   = drivers.map(d => d.pdProj);
  const formArr = drivers.map(d => d.recentFormScore);

  for (const d of drivers) {
    const domNorm = normalize(d.domPts, domArr);
    const domAdj  = clampAdj(
      ((domNorm - 50) / 50) * ADJ_PROJ_BOUNDS.MAX_DOM_ADJ,
      ADJ_PROJ_BOUNDS.MAX_DOM_ADJ
    );

    const pdNorm = normalize(d.pdProj, pdArr);
    const pdAdj  = clampAdj(
      ((pdNorm - 50) / 50) * ADJ_PROJ_BOUNDS.MAX_PD_ADJ,
      ADJ_PROJ_BOUNDS.MAX_PD_ADJ
    );

    const speedInverted = drivers.length - d.speedComposite;
    const speedInvArr   = drivers.map(d2 => drivers.length - d2.speedComposite);
    const speedNorm = normalize(speedInverted, speedInvArr);
    const speedAdj  = clampAdj(
      ((speedNorm - 50) / 50) * ADJ_PROJ_BOUNDS.MAX_SPEED_ADJ,
      ADJ_PROJ_BOUNDS.MAX_SPEED_ADJ
    );

    const histInverted = 40 - d.histAvgFinish;
    const histInvArr   = drivers.map(d2 => 40 - d2.histAvgFinish);
    const histNorm = normalize(histInverted, histInvArr);
    const histAdj  = clampAdj(
      ((histNorm - 50) / 50) * ADJ_PROJ_BOUNDS.MAX_HISTORY_ADJ,
      ADJ_PROJ_BOUNDS.MAX_HISTORY_ADJ
    );

    const formNorm = normalize(d.recentFormScore, formArr);
    const formAdj  = clampAdj(
      ((formNorm - 50) / 50) * ADJ_PROJ_BOUNDS.MAX_FORM_ADJ,
      ADJ_PROJ_BOUNDS.MAX_FORM_ADJ
    );

    d.adjProj = d.proj
      + (domAdj   * w.dom)
      + (pdAdj    * w.pd)
      + (speedAdj * w.speed)
      + (histAdj  * (w.history || 0))
      + formAdj;

    const floorClamp = d.floor > 0
      ? Math.min(d.floor, d.proj * 0.80)
      : d.proj * 0.80;

    d.adjProj = Math.max(d.adjProj, floorClamp);
    d.adjProj = Math.round(d.adjProj * 100) / 100;
  }
}

function clampAdj(val, maxAbs) {
  return Math.max(-maxAbs, Math.min(maxAbs, val));
}


/* -------------------------------------------------------
 *  4b. TrackType Blend
 *
 *  Blends domPts, pdProj, and adjProj with TrackType-derived
 *  signals after each site-specific computation.
 *
 *  Blend ratio by siteRaces:
 *    10+  → 75% site / 25% TrackType
 *    5-9  → 60% site / 40% TrackType
 *    3-4  → 50% site / 50% TrackType
 *    1-2  → 35% site / 65% TrackType
 *    0    →  0% site / 100% TrackType
 *
 *  If the TrackType sheet was not filled in (all tt fields = 0),
 *  each blend function exits early and leaves values unchanged.
 * ------------------------------------------------------- */

function getTrackTypeBlendWeights(siteRaces) {
  if (siteRaces >= 10) return { siteW: 0.75, ttW: 0.25 };
  if (siteRaces >= 5)  return { siteW: 0.60, ttW: 0.40 };
  if (siteRaces >= 3)  return { siteW: 0.50, ttW: 0.50 };
  if (siteRaces >= 1)  return { siteW: 0.35, ttW: 0.65 };
  return                      { siteW: 0.00, ttW: 1.00 };
}

function applyTTDomBlend(drivers) {
  const ttArr = drivers.map(d => d.ttLapsLedPerRace);
  if (ttArr.every(v => v === 0)) return;

  for (const d of drivers) {
    const { siteW, ttW } = getTrackTypeBlendWeights(d.siteRaces);
    const ttNorm = normalize(d.ttLapsLedPerRace, ttArr);
    d.domPts = Math.round((siteW * d.domPts + ttW * ttNorm) * 100) / 100;
  }

  // Re-rank after blend
  const sorted = drivers.slice().sort((a, b) => b.domPts - a.domPts);
  for (let i = 0; i < sorted.length; i++) sorted[i].domRank = i + 1;
}

function applyTTPDBlend(drivers) {
  if (drivers.every(d => d.ttSFDiff === 0)) return;

  for (const d of drivers) {
    const { siteW, ttW } = getTrackTypeBlendWeights(d.siteRaces);
    d.pdProj = Math.round((siteW * d.pdProj + ttW * d.ttSFDiff) * 100) / 100;
  }
}

function applyTTAdjBlend(drivers) {
  const ttArr = drivers.map(d => d.ttRating);
  if (ttArr.every(v => v === 0)) return;

  for (const d of drivers) {
    const { siteW, ttW } = getTrackTypeBlendWeights(d.siteRaces);
    const ttRatingNorm = normalize(d.ttRating, ttArr);
    const ttBonus      = clampAdj(((ttRatingNorm - 50) / 50) * ADJ_PROJ_BOUNDS.MAX_HISTORY_ADJ,
                                   ADJ_PROJ_BOUNDS.MAX_HISTORY_ADJ);
    const ttAdjProj    = d.proj + ttBonus;
    d.adjProj = siteW * d.adjProj + ttW * ttAdjProj;

    // Preserve floor clamp from computeAdjProjection
    const floorClamp = d.floor > 0
      ? Math.min(d.floor, d.proj * 0.80)
      : d.proj * 0.80;
    d.adjProj = Math.max(d.adjProj, floorClamp);
    d.adjProj = Math.round(d.adjProj * 100) / 100;
  }
}


/* -------------------------------------------------------
 *  5. Edge Calculation
 *
 *  Sets edge = 0 for all when no ownership data entered.
 * ------------------------------------------------------- */

function computeEdge(drivers) {
  const ownArr = drivers.map(d => d.ownPct);
  const avgOwn = ownArr.reduce((a, b) => a + b, 0) / ownArr.length;

  if (avgOwn === 0) {
    for (const d of drivers) d.edge = 0;
    return;
  }

  const projArr = drivers.map(d => d.adjProj);
  const avgProj = projArr.reduce((a, b) => a + b, 0) / projArr.length;

  for (const d of drivers) {
    const ownershipImplied = (d.ownPct / avgOwn) * avgProj;
    d.edge = Math.round((d.adjProj - ownershipImplied) * 10) / 10;
  }
}


/* -------------------------------------------------------
 *  6. Group Assignment
 *
 *  Active groups: DOM, PD, LEVERAGE, UNDER.
 *  CORE removed — those drivers are ungrouped (Fill).
 *  Priority: DOM → PD → LEVERAGE → UNDER
 * ------------------------------------------------------- */

function assignGroups(drivers, raceContext) {
  const trackType = raceContext.trackType;
  const T = GROUP_THRESHOLDS;

  const targetDomsMax = (raceContext.targetDoms && raceContext.targetDoms.max !== undefined)
    ? raceContext.targetDoms.max
    : (raceContext.targetDoms || 0);

  const hasOwnership = drivers.some(d => d.ownPct > 0);

  const edgeArr  = drivers.map(d => d.edge);
  const edgeP75  = percentile(edgeArr, 75);
  const edgeP25  = percentile(edgeArr, 25);

  const avgOwn = hasOwnership
    ? drivers.reduce((s, d) => s + d.ownPct, 0) / drivers.length
    : 0;

  const adjProjArr    = drivers.map(d => d.adjProj);
  const medianAdjProj = percentile(adjProjArr, 50);

  for (const d of drivers) {
    const tags = [];

    // --- DOMINATOR ---
    if (targetDomsMax > 0
        && d.domPts   >= 30
        && d.startPos <= getDOMMaxStart(trackType)) {
      tags.push("DOM");
    }

    // --- PD VALUE ---
    // Qualifies via current projection, proven track history, or start deviation
    const pdByProj      = d.pdProj >= T.PD_MIN_PROJ_PD && d.histAvgStartFinishDiff > 0;
    const pdMinStart    = getPDMinStart(trackType);
    const pdByHistory   = d.histAvgStartFinishDiff >= 5 && d.startPos >= pdMinStart;
    d.startDeviation    = d.histAvgStart > 0 ? d.histAvgStart - d.startPos : 0;
    const pdByDeviation = d.startDeviation <= -10;
    if (pdByProj || pdByHistory || pdByDeviation) {
      tags.push("PD");
    }

    // --- LEVERAGE (ownership-dependent) ---
    if (hasOwnership
        && d.edge    >= edgeP75
        && d.edge    >  T.LEVERAGE_MIN_EDGE
        && d.ownPct  <  T.LEVERAGE_MAX_OWN
        && d.adjProj >  medianAdjProj) {
      tags.push("LEVERAGE");
    }

    // --- UNDER (ownership-dependent) ---
    if (hasOwnership
        && d.edge   <= edgeP25
        && d.edge   <  0
        && d.ownPct >  avgOwn) {
      tags.push("UNDER");
    }

    // Everything else is ungrouped — falls into Fill pool
    d.group = tags.join(" ");
  }
}


/* -------------------------------------------------------
 *  7. Cash Score Calculation
 * ------------------------------------------------------- */

function computeCashScores(drivers) {
  const cw = CASH_WEIGHTS;

  const ownArr  = drivers.map(d => d.ownPct);
  const valArr  = drivers.map(d => d.salary > 0 ? d.adjProj / (d.salary / 1000) : 0);
  const formArr = drivers.map(d => d.recentFormScore);

  for (const d of drivers) {
    const chalkNorm = normalize(d.ownPct,            ownArr);
    const rawValue  = d.salary > 0 ? d.adjProj / (d.salary / 1000) : 0;
    const valueNorm = normalize(rawValue,            valArr);
    const formNorm  = normalize(d.recentFormScore,   formArr);

    d.cashScore = (d.floor   * cw.floorW)
                + (d.adjProj * cw.projW)
                + (Math.max(0, d.pdProj) * cw.pdW)
                - (d.dkStd   * cw.stdPenalty)
                + (chalkNorm * cw.chalkW)
                + (valueNorm * cw.valueW)
                + (formNorm  * cw.formW);

    d.cashScore = Math.round(d.cashScore * 100) / 100;
  }
}


/* -------------------------------------------------------
 *  Track History Score
 *
 *  Composite of historical signals at this track.
 *  Normalized to 0-100. Drivers with no history score ~0.
 *
 *  Components:
 *    histAvgFinish  — lower = better (inverted, weighted 40%)
 *    histRating     — higher = better (weighted 35%)
 *    histTop15Pct   — higher = better (weighted 25%)
 *
 *  Drivers with all default values (histAvgFinish=25,
 *  histRating=0, histTop15Pct=0) score near zero.
 * ------------------------------------------------------- */

function computeTrackHistScore(drivers) {
  // Build arrays for normalization
  // Invert avgFinish: lower finish = better = higher score
  const finishInv = drivers.map(d => 40 - d.histAvgFinish);
  const ratingArr  = drivers.map(d => d.histRating);
  const top15Arr   = drivers.map(d => d.histTop15Pct);

  for (const d of drivers) {
    const finishNorm = normalize(40 - d.histAvgFinish, finishInv);
    const ratingNorm = normalize(d.histRating,         ratingArr);
    const top15Norm  = normalize(d.histTop15Pct,       top15Arr);

    // Penalty for no history: if all three are at defaults, score is near zero
    const hasHistory = d.histAvgFinish !== 25 || d.histRating > 0 || d.histTop15Pct > 0;
    const histPenalty = hasHistory ? 1.0 : 0.2;

    const raw = (finishNorm * 0.40)
              + (ratingNorm * 0.35)
              + (top15Norm  * 0.25);

    d.trackHistScore = Math.round(raw * histPenalty * 10) / 10;
  }
}


/* -------------------------------------------------------
 *  TrackType History Score
 *
 *  Composite of TrackType signals, normalized 0-100.
 *  Uses the confidence-discounted tt fields stored on each driver.
 *
 *  Components:
 *    ttRating         — higher = better (weighted 50%)
 *    ttLapsLedPerRace — higher = better (weighted 30%)
 *    ttSFDiff         — higher = better (weighted 20%)
 *
 *  Then multiplied by the same confidence factor as Step 1:
 *    ttRaces >= 10 → 1.00   5-9 → 0.85   3-4 → 0.70
 *    ttRaces 1-2  → 0.50   ttRaces 0 → 0.00
 * ------------------------------------------------------- */

function computeTTHistScore(drivers) {
  const ratingArr = drivers.map(d => d.ttRating);
  const lapsArr   = drivers.map(d => d.ttLapsLedPerRace);
  const sfArr     = drivers.map(d => d.ttSFDiff);

  for (const d of drivers) {
    const ratingNorm = normalize(d.ttRating,         ratingArr);
    const lapsNorm   = normalize(d.ttLapsLedPerRace, lapsArr);
    const sfNorm     = normalize(d.ttSFDiff,         sfArr);

    const raw = (ratingNorm * 0.50)
              + (lapsNorm   * 0.30)
              + (sfNorm     * 0.20);

    let factor;
    if      (d.ttRaces >= 10) factor = 1.00;
    else if (d.ttRaces >= 5)  factor = 0.85;
    else if (d.ttRaces >= 3)  factor = 0.70;
    else if (d.ttRaces >= 1)  factor = 0.50;
    else                      factor = 0.00;

    d.ttHistScore = Math.round(raw * factor * 10) / 10;
  }
}


/* -------------------------------------------------------
 *  Cash Core Grade & Group Assignment
 *
 *  Cash Core grade rewards floor value + track history.
 *  Formula:
 *    cashCoreGrade = (floor × 0.35)
 *                  + (adjProj × 0.25)
 *                  + (valueNorm × 0.20)    ← proj*1000/salary
 *                  + (trackHistNorm × 0.20)
 *
 *  Top 15% of all drivers by cash grade are tagged CASHCORE.
 *  A driver can hold CASHCORE alongside any other group tag.
 * ------------------------------------------------------- */

function assignCashCoreGroup(drivers) {
  const eligible = drivers.slice();
  if (eligible.length === 0) return;

  // Normalize value and track history across eligible pool
  const valueArr    = eligible.map(d => d.salary > 0 ? d.adjProj / (d.salary / 1000) : 0);
  const trackArr    = eligible.map(d => d.trackHistScore);
  const floorArr    = eligible.map(d => d.floor);
  const adjProjArr  = eligible.map(d => d.adjProj);

  // Compute cash core grade for each eligible driver
  for (const d of eligible) {
    const rawValue   = d.salary > 0 ? d.adjProj / (d.salary / 1000) : 0;
    const valueNorm  = normalize(rawValue,        valueArr);
    const trackNorm  = normalize(d.trackHistScore, trackArr);
    const floorNorm  = normalize(d.floor,          floorArr);
    const projNorm   = normalize(d.adjProj,        adjProjArr);

    d.cashCoreGrade = Math.round(
      (floorNorm  * 0.35)
    + (projNorm   * 0.25)
    + (valueNorm  * 0.20)
    + (trackNorm  * 0.20)
    );
  }

  // Top 15% → CASHCORE group
  const threshold = Math.max(1, Math.ceil(eligible.length * 0.15));
  const sorted    = eligible.slice().sort((a, b) => b.cashCoreGrade - a.cashCoreGrade);

  for (let i = 0; i < threshold; i++) {
    const d = sorted[i];
    d.group = d.group ? d.group + " CASHCORE" : "CASHCORE";
  }
}


/* -------------------------------------------------------
 *  8. Cash Lineup Builder
 * ------------------------------------------------------- */

function buildCashLineup(drivers) {
  const cap  = CASH_SALARY_CAP;
  const size = CASH_ROSTER_SIZE;

  const pool = drivers.filter(d => d.salary > 0 && d.floor > 0);
  if (pool.length < size) return pool.slice(0, size);

  const hasOwnership = pool.some(d => d.ownPct > 0);
  const freeSquare   = pool.slice().sort((a, b) =>
    hasOwnership ? b.ownPct - a.ownPct : b.floor - a.floor
  )[0];

  let lineup   = [freeSquare];
  let totalSal = freeSquare.salary;

  const remaining = pool
    .filter(d => d.name !== freeSquare.name)
    .sort((a, b) => {
      const aVal = a.cashScore / (a.salary / 1000);
      const bVal = b.cashScore / (b.salary / 1000);
      return bVal - aVal;
    });

  for (const d of remaining) {
    if (lineup.length >= size) break;
    if (totalSal + d.salary <= cap) {
      lineup.push(d);
      totalSal += d.salary;
    }
  }

  if (lineup.length < size) {
    for (const d of remaining) {
      if (lineup.length >= size) break;
      if (!lineup.some(ld => ld.name === d.name)) lineup.push(d);
    }
  }

  let improved = true, iterations = 0;
  while (improved && iterations < 50) {
    improved = false;
    iterations++;
    for (let i = 1; i < lineup.length; i++) {
      const salWithout = lineup.reduce((s, d, j) => j !== i ? s + d.salary : s, 0);
      const budget     = cap - salWithout;
      const curScore   = lineup[i].cashScore;
      let   bestSwap   = null;

      for (const candidate of pool) {
        if (lineup.some(d => d.name === candidate.name)) continue;
        if (candidate.salary > budget) continue;
        if (candidate.cashScore > curScore) {
          if (!bestSwap || candidate.cashScore > bestSwap.cashScore) bestSwap = candidate;
        }
      }

      if (bestSwap) { lineup[i] = bestSwap; improved = true; }
    }
  }

  lineup.sort((a, b) => b.salary - a.salary);
  return lineup;
}