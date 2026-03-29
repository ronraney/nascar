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

  computeDomPoints(drivers, raceContext);
  computePDProjection(drivers);
  computeAdjProjection(drivers, raceContext);
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

  // ---- Step 10: Cash Core Grade + Group Assignment ----
  // Must run after track history score and cash score are computed.
  // Reassigns top 15% of non-DOM/PD drivers to CASHCORE group.
  assignCashCoreGroup(drivers);

  return drivers;
}


/* -------------------------------------------------------
 *  2. Dominator Points Calculation
 *
 *  Six normalized signals weighted and combined.
 *  Reliability-adjusted by site race count.
 *  domScore/domLabel from Playability are NOT used here.
 * ------------------------------------------------------- */

function computeDomPoints(drivers, raceContext) {
  const histPctArr    = drivers.map(d => d.histPctLapsLed);
  const projArr       = drivers.map(d => d.proj);
  const qualSpeedArr  = drivers.map(d => d.qualSpeed);
  const histRatingArr = drivers.map(d => d.histRating);
  const siteLapsArr   = drivers.map(d => d.siteLapsLed);

  // startPos penalty: logarithmic — P1 gets highest raw penalty value (best)
  // rawPenalty = 1 / log(startPos + e - 1), then normalize across field
  const startPenaltyArr = drivers.map(d => 1.0 / Math.log(d.startPos + Math.E - 1));

  for (const d of drivers) {
    const histPctNorm      = normalize(d.histPctLapsLed, histPctArr);
    const startPenaltyNorm = normalize(1.0 / Math.log(d.startPos + Math.E - 1), startPenaltyArr);
    const qualSpeedNorm    = normalize(d.qualSpeed,      qualSpeedArr);
    const projNorm         = normalize(d.proj,           projArr);
    const histRatingNorm   = normalize(d.histRating,     histRatingArr);
    const siteLapsNorm     = normalize(d.siteLapsLed,    siteLapsArr);

    const raw = (histPctNorm      * 0.30)
              + (startPenaltyNorm * 0.25)
              + (qualSpeedNorm    * 0.20)
              + (projNorm         * 0.15)
              + (histRatingNorm   * 0.10)
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
  const fieldSize         = drivers.length;
  const projArr           = drivers.map(d => d.proj);
  const startArr          = drivers.map(d => d.startPos);
  const histTop15Arr      = drivers.map(d => d.histTop15Pct);
  const fieldMedianStart  = percentile(startArr, 50);

  for (const d of drivers) {
    if (d.siteRaces === 0) {
      d.pdProj = Math.round((d.startPos - 20) * 100) / 100;
      continue;
    }

    const projNorm       = normalize(d.proj, projArr);
    const impliedFinish  = 1 + (1 - projNorm / 100) * (fieldSize - 1);
    const projContrib    = d.startPos - impliedFinish;

    const histTop15Norm  = normalize(d.histTop15Pct, histTop15Arr);
    const histTop15Contrib = (histTop15Norm / 100) * 10;

    const startContrib   = (d.startPos - fieldMedianStart) / 2;

    const raw = (d.histAvgStartFinishDiff * 0.35)
              + (projContrib              * 0.30)
              + (histTop15Contrib         * 0.20)
              + (startContrib             * 0.15);

    let factor;
    if (d.siteRaces >= 5)      factor = 1.00;
    else if (d.siteRaces >= 3) factor = 0.85;
    else                       factor = 0.70;   // siteRaces 1-2

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

    d.adjProj = d.proj
      + (domAdj   * w.dom)
      + (pdAdj    * w.pd)
      + (speedAdj * w.speed)
      + (histAdj  * (w.history || 0));

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
    d.group = "";

    // --- DOMINATOR ---
    if (targetDomsMax > 0
        && d.domRank  <= (targetDomsMax + 2)
        && d.domPts   >  0
        && d.startPos <= T.DOM_MAX_START_POS) {
      d.group = "DOM";
      continue;
    }

    // --- PD VALUE ---
    // Qualifies via current projection OR proven track history
    const pdByProj    = d.pdProj >= T.PD_MIN_PROJ_PD && d.histAvgStartFinishDiff > 0;
    const pdByHistory = d.histAvgStartFinishDiff >= 5 && d.startPos >= T.PD_MIN_START_POS;
    if (pdByProj || pdByHistory) {
      d.group = "PD";
      continue;
    }

    // --- LEVERAGE (ownership-dependent) ---
    if (hasOwnership
        && d.edge    >= edgeP75
        && d.edge    >  T.LEVERAGE_MIN_EDGE
        && d.ownPct  <  T.LEVERAGE_MAX_OWN
        && d.adjProj >  medianAdjProj) {
      d.group = "LEVERAGE";
      continue;
    }

    // --- UNDER (ownership-dependent) ---
    if (hasOwnership
        && d.edge   <= edgeP25
        && d.edge   <  0
        && d.ownPct >  avgOwn) {
      d.group = "UNDER";
      continue;
    }

    // Everything else is ungrouped — falls into Fill pool
  }
}


/* -------------------------------------------------------
 *  7. Cash Score Calculation
 * ------------------------------------------------------- */

function computeCashScores(drivers) {
  const cw = CASH_WEIGHTS;

  const ownArr = drivers.map(d => d.ownPct);
  const valArr = drivers.map(d => d.salary > 0 ? d.adjProj / (d.salary / 1000) : 0);

  for (const d of drivers) {
    const chalkNorm = normalize(d.ownPct, ownArr);
    const rawValue  = d.salary > 0 ? d.adjProj / (d.salary / 1000) : 0;
    const valueNorm = normalize(rawValue, valArr);

    d.cashScore = (d.floor   * cw.floorW)
                + (d.adjProj * cw.projW)
                + (Math.max(0, d.pdProj) * cw.pdW)
                - (d.dkStd   * cw.stdPenalty)
                + (chalkNorm * cw.chalkW)
                + (valueNorm * cw.valueW);

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
 *  Cash Core Grade & Group Assignment
 *
 *  Cash Core grade rewards floor value + track history.
 *  Formula:
 *    cashCoreGrade = (floor × 0.35)
 *                  + (adjProj × 0.25)
 *                  + (valueNorm × 0.20)    ← proj*1000/salary
 *                  + (trackHistNorm × 0.20)
 *
 *  Top 15% of non-DOM, non-PD drivers are assigned
 *  to the CASHCORE group. ~5-6 drivers at a 37-driver slate.
 * ------------------------------------------------------- */

function assignCashCoreGroup(drivers) {
  // Only consider drivers not already in DOM or PD
  const eligible = drivers.filter(d => d.group !== "DOM" && d.group !== "PD");
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
    sorted[i].group = "CASHCORE";
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