Task: Add TrackType as a new data source in DataPipeline.gs, blend into existing scoring in Analysis.gs.
Sheet: TrackType — manual paste each week, columns: Driver, Avg Fin, Races, Laps Led, Avg St, Rating
Step 1 — loadTrackType() in DataPipeline.gs:
Compute these fields per driver:

ttLapsLedPerRace = Laps Led / Races
ttSFDiff = Avg St - Avg Fin
ttRating = Rating
ttRaces = Races

Apply confidence discount to all three signals:

ttRaces >= 10: factor 1.0
ttRaces 5-9: factor 0.85
ttRaces 3-4: factor 0.70
ttRaces < 3: factor 0.50

Step 2 — Blend ratio driven by siteRaces in Analysis.gs:
siteRacesTrack-SpecificTrackType10+75%25%5-960%40%3-450%50%1-235%65%00%100%
Step 3 — Apply blend to:

adjProj
domPts
pdProj

No new Dashboard columns. Effect surfaces through existing adjProj, DOM, and PD values.
Read Config.gs for all sheet name constants before writing anything.

Task: Add TrackType as a new data source in DataPipeline.gs, blend into existing scoring in Analysis.gs, add ttHistScore column to Dashboard.
Sheet: TrackType — manual paste each week, columns: Driver, Avg Fin, Races, Laps Led, Avg St, Rating
Step 1 — loadTrackType() in DataPipeline.gs:
Compute these fields per driver:

ttLapsLedPerRace = Laps Led / Races
ttSFDiff = Avg St - Avg Fin
ttRating = Rating
ttRaces = Races

Apply confidence discount to all three signals:

ttRaces >= 10: factor 1.0
ttRaces 5-9: factor 0.85
ttRaces 3-4: factor 0.70
ttRaces < 3: factor 0.50

Step 2 — Blend ratio driven by siteRaces in Analysis.gs:
siteRacesTrack-SpecificTrackType10+75%25%5-960%40%3-450%50%1-235%65%00%100%
Step 3 — Apply blend to:

adjProj
domPts
pdProj

Step 4 — computeTTHistScore() in Analysis.gs:
Formula (normalized 0-100):

ttRating: 50%
ttLapsLedPerRace: 30%
ttSFDiff: 20%

Multiplied by confidence discount based on ttRaces (same as Step 1).
Step 5 — Dashboard column:
Add ttHistScore as a visible column immediately before Notes.
No other new Dashboard columns. Effect of blend surfaces through existing adjProj, DOM, and PD values.
Read Config.gs for all sheet name constants before writing anything.

Task
Replace DOM_MAX_START_POS: 15 in GROUP_THRESHOLDS with a new function getDOMMaxStart(trackType) in Config.gs:
Superspeedway:            0  (no DOM group at all)
Superspeedway (Drafting): 0
Intermediate:             15
Short Track (Flat):       12
Short Track (Steep):      12
Short Track (Fast):       12
Short Track:              12
Short Track (Wear):       12
Road Course:              10
Street Course:            0
Large Oval:               15
Large Triangle:           15
1-Mile Flat:              15
In assignGroups in Analysis.gs, replace:
javascriptd.startPos <= T.DOM_MAX_START_POS
With:
javascriptd.startPos <= getDOMMaxStart(trackType)
Also replace the domRank condition with domPts threshold:
javascriptd.domPts >= 40
So the full new DOM condition becomes:
javascriptd.domPts >= 40
&& d.startPos <= getDOMMaxStart(trackType)
Condition 1 (targetDomsMax > 0) stays as-is — it correctly suppresses DOM entirely on superspeedways and street courses via Config.gs track type mapping.

Task
Add startDeviation signal in Analysis.gs:
startDeviation = d.histAvgStart - d.startPos
Positive = starting worse than historical average (situational PD opportunity).
Add to PD qualification in assignGroups: a driver qualifies for PD if they meet the existing criteria OR if startDeviation >= 10 (starting 10+ positions worse than their historical average).
Multi-tag behavior: these drivers would get "PD CASHCORE" if they also meet CASHCORE grade — same pattern as Hamlin and Bell already have.
histAvgStart is already on the driver object from the Data_Avg_Start loader, so no new data pipeline work needed.