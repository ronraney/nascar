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