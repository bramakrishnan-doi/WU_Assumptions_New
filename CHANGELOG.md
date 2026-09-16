# Changelog

## 2026.09.1

- Made the header Rules badge clickable and added a read-only active-rules panel, including the complete loaded JSON configuration and maintenance/edit instructions.
- Clarified Probable Max selection: Max remains optional, but its model picker is no longer blocked; selecting or dropping a Max model automatically includes that scenario.
- Renamed the Min report display label from `Probable Min` to `Probable Minimum`.
- Removed the SEIS ROD input and removed that second Notes and Disclaimers bullet from previews and Word output.
- Limited the `For the 6 maf & 7 maf Powell Release Scenarios` subtitle to Most Probable reports.
- Added a policy-driven California conservation roll-up: a water-left-in-Mead entry labeled `MWD System Conservation` is added to Total California System Conservation and no longer printed separately.
- Set ICS and conservation-table rows to an exact 0.25 inch and centered table content horizontally and vertically.
- Removed the forced page break between the ICS table and the conservation summary table.
- Expanded regression coverage for the new policy, report-content, and Word-layout behavior.

## 2026.09

- Added versioned `config/report_policy.json` and policy validation.
- Corrected Most/Min/Max Excel year-sheet block selection.
- Corrected Arizona state and basin totals to use `AnnualWaterUse.AzTotalAnnual` rather than Arizona apportionment.
- Added multi-scenario batch preflight and generation for any selected combination of Most, Min, and Max.
- Added model filename/scenario mismatch checks and common-start-year validation.
- Added a 30-minute in-memory parsed-input cache so generation does not re-parse preflighted sources.
- Added source traceability for RiverWare slots, Excel cells, calculations, and policy overrides.
- Replaced editable HTML with a read-only preview and removed the alternate HTML-to-Word renderer.
- Unified HTML and Word rendering around a canonical `ReportContext`.
- Preserved the state-level conservation summary and intentionally omitted the older detailed conservation-actions table.
- Hardened RiverWare unit handling; missing/unsupported units now fail rather than silently assuming acre-feet.
- Moved effective-year Nevada shortage/system-conservation wording into policy and aligned the post-2026 language with the supplied August reference report.
- Made System Conservation block header discovery tolerant of rows inserted above configured tables.
- Added conversion-noise snapping before integer report rounding to avoid one-acre-foot boundary artifacts.
- Scrubbed the reference template metadata and removed stale/custom/timestamp metadata from generated Word packages.
- Made optional RiverWare slots part of the parsed policy inventory instead of a validation-only concept.
- Made Word bullet numbering discovery template-aware instead of relying on a hard-coded `numId`.
- Added regression tests plus an optional external real-model integration test.
- Disabled Flask debug mode for normal launches and kept the application bound to localhost.
- Replaced the bundled virtual environment with a clean Windows setup/launcher workflow.
- Removed monthly workbook/sample/debug artifacts from the distributable project.

## Maintenance tooling

- Added `report_change_planner.py`, a non-destructive interactive planner for new RiverWare-derived bullets and calculation changes.
- Added `run_change_planner.bat` and `REPORT_CHANGE_PLANNER_GUIDE.md` for Windows maintainers.
- The planner can detect already-registered slots, propose policy/source/calculation/test changes, and generate a Markdown implementation plan without modifying operational code.
