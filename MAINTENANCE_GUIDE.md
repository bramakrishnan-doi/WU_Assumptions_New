# Maintenance Guide

## 1. Report policy is explicit and versioned

Business rules that are expected to change because of report policy, scenario conventions, operating guidance, or source layout belong in `config/report_policy.json`, not in Flask routes or Word/HTML rendering code. Increment `policy_version` whenever a rule change can affect generated content or source interpretation.

The current policy contains:

- scenario display names and Excel label/value/note columns;
- default-enabled scenarios in the UI;
- which scenarios include the state-level conservation summary or Powell-release subtitle;
- the Notes and Disclaimers conservation text;
- the California water-left-in-Mead note labels that roll into Total California System Conservation;
- effective-year total-use disclaimer wording;
- the effective year and displayed value for the Arizona reduction override;
- effective-year Mexico shortage/reduction wording;
- effective-year Nevada shortage and system-conservation wording;
- configured workbook sheet names, conservation ranges, and expected summary rows;
- the complete required/optional RiverWare slot inventory; and
- the number of report years extracted from each model.

Keep processing mechanics in Python. Do not turn the JSON file into a scripting language. A policy rule should describe **what changes and when**; parsing algorithms, arithmetic utilities, DOCX formatting, and HTTP behavior stay in code.

## 2. Scenario Excel mapping

The active mappings are (the Min display label is **Probable Minimum**):

| Scenario | Label | Value | Notes |
|---|---:|---:|---:|
| Most | A | B | D |
| Min | F | G | I |
| Max | K | L | N |

If the workbook layout changes, update the scenario entries in `report_policy.json`, increment the policy version, and update/run `tests/test_excel_source.py`.

System Conservation ranges are also configured in policy. Within each configured conservation block, the reader locates the `Contractor` header dynamically within the allowed range instead of assuming a fixed row, which tolerates rows inserted above the table. The state-level conservation summary is likewise located by finding a `State` header immediately followed by the model start year rather than relying on a fixed row number. Expected state/total row names are validated.

## 3. RiverWare slot changes

`riverware.required_slots` is the contract between the report and a compatible RiverWare model. Preflight fails if a required slot cannot be parsed. Add a slot to `optional_slots` only when omission is genuinely acceptable and the report builder safely handles its absence.

The parser currently supports annual DSeries/CSeries data, aggregate-series members, scalar values, `.mdl` and `.mdl.gz`, RiverWare's `24:00` annual timestamps, and configured volume-unit conversions. Unknown or missing units fail validation rather than silently assuming acre-feet.

Before accepting a RiverWare upgrade or a serialization-format change, run the optional integration test against a representative model:

```bat
set WU_TEST_MODEL=C:\path\to\representative.mdl.gz
.venv\Scripts\python.exe -m pytest -q -m integration
```

Direct model parsing is intentionally isolated in `riverware_annual.py`. If a supported RiverWare programmatic export/API becomes preferable later, that module can be replaced while leaving the report model and renderers intact.

## 4. Canonical report model

`report_models.ReportContext` is the single source of truth after calculations are complete. `report_builder.py` constructs it. Both `templates/report_preview.html` and `docx_writer.py` render it.

Do not put scenario/year calculations directly in either renderer. Doing so can make the preview and Word report disagree.

Each visible value can have one or more `SourceRef` records. These are used by the UI's **Source** drawer and can identify:

- a RiverWare model filename, slot, year, converted acre-foot value, and original unit;
- an Excel workbook, sheet/cell, label, note, and cached value;
- a calculated value and its formula description; or
- a policy-driven override/rule and active policy version.

These provenance records are intentionally UI-only and must not be printed into the generated Word report.

## 5. Arizona totals

For state totals, U.S. contractor totals, and Lower Basin total projected use, Arizona is sourced from `AnnualWaterUse.AzTotalAnnual`. `AnnualWaterUse.Arizona_Apportionment` is not interchangeable with modeled annual use and must not be substituted for it.

The displayed post-2026 Arizona water-use reduction override is a separate policy rule in `report_policy.json`; its source trace includes both the underlying RiverWare shortage slot and the policy override. Nevada's 2026-versus-2027 shortage terminology and system-conservation presentation are also effective-year policy rules, so future wording changes should be made in the policy rather than embedded in `report_builder.py`.

California conservation roll-up behavior is also policy-driven. If the `California- water left in Mead` entry has a note matching `report.rules.california_system_conservation_rollup.year_sheet_note_labels`, its sign-adjusted value is added into **Total California System Conservation** and the separate note-labeled bullet is suppressed. If the note does not match, the water-left-in-Mead entry remains a separate bullet.

## 6. Word template

`reference_report.docx` is the authoritative formatting source. `docx_writer.py` clears only the body content and preserves the document's styles, numbering definitions, and section/page setup.

The writer no longer relies on a hard-coded Word numbering ID. It searches the reference document for the bullet/dash hierarchy (`•` at level 0 and `–` at level 1) and uses a numbering instance associated with that definition. The shipped reference template is metadata-scrubbed, and generated documents clear stale author/custom metadata and created/modified timestamps without adding application provenance.


The ICS Totals and Modeled Lower Basin Conservation Actions tables use an exact 0.25-inch row height with horizontal and vertical cell centering. There is no forced page break before the conservation summary, allowing it to follow the ICS table on the same page when space permits. The Powell-release second subtitle line is emitted only for scenarios configured in `report.powell_release_subtitle_scenarios`.

After replacing `reference_report.docx`, always:

1. run the automated tests;
2. generate a representative Most report;
3. render the DOCX to PDF/PNG and visually inspect every page; and
4. verify bullets, state-heading blue values, highlighted disclaimer text, tables, margins, and page breaks.

## 7. Workbook formula cache

`openpyxl` does not calculate formulas. The application uses `data_only=True`, which reads the last values cached by Excel. Before a monthly workbook is used, it should be recalculated and saved in Excel. Preflight identifies missing key cached values and warns when the workbook's calculation mode is Manual.

## 8. Batch preflight/cache

A preflight parses each enabled RiverWare model one at a time and loads the Excel workbook once for all selected scenarios. Parsed results—not the original uploaded model bytes—are cached in memory for 30 minutes. This keeps report generation fast and avoids repeatedly expanding/processing large model files.

The cache is intentionally process-local. Restarting the Flask app clears it. No source files or provenance logs are persisted by the application.

## 9. Regression expectations

At minimum, regression tests should continue to verify:

- Most/Min/Max year-sheet columns are distinct;
- Arizona totals use `AzTotalAnnual`;
- policy effective-year wording and overrides, including Nevada terminology;
- Most-only state-level conservation summary visibility;
- preview/Word content originates from the canonical report model;
- the Word file does not contain source/provenance metadata; and
- a representative real model still exposes every configured required slot.

When an intentional policy change makes an old historical report differ from a new report, update the policy version and tests to document the new expected behavior rather than weakening validation.

## 10. Rules panel and editability

The header **Rules** control calls `/api/rules` and displays the exact `report_policy.json` content loaded by the current Flask process, along with a human-readable summary. This view is intentionally read-only to avoid accidental operational changes from the report-generation screen.

To edit report rules, update `config/report_policy.json`, increment `policy_version` when report content or source interpretation changes, run the tests, and restart the application. The restarted Rules panel should show the new version and values.
