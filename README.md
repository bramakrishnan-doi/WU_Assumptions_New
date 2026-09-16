# Lower Basin Water Use Report Generator

> **Detailed documentation:** See `APP_USER_RULES_MAINTENANCE_GUIDE.md` for end-user instructions, rule editing, testing, troubleshooting, and application maintenance.
> **Developer calculation/slot changes:** See `DEVELOPER_CALCULATION_SLOT_CHANGE_GUIDE.md` for step-by-step instructions for adding RiverWare slots, changing formulas, adding bullets, Source tracing, and regression testing.
> **Guided change planner:** Run `run_change_planner.bat` (or `report_change_planner.py`) to answer a structured questionnaire and generate a reviewable Markdown implementation plan. See `REPORT_CHANGE_PLANNER_GUIDE.md`.


A local Flask application for generating Lower Basin Projected Water Use Word reports from RiverWare 24-Month Study model files and the current monthly Projected State Use workbook.

## What changed in this revision

The application now supports batch generation for any deliberate combination of Most Probable, Probable Minimum, and Probable Max models. It validates all selected files before generation, reads each scenario from the correct Excel block, and renders both the browser preview and Word output from one canonical report model. The preview is read-only and includes **Source** controls that trace report values to the RiverWare slot, Excel cell, calculation, or policy rule that produced them.

The report-policy rules are versioned in `config/report_policy.json`. Click the **Rules** control in the application header to review the exact active configuration. The active version is displayed in the UI but is **not written into the generated Word report**. The Word renderer continues to use `reference_report.docx` as the authoritative visual template.

## Windows quick start

1. Install **Python 3.11 or newer** if it is not already installed. During a python.org installation, enable the option to add Python to PATH.
2. Extract this project to a normal writable folder.
3. Double-click **`run_app.bat`**. On first use it creates a local `.venv` and installs the dependencies. Subsequent launches reuse that environment.
4. The application opens at `http://127.0.0.1:5051`. Keep the command window open while using it.

The server binds only to `127.0.0.1`; it is intended as a local workstation application.

## Monthly workflow

1. Upload the current Projected State Use `.xlsx` workbook.
2. Enable the scenarios needed for that run. Most Probable and Probable Minimum are enabled by default. Probable Max is optional; either switch its **Include** toggle on or simply choose/drop a Max model and the application will enable it automatically.
3. Assign the corresponding `.mdl` or `.mdl.gz` RiverWare model to each enabled scenario.
4. Select the 24-Month Study month.
5. Click **Run preflight**. Generation remains disabled until validation succeeds.
6. Review the detected study period, report years, model filenames, slot counts, workbook, policy version, and any warnings. Use the header **Rules** control whenever you need to inspect the configuration currently in force.
7. Click **Generate report(s)**. Use the scenario tabs to review each read-only preview and the **Source** controls to inspect provenance.
8. Download an individual Word report or, for a multi-scenario run, the batch ZIP.

A successful preflight is cached in memory for 30 minutes, so generation does not re-parse the uploaded models and workbook. Changing a source file or scenario selection invalidates the preflight and requires validation again.

## Important validation behavior

The application reads the year-sheet blocks defined by the active policy: Most from A/B/D, Min from F/G/I, and Max from K/L/N. It checks required year sheets, required report labels, formula-cache availability for key values, System Conservation year columns, the state-level conservation summary table, all configured RiverWare annual slots, model study dates, supported volume units, and obvious scenario/model filename mismatches.

If a workbook contains formulas, Python does not calculate them. The application reads values cached by Excel. If values are missing—or if the workbook is set to Manual calculation—the preflight warns or fails as appropriate. Recalculate and save the workbook in Excel before using it for a report.

## Generated Word documents

Word files inherit styles, page setup, list definitions, and table formatting from `reference_report.docx`. Source traceability, model filenames, workbook filenames, policy version, and generation timestamps are intentionally **not** inserted into the Word document.

The intentionally retained conservation content is the **state-level “Modeled Lower Basin Conservation Actions” summary table only**. The older detailed conservation-actions table is not generated. The SEIS ROD note is no longer collected or printed. The Notes and Disclaimers section contains only the configured conservation disclaimer.

The Powell-release subtitle is shown only for the Most Probable scenario. Probable Minimum and Probable Max omit that second subtitle line. ICS and conservation-table rows are rendered at 0.25 inch with centered cell content, and there is no forced page break between those two tables.

When the calendar-year sheet identifies the California water-left-in-Mead entry with the configured note `MWD System Conservation`, that value is rolled into **Total California System Conservation** and is not printed as a separate MWD bullet. Other labels, such as `Other water left in Mead`, remain separate.

## Project structure

- `app.py` — Flask routes and HTTP concerns only.
- `report_service.py` — batch preflight, scenario assignment checks, and short-lived parsed-input cache.
- `riverware_annual.py` — direct `.mdl`/`.mdl.gz` annual-slot extraction and unit conversion.
- `excel_source.py` — scenario-aware Excel reading and workbook validation.
- `policy.py` + `config/report_policy.json` — versioned report/configuration rules.
- `report_models.py` — canonical report and source-trace data structures.
- `report_builder.py` — report calculations and narrative construction.
- `docx_writer.py` — Word-only rendering from the canonical report model.
- `templates/report_preview.html` — read-only HTML renderer of that same model.
- `templates/index.html`, `static/` — application UI.
- `tests/` — regression tests, including an optional real-model integration test.
- `report_change_planner.py` + `run_change_planner.bat` — non-destructive guided planner for new RiverWare-derived bullets/calculation changes.
- `REPORT_CHANGE_PLANNER_GUIDE.md` — instructions for using the planner.
- `MAINTENANCE_GUIDE.md` — guidance for policy, workbook, RiverWare, and template changes.

## Editing report rules

The business/report rules are intentionally editable in `config/report_policy.json`. This includes scenario display names, Excel scenario columns, default-enabled scenarios, which scenarios show the Powell-release subtitle or conservation summary, effective-year wording, the Arizona display override, the California conservation roll-up labels, workbook source locations, and RiverWare slot requirements.

After editing the JSON file, **restart the application** so the new rules are loaded. Increment `policy_version` whenever a change can affect report content or source interpretation. The Rules panel shows the exact configuration loaded by the running app.

Not every application behavior belongs in that JSON file. Parsing algorithms, arithmetic mechanics, DOCX layout code, Flask behavior, and general UI mechanics remain Python/CSS/JavaScript code so the policy file stays understandable rather than becoming a programming language.

## Tests

From a command prompt in the project folder:

```bat
.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.venv\Scripts\python.exe -m pytest -q
```

Or run `run_tests.bat` after the application environment has been created.

To validate a real/sanitized RiverWare model without adding it to the repository:

```bat
set WU_TEST_MODEL=C:\path\to\model.mdl.gz
.venv\Scripts\python.exe -m pytest -q -m integration
```

See `MAINTENANCE_GUIDE.md` before changing report policy or replacing the Word reference document.
