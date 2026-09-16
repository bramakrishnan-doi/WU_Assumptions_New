# Lower Basin Water Use Report Generator
## User, Rules, and Maintenance Guide

**Application policy version documented:** `2026.09.1`  
**Primary platform:** Windows workstation  
**Application type:** Local Flask web application  
**Rules file:** `config/report_policy.json`

---

## 1. Purpose of this guide

This guide explains how to:

- install and start the application on Windows;
- generate one or more Lower Basin Water Use reports;
- use preflight validation and the Source trace feature;
- understand the role of `config/report_policy.json`;
- safely edit report rules;
- validate and test a rule change;
- maintain Excel mappings, RiverWare slot requirements, and the Word template;
- troubleshoot common problems; and
- prepare a controlled application release after changes.

The central maintenance principle is:

> **Report/business policy belongs in `config/report_policy.json` when it can be expressed as configuration. Processing algorithms, calculations that require programming logic, RiverWare parsing, Word layout mechanics, Flask behavior, and UI behavior remain in code.**

The Rules panel in the application is intentionally read-only. It shows the exact policy that the currently running Flask process loaded.

---

## 2. Is `report_policy.json` where rules should be changed?

**Yes.** For the rule types described in this guide, `config/report_policy.json` is the intended source of truth.

Examples of rules that belong in JSON include:

- scenario display names;
- scenario-specific Excel columns;
- which scenarios are enabled by default;
- which scenarios show the Powell release subtitle;
- which scenarios include the conservation summary table;
- report disclaimer wording by effective year;
- Mexico and Nevada wording by effective year;
- the Arizona displayed reduction override;
- California conservation roll-up note labels;
- Excel sheet names, configured conservation ranges, and required labels;
- required and optional RiverWare slot names; and
- the list of accepted RiverWare volume-unit names.

Examples of items that are **not** controlled only by JSON include:

- how a RiverWare `.mdl` or `.mdl.gz` file is parsed;
- formulas and calculation mechanics implemented in `report_builder.py`;
- Word paragraph, table, row-height, alignment, numbering, or page-layout mechanics;
- Flask routes, upload handling, caching, or browser behavior;
- the Source drawer UI; and
- support for a completely new RiverWare unit conversion.

Changing a JSON rule changes behavior only if the Python application already knows how to interpret that rule.

---

# Part I - Using the Application

## 3. Windows installation and first launch

### 3.1 Prerequisites

Install Python 3.11 or newer. If installing Python from python.org, enable the option to add Python to PATH.

The application dependencies are listed in `requirements.txt`:

- Flask;
- openpyxl; and
- python-docx.

Development/testing also uses pytest through `requirements-dev.txt`.

### 3.2 Extract the application

Extract the application ZIP into a normal writable folder, for example:

```text
C:\WaterUse\WU_Assumptions_Generator_v2_1
```

Avoid running directly from inside the ZIP file.

### 3.3 Start the application

Double-click:

```text
run_app.bat
```

On the first launch, the batch file:

1. looks for the Windows `py` launcher or `python`;
2. creates a local `.venv` virtual environment if one does not exist;
3. installs the packages in `requirements.txt`; and
4. starts the local application.

The browser opens at:

```text
http://127.0.0.1:5051
```

Keep the command-window open while using the application. Closing that window stops the local server.

The server is bound to `127.0.0.1`, so it is intended to run locally on the user's workstation.

---

## 4. Monthly report workflow

### 4.1 Prepare the monthly Excel workbook

Before uploading the Projected State Use workbook:

1. open it in Microsoft Excel;
2. allow formulas to recalculate;
3. verify the expected year sheets are present;
4. save the workbook; and
5. close it if desired before uploading.

This matters because `openpyxl` reads cached formula results. It does not calculate Excel formulas itself.

If the workbook calculation mode is Manual, the application warns the user during preflight.

### 4.2 Prepare RiverWare model files

The application accepts:

```text
.mdl
.mdl.gz
```

Compressed `.mdl.gz` files are preferable for large models.

Use the correct model for each selected scenario. Filenames containing `MOST`, `MIN`, or `MAX` help the application detect an obvious scenario assignment error.

### 4.3 Select the Excel workbook

Under **Source files**, choose the current monthly Projected State Use `.xlsx` workbook.

A workbook must be supplied for every run. The application does not rely on a bundled default monthly workbook.

### 4.4 Select scenarios

The current default-enabled scenarios are:

- Most Probable; and
- Probable Minimum.

Probable Max is optional.

For Max, either:

1. turn on its **Include** switch and select the Max model; or
2. directly choose or drop a Max model into the Max model picker; selecting the file automatically enables Max.

Any deliberate combination can be generated, provided every enabled scenario has a model file.

### 4.5 Select the 24-Month Study month

Choose the study month using the month selector.

This value is used in the report subtitle and output filename. It does not change the model start year extracted from RiverWare.

### 4.6 Run preflight

Click **Run preflight**.

Preflight validates the uploaded sources before Word generation. It checks, among other things:

- model run dates;
- the model start year;
- common start year across a multi-scenario batch;
- required RiverWare slots;
- RiverWare annual timestep expectations;
- accepted volume units;
- scenario clues in model filenames;
- required Excel year sheets;
- scenario-specific Excel blocks;
- expected year-sheet labels;
- cached Excel values needed by the report;
- System Conservation year columns;
- the conservation summary structure; and
- availability of summary values when the selected scenario requires the summary table.

Generation remains disabled until preflight succeeds.

### 4.7 Review warnings carefully

A warning is not the same as a failed preflight. For example, if a model filename does not contain a recognizable scenario token, the application can continue but asks the user to verify the assignment.

A clear filename mismatch, such as assigning a filename identified as `MAX` to the Min scenario, is treated as an error.

### 4.8 Generate reports

After successful preflight, select **Generate report(s)**.

For one scenario, the application generates one Word report.

For multiple scenarios, it generates:

- one Word file per scenario; and
- a batch ZIP containing all generated Word files.

### 4.9 Review the browser preview

The preview is intentionally read-only. It and the Word report are generated from the same canonical `ReportContext`, reducing the risk that browser content and Word content diverge.

### 4.10 Use the Source feature

Select **Source** beside a traceable report value to see provenance such as:

- RiverWare model and full slot name;
- report year;
- original/converted value and units;
- Excel workbook, sheet, and cell;
- Excel label or note;
- calculated-value explanation; or
- policy rule and policy version.

Source/provenance information is a UI QA feature. It is intentionally not written into the generated Word report.

### 4.11 Download output

Download either:

- the individual scenario Word file; or
- the batch ZIP for a multi-scenario run.

The generated Word document does not intentionally include source filenames, policy version, generation timestamp, or source-trace records.

---

## 5. Current scenario behavior

Under policy `2026.09.1`:

| Scenario key | Display name | Label column | Value column | Notes column | Conservation summary | Powell subtitle |
|---|---|---:|---:|---:|---|---|
| `Most` | Most Probable | A | B | D | Yes | Yes |
| `Min` | Probable Minimum | F | G | I | No | No |
| `Max` | Probable Max | K | L | N | No | No |

The keys `Most`, `Min`, and `Max` are internal identifiers. Do not rename those keys casually. Changing a display name is much safer than changing an internal key.

---

# Part II - Editing Rules

## 6. Recommended rule-editing procedure

Use this process every time report policy changes.

### Step 1 - Finish or stop the current app session

Close the application browser tab if desired, then stop the Flask server by closing its command window or pressing `Ctrl+C` in that window.

The policy is loaded when the application starts. Editing the JSON while the app is already running does not change the policy already in memory.

### Step 2 - Back up the current policy file

Before editing:

```text
config\report_policy.json
```

make a backup copy, for example:

```text
config\report_policy.2026.09.1.backup.json
```

For a controlled production process, also archive the entire released application ZIP.

### Step 3 - Open the JSON in a plain-text editor

Recommended editors include Visual Studio Code or Notepad++. Windows Notepad also works.

Do not edit the JSON in Microsoft Word.

### Step 4 - Make the smallest practical change

Change only the rule needed for the new policy. Avoid mixing unrelated changes in one edit. This makes testing and rollback much easier.

### Step 5 - Increment `policy_version`

At the top of the file, update:

```json
"policy_version": "2026.09.1"
```

when the change can alter report content, source interpretation, scenario behavior, required inputs, or validation.

A suggested convention is:

```text
YYYY.MM.REVISION
```

For example:

```text
2026.10.0
2026.10.1
```

The exact numbering convention is an administrative choice, but the important point is that a behavior-changing policy edit should have a new version.

### Step 6 - Save the JSON

Save as plain UTF-8 text.

### Step 7 - Validate JSON syntax

From a Command Prompt in the project folder, run:

```bat
.venv\Scripts\python.exe -m json.tool config\report_policy.json > NUL
```

If the command returns without an error message, the JSON is syntactically valid.

If it reports an error, fix the JSON before starting the application.

### Step 8 - Run automated tests

Double-click:

```text
run_tests.bat
```

or run:

```bat
.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.venv\Scripts\python.exe -m pytest -q
```

A policy change may intentionally require a test expectation to be updated. Do not simply delete or weaken a failing regression test. First determine whether the failure correctly identifies a changed business rule.

### Step 9 - Restart the application

Run:

```text
run_app.bat
```

The restarted Flask process loads the edited policy.

### Step 10 - Confirm the active rule in the Rules panel

Click **Rules** in the application header.

Verify:

- the displayed policy version;
- the human-readable rule summary; and
- if necessary, **Show full active configuration (JSON)**.

The Rules drawer is the best way to confirm what the running process actually loaded.

### Step 11 - Run a representative preflight

Use a known workbook and representative RiverWare model(s). Confirm that preflight succeeds and that the expected scenario mappings, report years, and slot counts appear.

### Step 12 - Generate and QA a report

Review:

- headings and subtitles;
- annual totals;
- state sections;
- policy-controlled wording;
- conservation totals;
- ICS table;
- conservation summary table, when applicable;
- Notes and Disclaimers; and
- Source records for values affected by the changed rule.

Open the generated `.docx` in Microsoft Word and visually check pagination and formatting.

### Step 13 - Record the change

For a maintained release, update `CHANGELOG.md` with:

- the new policy version;
- what rule changed;
- why it changed;
- effective year/date if applicable; and
- what validation was performed.

Archive the prior policy/release so a report can be reproduced later if needed.

---

## 7. JSON syntax rules

JSON is strict. The most common editing errors are commas and quotation marks.

### 7.1 Strings require double quotes

Correct:

```json
"display_name": "Probable Minimum"
```

Incorrect:

```text
'display_name': 'Probable Minimum'
```

### 7.2 Items are separated by commas

Correct:

```json
{
  "effective_from_year": 2027,
  "display_kaf": 760
}
```

### 7.3 No trailing comma after the last item

Correct:

```json
[
  "Most",
  "Min"
]
```

Incorrect:

```text
[
  "Most",
  "Min",
]
```

### 7.4 JSON has no comments

Do not add lines beginning with `#` or `//` inside `report_policy.json`.

Use the existing `description` fields or maintenance documentation for explanations.

### 7.5 Booleans are lowercase

Correct:

```json
"decompose": false
```

Incorrect:

```text
"decompose": False
```

---

## 8. Policy file structure

The top-level policy has these major sections:

```text
policy_version
  description
  scenarios
  report
  excel
  riverware
```

The following sections explain each area.

---

## 9. `policy_version`

Example:

```json
"policy_version": "2026.09.1"
```

### Purpose

Identifies the report/configuration rules loaded by the app.

### Edit risk

**Low**, but it should accurately correspond to the actual rule set.

### Recommendation

Increment it whenever a change can affect generated content or source interpretation.

The version appears in the application and source tracing but is intentionally not printed into the generated Word report.

---

## 10. `description`

Example:

```json
"description": "Lower Basin Water Use assumptions report rules..."
```

### Purpose

Human-readable description shown in the Rules view.

### Edit risk

**Low.** It does not drive report calculations.

---

## 11. `scenarios`

Current keys:

```text
Most
Min
Max
```

Each scenario contains a display name and Excel columns.

Example:

```json
"Min": {
  "display_name": "Probable Minimum",
  "excel": {
    "label_column": "F",
    "value_column": "G",
    "notes_column": "I"
  }
}
```

### 11.1 `display_name`

Controls user-facing scenario text and output naming.

**Edit risk: Low to Medium.** A wording change is normally safe, but regression expectations and filenames may change.

### 11.2 Excel column mappings

For each scenario:

- `label_column` tells the reader where field labels are located;
- `value_column` tells the reader where scenario values are located; and
- `notes_column` tells the reader where notes/qualifiers are located.

**Edit risk: High.** These columns determine which values the report reads.

Only change them when the monthly workbook layout has actually changed. Verify the workbook visually and run all tests plus a representative report.

### 11.3 Do not casually rename scenario keys

Changing `"Min"` to another internal key is not equivalent to changing `display_name`. Other application code and tests expect the scenario keys.

**Recommendation:** change `display_name` when the visible wording changes; treat internal-key changes as a code change.

---

## 12. `report.powell_release_subtitle`

Current value:

```text
For the 6 maf & 7 maf Powell Release Scenarios
```

### Purpose

Defines the actual text of the second subtitle when a scenario is configured to show it.

### Edit risk

**Low.** This is report wording.

---

## 13. `report.powell_release_subtitle_scenarios`

Current value:

```json
[
  "Most"
]
```

### Purpose

Controls which scenarios show the Powell-release second subtitle.

Examples:

Show it only for Most:

```json
["Most"]
```

Show it for Most and Max:

```json
["Most", "Max"]
```

Show it for none:

```json
[]
```

### Edit risk

**Low to Medium.** It changes report content but not calculations.

---

## 14. `report.conservation_summary_scenarios`

Current value:

```json
[
  "Most"
]
```

### Purpose

Controls which scenarios include the state-level **Modeled Lower Basin Conservation Actions** summary table.

The detailed contractor-level conservation-actions table is intentionally not generated by this application.

### Edit risk

**Medium.** Enabling this for another scenario also makes preflight require usable summary values for that scenario/report period.

---

## 15. `report.notes_conservation_disclaimer`

### Purpose

Controls the conservation text printed under **Notes and Disclaimers**.

### Edit risk

**Low**, as long as only wording is being changed.

The previous SEIS ROD input/note is no longer part of the application.

---

## 16. `report.disclaimers`

Example structure:

```json
"disclaimers": [
  {
    "effective_from_year": 2027,
    "text": "per the 2027-2028 Operating Guidelines ..."
  },
  {
    "effective_from_year": 0,
    "text": "Based on Lake Mead Operating Condition ..."
  }
]
```

### Purpose

Controls the annual text following **Total projected water use**.

### How effective-year selection works

The application evaluates entries by `effective_from_year`, newest/highest year first. For a report year, it uses the first entry whose effective year is less than or equal to that report year.

That means:

- a 2027 rule applies to 2027 and later until superseded by a newer rule;
- a year `0` entry acts as a fallback for earlier years.

### Adding a future rule

For example, if wording changes beginning in 2029, add a new entry:

```json
{
  "effective_from_year": 2029,
  "text": "New approved wording beginning in 2029."
}
```

Keep the older rules too unless they are no longer needed for reproducibility/testing.

### Edit risk

**Medium.** This changes official report narrative and should be approved and regression-tested.

---

## 17. `report.default_enabled_scenarios`

Current value:

```json
[
  "Most",
  "Min"
]
```

### Purpose

Controls which scenario Include switches are on when the page first loads.

Max can still be selected even when it is not in this list.

### Edit risk

**Low.** This changes the initial UI state, not report math.

Use only valid internal scenario keys.

---

## 18. `report.rules.post_2026_start_year`

Current value:

```json
"post_2026_start_year": 2027
```

### Purpose

This is a broad transition-year switch used by report construction for pre-transition versus post-transition California/Arizona presentation and wording.

It affects more than a single sentence. For example, it participates in deciding which California ICS/reduction presentation and Arizona DCP/reduction phrasing are used.

### Edit risk

**High.** Change only when the underlying business/report regime changes and then run full regression and visual QA.

---

## 19. `report.rules.az_water_use_reduction_override`

Current structure:

```json
"az_water_use_reduction_override": {
  "effective_from_year": 2027,
  "display_kaf": 760
}
```

### Purpose

When the underlying CAP shortage slot is positive for an applicable year, this rule replaces the displayed shortage amount with the configured report-policy value and changes the phrase to a water-use reduction.

The Source trace records both:

- the underlying RiverWare shortage slot; and
- the policy override.

### Edit risk

**High.** This changes a displayed numeric value. Require a policy basis, updated version, tests, and representative report review.

---

## 20. `report.rules.mexico_shortage_wording`

Current behavior:

- before 2027: `Shortage volume`;
- from 2027: `Reduced delivery`.

### Purpose

Changes the label used for the RiverWare Mexico annual shortage/reduction value.

### Edit risk

**Medium.** Numeric source remains RiverWare; the policy changes the wording.

Use the same effective-year pattern described for `report.disclaimers`.

---

## 21. `report.rules.nevada_shortage_wording`

Current behavior:

- before 2027: `Shortage volume`;
- from 2027: `Water use reduction`.

### Edit risk

**Medium.** This changes wording associated with the Nevada shortage value.

---

## 22. `report.rules.nevada_system_conservation_wording`

Example:

```json
{
  "effective_from_year": 2027,
  "label": "Other water left in Mead is",
  "decompose": false
}
```

### Purpose

Controls both:

- the displayed Nevada system-conservation label; and
- whether that total is decomposed into Tributary conservation and Other system conservation sub-bullets.

### `decompose`

- `true` means the app may show the detailed decomposition when the corresponding RiverWare values are positive;
- `false` suppresses that decomposition and shows only the configured top-level presentation.

### Edit risk

**Medium to High.** This changes section structure as well as wording.

---

## 23. `report.rules.california_system_conservation_rollup`

Current structure:

```json
"california_system_conservation_rollup": {
  "year_sheet_note_labels": [
    "MWD System Conservation"
  ],
  "description": "..."
}
```

### Purpose

The application finds the California water-left-in-Mead entry on the appropriate year sheet. If its note contains one of the configured labels and the sign-adjusted conservation amount is positive, the application:

1. adds that amount to **Total California System Conservation**; and
2. suppresses the separate note-labeled California bullet.

If the note does not match, the water-left-in-Mead amount remains a separate report bullet.

### Note matching behavior

Matching is:

- case-insensitive;
- whitespace-normalized; and
- based on the configured text being contained within the workbook note.

For example, a configured label of:

```text
MWD System Conservation
```

can match a note containing that phrase with different capitalization or extra surrounding text.

### Adding another approved roll-up note

Example:

```json
"year_sheet_note_labels": [
  "MWD System Conservation",
  "Another approved California conservation note"
]
```

### Edit risk

**High.** This can change a numeric total and whether a separate bullet appears. Always validate the Source records and generated output after changing it.

---

# Part III - Excel Configuration

## 24. `excel.year_sheet_required_labels`

Current configured labels include strings used to confirm that expected year-sheet content exists.

### Matching behavior

The year-sheet reader performs normalized, case-insensitive substring matching. It also normalizes whitespace and common dash variants.

This allows a configured fragment to validate a longer workbook label.

### Edit risk

**Medium to High.** These labels are validation guards. Removing one can make preflight less protective; adding an incorrect one can block otherwise valid workbooks.

---

## 25. `excel.system_conservation_sheet`

Current value:

```text
System Conservation
```

### Purpose

Names the worksheet from which conservation blocks are read.

### Edit risk

**High** if the workbook layout has changed. The configured sheet must exist.

---

## 26. `excel.conservation_blocks`

Current blocks:

```text
cawcd
non_cawcd
california
```

Each has an Excel range such as:

```json
"range": "R4:Y20"
```

### How the range is used

The configured range establishes the expected block columns and approximate data-block height. Within that structure, the reader dynamically locates the `Contractor` header in the relevant contractor column rather than depending on one fixed header row. This gives some tolerance to rows inserted above the table.

The year columns are then read from the headers to the right of `Contractor`.

### Edit risk

**High.** A wrong range can read the wrong table or omit data. Update only after inspecting the new workbook layout and then test with the actual monthly workbook.

---

## 27. Conservation summary configuration

Relevant keys include:

```text
summary_sheet
summary_header_label
summary_expected_rows
```

The application searches the configured summary sheet for the header label and a sequence of year columns beginning with the model start year. It validates expected rows such as AZ, CA, NV, Annual Total, and Cumulative Total.

### Edit risk

**High.** These settings control where the state-level conservation summary is obtained and how its structure is validated.

---

# Part IV - RiverWare Configuration

## 28. `riverware.required_slots`

This is the contract between a compatible RiverWare model and the report.

Each entry is a full RiverWare object/slot name, for example:

```text
AnnualWaterUse.CaTotalAnnual
```

### Purpose

Preflight expects every required slot to be successfully parsed with usable annual data and acceptable volume units.

### Edit risk

**Very High.** Do not remove or rename required slots simply to make preflight pass. A report calculation may rely on the slot even if missing values are sometimes treated as zero/blank in downstream logic.

When RiverWare changes:

1. confirm the actual new object/slot name;
2. determine where the old slot was used in the report calculation;
3. update policy and Python logic if necessary;
4. run unit/regression tests; and
5. run the real-model integration test.

---

## 29. `riverware.optional_slots`

The current policy has no optional slots:

```json
"optional_slots": []
```

### Purpose

Allows the parser to attempt a slot without failing preflight solely because the slot is absent.

### Edit risk

**Very High.** A slot should be optional only when the report builder safely and intentionally handles its absence.

A slot cannot be listed in both `required_slots` and `optional_slots`; the policy loader rejects that configuration.

---

## 30. `riverware.accepted_volume_units`

Current policy includes accepted representations for:

- cubic meters;
- acre-feet; and
- cubic feet.

### Important limitation

The policy list is an acceptance gate. It does **not** define conversion mathematics.

The current parser's conversion code supports known acre-foot, cubic-meter, and cubic-foot forms. Adding a completely new unit name to JSON does not teach the application how to convert that unit.

### Edit risk

**Very High.** Adding a genuinely new unit type requires a code change in `riverware_annual.py` plus tests.

---

## 31. `riverware.target_year_count`

Current value:

```json
"target_year_count": 3
```

### Important limitation

The current policy loader explicitly requires this value to remain `3`.

Changing it to another number causes policy loading to fail.

The application can still produce a two-year report if the third Excel year sheet is absent, but the RiverWare extraction contract remains three target years.

### Edit risk

**Not currently configurable beyond 3.** Supporting another target-year count requires a code/design change and new tests.

---

# Part V - Practical Rule-Change Examples

## 32. Example: change only visible scenario wording

Suppose the approved visible label changes from:

```text
Probable Max
```

to:

```text
Probable Maximum
```

Change only:

```json
"Max": {
  "display_name": "Probable Maximum",
  ...
}
```

Then increment `policy_version`, validate JSON, run tests, restart, and inspect the Rules panel/output filename/subtitle.

Do not rename the internal `Max` key.

---

## 33. Example: show the conservation summary in Max reports

Change:

```json
"conservation_summary_scenarios": [
  "Most"
]
```

to:

```json
"conservation_summary_scenarios": [
  "Most",
  "Max"
]
```

Preflight will then expect the summary data to be usable for Max when Max is selected.

---

## 34. Example: add new effective-year disclaimer wording

If new approved wording begins in 2029, append a new object to `report.disclaimers`:

```json
{
  "effective_from_year": 2029,
  "text": "Approved wording for 2029 and later."
}
```

Leave the 2027 and fallback entries intact unless there is an intentional reason to remove historical behavior.

---

## 35. Example: update the Excel scenario block after workbook redesign

If the Min block moves, update all three column references together:

```json
"Min": {
  "display_name": "Probable Minimum",
  "excel": {
    "label_column": "NEW_LABEL_COLUMN",
    "value_column": "NEW_VALUE_COLUMN",
    "notes_column": "NEW_NOTES_COLUMN"
  }
}
```

Replace the placeholders with actual Excel column letters.

Do not change only the value column without confirming that labels and notes still align row-for-row with that same scenario block.

This is a structural change and should be tested with the new workbook before release.

---

## 36. Example: recognize another California roll-up note

If an approved monthly workbook begins using another note phrase for the same conservation concept, extend the list:

```json
"year_sheet_note_labels": [
  "MWD System Conservation",
  "New approved MWD conservation label"
]
```

Then verify a representative California section using **Source** to make sure:

- the correct Excel entry was matched;
- the amount was sign-adjusted correctly;
- the amount was included once in Total California System Conservation; and
- the separate bullet was suppressed only when intended.

---

# Part VI - What Requires a Code Change Instead of a JSON Edit

## 37. Calculation changes

If a business requirement changes the actual formula for a reported value, determine whether the existing policy structure can express it.

Examples that generally require code changes include:

- summing an entirely new combination of RiverWare slots;
- changing the arithmetic definition of basin/state totals;
- adding a new conditional calculation not represented by current policy fields;
- adding a new state/report section; or
- changing how missing data is mathematically treated.

These belong primarily in `report_builder.py`, with corresponding policy additions only when useful.

---

## 38. RiverWare parser changes

Changes to model serialization, slot encoding, timestep parsing, gzip handling, or unit conversion belong in:

```text
riverware_annual.py
```

Do not try to compensate for a parser incompatibility by simply deleting required slots from JSON.

---

## 39. Word-formatting changes

Formatting mechanics belong in:

```text
docx_writer.py
reference_report.docx
```

Examples include:

- row height;
- cell alignment;
- Word styles;
- page setup;
- paragraph spacing;
- bullet/numbering behavior; and
- page-break logic.

Current ICS and conservation-summary table rows use exact 0.25-inch height and centered cell alignment. There is no forced page break between those two tables.

---

## 40. UI changes

Browser interaction and layout are implemented in:

```text
templates/index.html
static/app.js
static/app.css
```

The report preview renderer is:

```text
templates/report_preview.html
```

The Rules drawer is read-only by design; policy edits are made in the JSON file, not in the browser.

---

# Part VII - Application Architecture

## 41. Main files and responsibilities

| File | Responsibility |
|---|---|
| `app.py` | Flask routes, upload/generation endpoints, local HTTP behavior |
| `launch_app.py` | Windows-friendly local launch and browser opening |
| `run_app.bat` | Creates/reuses `.venv`, installs dependencies, starts app |
| `policy.py` | Loads and validates `report_policy.json` and exposes policy helpers |
| `config/report_policy.json` | Versioned report/business/source configuration |
| `report_service.py` | Batch preflight, scenario validation, input cache |
| `riverware_annual.py` | `.mdl`/`.mdl.gz` parsing and volume conversion |
| `excel_source.py` | Excel scenario/year/conservation/summary reading and validation |
| `report_models.py` | Canonical report and source-trace data structures |
| `report_builder.py` | Report calculations, business narrative, section construction |
| `docx_writer.py` | Word rendering using `reference_report.docx` |
| `reference_report.docx` | Authoritative Word style/page/numbering template |
| `templates/report_preview.html` | Read-only HTML report renderer |
| `templates/index.html` | Main application page |
| `static/app.js` | Browser workflow, batch UI, Rules/Source drawers, downloads |
| `static/app.css` | Browser styling |
| `tests/` | Regression and integration tests |
| `CHANGELOG.md` | Release/change history |

---

## 42. Data flow

At a high level:

```text
RiverWare model(s) -----> riverware_annual.py ---+
                                                |
Excel workbook --------> excel_source.py --------+--> ReportContext
                                                |       |
report_policy.json -----> policy.py -------------+       +--> HTML preview
                                                        |
                                                        +--> Word report
```

The important design rule is that HTML and Word should not independently calculate report values. `report_builder.py` constructs one canonical report model, and both renderers use it.

---

# Part VIII - Testing and Validation

## 43. Standard automated tests

After the app has created `.venv`, run:

```text
run_tests.bat
```

or:

```bat
.venv\Scripts\python.exe -m pytest -q
```

The regression suite covers areas such as:

- scenario Excel blocks;
- policy effective-year behavior;
- Arizona total source logic;
- California conservation roll-up behavior;
- conservation-summary visibility;
- subtitle behavior;
- Notes and Disclaimers behavior;
- canonical report generation;
- Word table row height/alignment/page flow; and
- absence of source/provenance metadata from Word output.

---

## 44. Real RiverWare model integration test

To test a representative real or sanitized model without storing it in the repository:

```bat
set WU_TEST_MODEL=C:\path\to\model.mdl.gz
.venv\Scripts\python.exe -m pytest -q -m integration
```

Use this before accepting a significant RiverWare version change, slot change, or parser modification.

---

## 45. Recommended test matrix after significant changes

For a material rule/source/code change, test as many of these as applicable:

| Test case | Why |
|---|---|
| Most only | Baseline and conservation-summary behavior |
| Min only | Min Excel block and Probable Minimum presentation |
| Max only | Optional Max selection and Max Excel block |
| Most + Min | Common routine batch |
| Most + Min + Max | Full batch behavior |
| `.mdl` | Uncompressed model path |
| `.mdl.gz` | Normal compressed model path |
| workbook with third year | Three-year report |
| workbook without third-year sheet | Two-year report behavior |
| known model filename mismatch | Scenario validation |
| recalculated workbook | Normal cached-formula path |
| representative historical run | Regression/parity review |

---

# Part IX - Maintaining the Excel Interface

## 46. When the monthly workbook changes only in values

No policy change should be necessary if:

- sheet names remain the same;
- scenario blocks remain in the same columns;
- required labels remain recognizable;
- conservation blocks remain structurally compatible; and
- summary table structure remains compatible.

Simply recalculate/save the new workbook and use it for the monthly run.

---

## 47. When the workbook layout changes

Treat layout changes as a controlled maintenance event.

Review:

1. scenario label/value/note columns;
2. year-sheet required labels;
3. System Conservation sheet name;
4. conservation block columns/ranges;
5. conservation summary sheet name;
6. summary header label;
7. expected summary rows; and
8. formula-cache availability.

Update `report_policy.json` only for layout changes already supported by the existing Excel reader. If the workbook adopts a fundamentally new structure, `excel_source.py` may also need a code change.

---

# Part X - Maintaining RiverWare Compatibility

## 48. When a RiverWare upgrade is introduced

Before production use:

1. preserve a known-good older model/report pair;
2. test a model created by the new RiverWare version;
3. run preflight and note missing/unreadable slots;
4. run the integration test;
5. compare major report values with trusted sources;
6. inspect Source details for representative values; and
7. generate and visually review Word output.

If slot names changed, update the policy only after confirming the report's business meaning is still the same. If serialization changed, modify/test `riverware_annual.py`.

---

# Part XI - Maintaining the Word Template

## 49. `reference_report.docx` is authoritative

The application uses:

```text
reference_report.docx
```

as the authoritative Word formatting source.

The writer preserves styles, numbering definitions, section/page setup, and other template structures while replacing body content with the generated report.

### After replacing or editing the template

Always:

1. run the automated tests;
2. generate a representative Most report;
3. generate a representative Min report;
4. open the outputs in Microsoft Word;
5. inspect every page; and
6. verify headings, bullets, state-value styling, highlighted disclaimer text, table widths/heights/alignment, margins, and page breaks.

Do not assume that a DOCX that opens successfully is visually correct.

---

# Part XII - Troubleshooting

## 50. `Python 3 was not found`

Install Python 3.11 or newer and make sure either the `py` launcher or `python` is available from Command Prompt.

Then rerun `run_app.bat`.

---

## 51. First launch fails while installing packages

Check:

- Internet/network access needed for the initial package installation;
- corporate proxy/security restrictions;
- Python installation health; and
- write permission to the application folder.

If a partially created `.venv` is unusable, it can be removed and recreated by rerunning `run_app.bat` after the underlying issue is fixed.

---

## 52. Rules change does not appear

Most likely cause: the app was not restarted.

Procedure:

1. stop the Flask command window;
2. run `run_app.bat` again;
3. click **Rules**; and
4. confirm the new policy version/value.

---

## 53. Application fails immediately after editing JSON

Common causes:

- missing comma;
- extra trailing comma;
- unmatched brace/bracket;
- single quotes instead of double quotes;
- missing `policy_version`;
- missing scenario fields;
- empty required-slot list;
- duplicate required slots;
- a slot listed as both required and optional;
- no accepted volume units; or
- `target_year_count` changed from 3.

Validate syntax with:

```bat
.venv\Scripts\python.exe -m json.tool config\report_policy.json > NUL
```

If syntax is valid but startup still fails, restore the prior policy backup and compare the changed section.

---

## 54. Preflight says a RiverWare slot is missing

Do not immediately remove it from `required_slots`.

First determine whether:

- the wrong scenario model was uploaded;
- the RiverWare model is incomplete;
- the object/slot was renamed in a newer model version;
- the parser no longer recognizes the serialized slot form; or
- the report no longer legitimately requires that value.

Only the last case justifies changing required/optional status without a replacement mapping.

---

## 55. Preflight reports unsupported or missing units

The application requires explicit supported volume units for configured report slots.

Confirm the model slot's unit in RiverWare. Do not simply add an arbitrary unit string to `accepted_volume_units` unless the parser's conversion code already supports it.

---

## 56. Excel cached value is missing

Open the workbook in Excel, recalculate it, save it, then rerun preflight.

This is especially important for workbooks with formulas or dynamic arrays.

---

## 57. Workbook calculation mode is Manual

Preflight warns about Manual calculation mode because cached formula values may be stale.

Set/recalculate as appropriate in Excel and save the workbook before report generation.

---

## 58. Wrong values appear for Min or Max

First check the active Rules panel and verify the Excel mapping:

```text
Most: A / B / D
Min:  F / G / I
Max:  K / L / N
```

Then use **Source** on the questionable value to confirm the actual Excel cell or RiverWare slot used.

If the workbook format changed, update the scenario mapping only after verifying the new columns.

---

## 59. Preflight session expired

Prepared inputs are cached in memory for 30 minutes.

If the cache expires or the app is restarted, rerun preflight before generating.

The cache is process-local and is not a persistent database.

---

## 60. Upload is too large

The Flask application limits the combined request to approximately 1.2 GB.

The RiverWare gzip decoder also has a 512 MB decompressed-model safety limit per model.

Prefer `.mdl.gz` and, if necessary, process fewer scenarios in one batch.

---

# Part XIII - Release and Change-Control Practice

## 61. Suggested categories of changes

### Category A - Text-only policy change

Examples:

- disclaimer wording;
- visible scenario name;
- Powell subtitle text.

Minimum QA:

- increment policy version;
- validate JSON;
- run tests;
- inspect Rules panel;
- generate representative report;
- visually inspect affected text.

### Category B - Source mapping or numeric policy change

Examples:

- Excel columns;
- conservation ranges;
- Arizona numeric override;
- California roll-up labels;
- RiverWare slot requirement.

Minimum QA:

- all Category A checks;
- representative source-data comparison;
- Source trace review;
- multi-scenario test if applicable;
- real-model integration test for RiverWare changes.

### Category C - Code/parser/report-calculation change

Minimum QA:

- code review;
- new/updated regression tests;
- all standard tests;
- real-model integration test where relevant;
- representative Most/Min/Max generation;
- Word visual QA; and
- update changelog/documentation.

### Category D - Word template change

Minimum QA:

- automated tests;
- Most and Min generation;
- full-page visual inspection in Word/PDF;
- verify numbering, styles, tables, and page flow.

---

## 62. Recommended release checklist

Before distributing an updated application:

- [ ] `policy_version` reflects the active rule set.
- [ ] `CHANGELOG.md` describes the change.
- [ ] JSON syntax validation passes.
- [ ] Automated tests pass.
- [ ] Real-model integration test passes if RiverWare behavior changed.
- [ ] Most report generated and reviewed.
- [ ] Min report generated and reviewed.
- [ ] Max report generated/reviewed if the change affects Max or scenario logic.
- [ ] Full batch tested if batch logic changed.
- [ ] Rules panel displays the expected active values.
- [ ] Source trace checked for changed calculations/mappings.
- [ ] Word visual formatting checked.
- [ ] Monthly source workbook/model files are not unintentionally bundled in the release.
- [ ] `.venv`, caches, temporary QA files, and generated reports are excluded from the distributable.
- [ ] Previous release/policy is archived for rollback.

---

# Part XIV - Rollback and Reproducibility

## 63. Rolling back a bad rule change

If a new rule produces incorrect output:

1. stop the application;
2. restore the prior known-good `report_policy.json`;
3. restart the application;
4. confirm the prior policy version in **Rules**;
5. rerun preflight; and
6. regenerate/verify the report.

Do not rely on a browser tab that was opened before the rollback; confirm the running version in Rules.

---

## 64. Reproducing an older report

For strong reproducibility, retain together:

- the exact application release;
- the exact `report_policy.json` version;
- the monthly Projected State Use workbook;
- the scenario RiverWare model(s); and
- the trusted generated/final report when appropriate.

The generated Word report intentionally does not embed the policy version or source filenames, so reproducibility depends on operational archive practices outside the DOCX.

---

# Part XV - Important Current Design Constraints

## 65. Three-year RiverWare extraction contract

The policy currently requires `target_year_count = 3`. The third Excel year sheet may be absent, in which case report output can be limited to two years, but changing the RiverWare target-year count itself requires code work.

## 66. Policy is loaded at startup

The app does not hot-reload JSON rule changes. Restart after every policy edit.

## 67. Rules UI is read-only

This is intentional. It prevents accidental operational policy changes while a user is simply generating a report.

## 68. Excel formulas are not recalculated by Python

Use Excel to recalculate/save the monthly workbook before production use.

## 69. Source tracing is not embedded in Word

Source details are available in the browser for QA but intentionally excluded from the generated DOCX.

## 70. The application is local-only by design

Normal launch binds Flask to `127.0.0.1:5051` with debug mode off.

---

# Part XVI - Recommended Governance for Rule Editing

## 71. Who should edit rules?

A rule file is easy to edit technically, but some entries affect official numeric reporting. A sensible operational practice is to distinguish:

- **text administrators** - approved wording/display changes;
- **report/data maintainers** - Excel mappings and source layout;
- **model/report developers** - RiverWare slots, numeric overrides, calculation logic, parser changes; and
- **release reviewers** - regression/Word QA before distribution.

This is a process recommendation, not an application permission system. The current JSON file has no built-in role-based access control.

---

## 72. Rule-editing risk summary

| Policy area | Typical risk | JSON-only change usually sufficient? |
|---|---|---|
| `policy_version` | Low | Yes |
| `description` | Low | Yes |
| scenario `display_name` | Low/Medium | Yes |
| default-enabled scenarios | Low | Yes |
| Powell subtitle text/visibility | Low/Medium | Yes |
| conservation disclaimer wording | Low | Yes |
| effective-year disclaimer wording | Medium | Yes |
| Mexico/Nevada wording | Medium | Yes |
| Nevada `decompose` | Medium/High | Yes, within existing behavior |
| Arizona numeric display override | High | Yes, within existing behavior |
| California roll-up labels | High | Yes, within existing behavior |
| scenario Excel columns | High | Yes, if workbook structure remains compatible |
| Excel sheet/range/header settings | High | Yes, if structure remains compatible |
| RiverWare required/optional slots | Very High | Sometimes; verify report-builder dependency |
| accepted unit aliases | Very High | Only for units the parser already converts |
| new unit conversion | Very High | No - code change required |
| `target_year_count` other than 3 | Not supported | No - code change required |
| report arithmetic/formulas | Very High | Usually no - code change required |
| Word row height/alignment/layout | Medium/High | No - code/template change |
| Flask/UI workflow | Medium/High | No - code change |

---

# Part XVII - Short Rule-Edit Checklist

For experienced maintainers, the abbreviated process is:

1. stop the app;
2. back up `config/report_policy.json`;
3. edit the smallest necessary rule;
4. increment `policy_version`;
5. validate JSON with `python -m json.tool`;
6. run `run_tests.bat`;
7. restart with `run_app.bat`;
8. click **Rules** and verify the active configuration;
9. preflight known inputs;
10. generate and inspect the affected report(s);
11. use **Source** to validate affected values; and
12. record/archive the release change.

---

## 73. Final maintenance principle

Use JSON for **declarative report policy**: what a scenario is called, where a supported source is located, when approved wording changes, what existing section is visible, or what existing policy override applies.

Use Python/template code for **application mechanics**: how data is parsed, calculated, converted, rendered, validated, cached, or displayed.

Keeping that boundary intact makes the application easier to audit, test, and maintain over time.

---

## Developer companion for calculations and new RiverWare slots

For step-by-step developer instructions covering calculation changes, adding new RiverWare slots, adding bullets to existing report sections, preserving Source tracing, and writing regression tests, see:

```text
DEVELOPER_CALCULATION_SLOT_CHANGE_GUIDE.md
```
