# WU Assumptions Generator
## Developer Guide for Calculation Changes, New RiverWare Slots, and New Report Bullets

**Applies to:** WU Assumptions Generator v2.1 / policy 2026.09.1 architecture  
**Audience:** Maintainers who need to change report calculations, add RiverWare inputs, or add report content  
**Primary platform:** Windows  

---

## 1. Purpose of this guide

This guide explains how to make three common kinds of technical changes safely:

1. change a calculation that already uses existing RiverWare or Excel inputs;
2. add a new RiverWare slot and use its value in an existing report section; and
3. add a calculated bullet that combines two or more source values.

The application deliberately separates **policy/configuration** from **calculation and rendering code**. That means `config/report_policy.json` is the correct place to register RiverWare slots and store configurable business-rule parameters, but the JSON file is **not a formula engine**. A new formula or a new bullet normally requires a change in `report_builder.py` as well.

This document is a developer/change-management companion to `APP_USER_RULES_MAINTENANCE_GUIDE.md`.

---

# Part I - Understand the Data Flow Before Editing

## 2. How a RiverWare value reaches the report

For a normal RiverWare-based report value, the path is:

```text
config/report_policy.json
        |
        | required_slots / optional_slots
        v
policy.py
        |
        | policy.all_slots
        v
report_service.py
        |
        | calls load_annual_slots(...)
        v
riverware_annual.py
        |
        | parses .mdl / .mdl.gz
        | validates annual timestep and units
        | converts volume to acre-ft
        v
AnnualExtractionResult
        |
        v
report_builder.py
        |
        | retrieves slot value with _slot(...)
        | performs report calculations
        | creates Bullet / StateSection / table objects
        | registers Source references
        v
ReportContext
       / \
      /   \
     v     v
HTML     Word
preview  document
```

The key architectural rule is:

> **Do calculations once in `report_builder.py`. Do not independently recreate the same calculation in the HTML template or Word writer.**

Both the HTML preview and Word output consume the same `ReportContext`.

---

## 3. Which file should I change?

Use this table as a decision guide.

| Change | Primary file(s) |
|---|---|
| Rename a scenario, change an existing wording rule, change an existing effective year | `config/report_policy.json` |
| Add a RiverWare slot to be extracted | `config/report_policy.json` |
| Change a mathematical formula | `report_builder.py` |
| Add a new bullet to California | `report_builder.py` -> `_build_california()` |
| Add a new bullet to Arizona | `report_builder.py` -> `_build_arizona()` |
| Add a new bullet to Nevada | `report_builder.py` -> `_build_nevada()` |
| Add a new bullet to Mexico | `report_builder.py` -> `_build_mexico()` |
| Change state/basin totals | `report_builder.py` -> `_build_year_section()` |
| Change ICS-table arithmetic | `report_builder.py` -> `_build_ics_table()` |
| Change conservation summary source logic | `excel_source.py` and/or `report_builder.py` |
| Add a new configurable rule parameter | `config/report_policy.json`, usually `policy.py`, then `report_builder.py` |
| Change how RiverWare model text is parsed | `riverware_annual.py` |
| Add support for a new physical unit conversion | `riverware_annual.py`, policy, tests |
| Change Word formatting | `docx_writer.py` and/or `reference_report.docx` |
| Add an entirely new report section or new table type | `report_models.py`, `report_builder.py`, `templates/report_preview.html`, `docx_writer.py`, tests |
| Add a normal bullet to an existing state section | Usually only policy + `report_builder.py` + tests |

A normal new bullet does **not** require a change to `templates/report_preview.html` or `docx_writer.py`. Both renderers already loop through the `Bullet` objects supplied by the canonical report model.

---

# Part II - Before Any Calculation or Slot Change

## 4. Make a controlled working copy

Do not edit the only operational copy of the application.

Recommended process:

1. Stop the Flask application.
2. Copy the entire project folder to a development folder.
3. Keep the last known-good ZIP or project folder unchanged.
4. Make your changes only in the development copy.
5. Do not modify `reference_report.docx` unless the requested change is specifically a Word-formatting change.

If Git is available, create a branch before changing code. Git is strongly recommended but is not required by the application.

---

## 5. Write the business rule in plain language first

Before editing Python, write down the requested rule in one sentence.

Examples:

- "Display Slot X as a California bullet when its annual value is greater than zero."
- "New California total = Slot A + Slot B."
- "Beginning in 2027, display Slot C minus Slot D as a Nevada reduction."
- "The new bullet applies to Most and Minimum but not Maximum."

Also answer these questions:

1. What is the exact source: RiverWare, Excel, policy constant, or a calculation?
2. Is the source required for a valid report, or optional?
3. What units should be displayed: AF, kaf, or maf?
4. What rounding is required?
5. Should zero values be shown or suppressed?
6. Should negative values be displayed as negative, converted to positive, or interpreted as a reduction?
7. Does the rule vary by year?
8. Does the rule vary by scenario?
9. Where should the bullet appear relative to existing bullets?
10. What should the **Source** panel show for QA?

Do not begin coding until these are clear. Most report defects occur because one of these assumptions was implicit.

---

# Part III - Adding a New RiverWare Slot

## 6. Determine the exact RiverWare slot name

The application expects the fully qualified RiverWare name in the form used by the model, for example:

```text
ObjectName.SlotName
```

Existing examples include:

```text
ForecastUse.MWDDiversionAnnualFC
AnnualWaterUse.AzTotalAnnual
ICS Credits.Bank_CA
Mexico Shortage and Surplus.MexicoAdjustedSched
```

Use the exact object and slot naming from RiverWare. Spaces and capitalization should be preserved.

For an aggregate series member, the parser can also resolve a configured member path using the aggregate slot/member naming encoded in the RiverWare model.

### Verify these properties before adding the slot

For a normal time-series slot used by this report, confirm that:

- it exists in all model scenarios where the report needs it;
- it contains values for the model start year and the following two calendar years;
- it is an annual series with a **1 YEAR** timestep;
- it has an explicit supported volume unit; and
- its meaning is stable across the RiverWare model versions you expect to use.

The current parser supports the normal annual series forms used by this project, scalar slots, and supported aggregate-series members. A table/periodic/table-series source or a nonannual series may require a change to `riverware_annual.py` rather than only adding the name to JSON.

---

## 7. Decide whether the new slot is required or optional

Open:

```text
config/report_policy.json
```

The RiverWare section contains:

```json
"riverware": {
  "required_slots": [
    "..."
  ],
  "optional_slots": [],
  "accepted_volume_units": [
    "m3",
    "acre-ft",
    "ac-ft",
    "acreft",
    "ft3"
  ],
  "target_year_count": 3
}
```

### Use `required_slots` when

The report cannot be considered valid without the value.

If a required slot is missing, has an unacceptable unit, has the wrong timestep, or lacks a target-year value, preflight fails and report generation is blocked.

### Use `optional_slots` when

The value is genuinely optional and the report is still valid if it does not exist.

If an optional slot cannot be extracted, the application records a warning rather than failing preflight.

### Do not put the same slot in both lists

`policy.py` explicitly rejects overlap between `required_slots` and `optional_slots`.

### Recommended default

If the new bullet is an official required assumption, make its source slot **required**. Use optional only when absence has an approved business meaning such as "do not display this optional activity."

---

## 8. Add the slot to `report_policy.json`

Suppose the new RiverWare slot is hypothetically:

```text
ExampleObject.ExampleAnnualVolume
```

This is a placeholder example, not a slot currently used by the application.

To make it required, add it once to `required_slots`:

```json
"required_slots": [
  "AnnualWaterUse.California_Apportionment",
  "AnnualWaterUse.CaTotalAnnual",
  "ExampleObject.ExampleAnnualVolume"
]
```

Or, if genuinely optional:

```json
"optional_slots": [
  "ExampleObject.ExampleAnnualVolume"
]
```

JSON rules to remember:

- every string is enclosed in double quotes;
- every item except the last item in an array needs a comma;
- do not leave comments inside JSON;
- do not duplicate a slot name; and
- do not put a slot in both required and optional lists.

---

## 9. Increment the policy version

Even when the calculation code also changes, the active source contract has changed when a new slot is registered.

Update:

```json
"policy_version": "2026.09.1"
```

to the next locally approved version, for example:

```json
"policy_version": "2026.09.2"
```

Use your organization's versioning convention consistently.

The policy version is visible in the application but is intentionally not embedded in generated Word reports.

---

## 10. Validate the JSON before starting the app

From Windows Command Prompt in the project folder:

```bat
.venv\Scripts\python.exe -m json.tool config\report_policy.json > NUL
```

If the command returns without an error, the JSON syntax is valid.

This checks syntax only. It does **not** prove that the RiverWare slot exists or that the business rule is correct.

---

## 11. Understand what happens automatically after the slot is registered

You normally do **not** add a slot list to `riverware_annual.py`.

The current flow is already generic:

1. `policy.py` loads `required_slots` and `optional_slots`.
2. `ReportPolicy.all_slots` combines them.
3. `report_service.py` passes `policy.all_slots` to `load_annual_slots()`.
4. `riverware_annual.py` parses those names from each uploaded model.
5. Values are converted to acre-feet before they reach `report_builder.py`.

Therefore, for a normal supported annual volume slot, registration in JSON is enough to make the extracted value available to report logic.

---

# Part IV - Using the New Slot in an Existing Report Section

## 12. Choose the correct report-builder function

Open:

```text
report_builder.py
```

Use the section that corresponds to where the bullet belongs:

```text
_build_california()  -> California bullets
_build_arizona()     -> Arizona bullets
_build_nevada()      -> Nevada bullets
_build_mexico()      -> Mexico bullets
_build_year_section()-> state headings, U.S. total, basin total, Mexico heading
_build_ics_table()   -> ICS table
_build_conservation_summary() -> conservation summary table
```

The order of `section.bullets.append(...)` calls is the order in which bullets appear in both HTML and Word.

---

## 13. Retrieve the slot value

For the hypothetical slot:

```python
example_slot = "ExampleObject.ExampleAnnualVolume"
example_af = _slot(annual, example_slot, year)
```

`_slot(...)` returns the extracted value in **acre-feet**, or `None` if there is no usable value.

### Important: distinguish missing from zero

For new code, prefer:

```python
example_af = _slot(annual, example_slot, year)
```

rather than immediately writing:

```python
example_af = _slot(annual, example_slot, year) or 0.0
```

The second form makes a missing optional value indistinguishable from a legitimate zero. Existing code uses `or 0.0` in several established rules, but new logic should preserve `None` when missing-versus-zero matters.

---

## 14. Convert only for display

RiverWare volume values reaching `report_builder.py` are already in acre-feet.

Typical display conversions are:

```text
AF  -> value unchanged
kaf -> value / 1,000
maf -> value / 1,000,000
```

The current formatting helpers are:

```python
_fmt_int(value)
_fmt_dec(value, digits)
```

Examples:

```python
_fmt_int(example_af)                  # AF, integer display
_fmt_dec(example_af / 1000, 1)       # kaf, one decimal
_fmt_dec(example_af / 1_000_000, 3)  # maf, three decimals
```

Keep the mathematical value separate from its formatted text. Do not perform subsequent arithmetic on a formatted string.

---

## 15. Add a simple new bullet

A typical existing-section bullet pattern is:

```python
example_slot = "ExampleObject.ExampleAnnualVolume"
example_af = _slot(annual, example_slot, year)

if example_af is not None and example_af > 0:
    section.bullets.append(
        Bullet(
            f"Example activity of {_fmt_dec(example_af / 1000, 1)} kaf",
            source_id=_register(
                registry,
                [_rw_ref(annual, example_slot, year)],
            ),
        )
    )
```

This single block does four jobs:

1. gets the source value;
2. applies the display condition;
3. formats the report sentence; and
4. connects the bullet to the **Source** panel.

### Bullet hierarchy

A normal top-level bullet uses the default:

```python
level=0
```

A nested bullet uses:

```python
level=1
```

For example:

```python
Bullet(
    "Example nested detail ...",
    level=1,
    source_id=...
)
```

Do not add HTML list markup or Word bullet formatting manually. The renderers handle bullet presentation from the `Bullet.level` value.

---

## 16. Why no HTML or Word change is normally needed

The HTML preview already loops over every state bullet:

```text
state.bullets
```

and displays the Source control when `bullet.source_id` exists.

The Word writer also loops over every state bullet and calls the common bullet-formatting routine.

Therefore, adding a `Bullet(...)` to an existing `StateSection` automatically reaches both outputs.

You only need renderer/model changes if you are introducing a **new type of content**, such as a new table, a new independent section, a special multi-column block, or a new formatting concept.

---

# Part V - Source Tracing for New Values

## 17. Always add Source tracing

The browser's **Source** feature is part of the application's QA design. A new report value should normally be traceable.

For a direct RiverWare value, use:

```python
_rw_ref(annual, slot_name, year)
```

and register it:

```python
source_id=_register(
    registry,
    [_rw_ref(annual, slot_name, year)],
)
```

The Source panel can then show the RiverWare model source, fully qualified slot, year, extracted acre-foot value, and original model unit information.

---

## 18. Source tracing for calculations

If a displayed value is calculated from multiple inputs, register both the formula explanation and the underlying inputs.

Example:

```python
slot_a = "ExampleObject.AnnualA"
slot_b = "ExampleObject.AnnualB"

a_af = _slot(annual, slot_a, year)
b_af = _slot(annual, slot_b, year)

if a_af is not None and b_af is not None:
    combined_af = a_af + b_af
    display_kaf = _fmt_dec(combined_af / 1000, 1)

    source_id = _register(
        registry,
        [
            _calc_ref(
                "AnnualA + AnnualB",
                display_kaf,
                "kaf",
                "Approved combined activity calculation.",
            ),
            _rw_ref(annual, slot_a, year),
            _rw_ref(annual, slot_b, year),
        ],
    )

    section.bullets.append(
        Bullet(
            f"Combined activity of {display_kaf} kaf",
            source_id=source_id,
        )
    )
```

The calculation reference explains *what the app did*, while the RiverWare references show *which model inputs were used*.

This is the preferred pattern for calculated bullets.

---

## 19. Source tracing for policy values

If part of the result comes from a policy constant or override, include `_policy_ref(...)`.

Pattern:

```python
_policy_ref(
    policy,
    "report.rules.some_rule",
    "760 kaf",
    "Effective-year approved display override.",
)
```

This makes it clear in the Source panel that the value is not simply a raw model value.

---

## 20. Source tracing for Excel values

The current report builder has specialized helpers for Excel data already loaded by `excel_source.py`:

```text
_excel_ref(...)        -> a year-sheet ExcelValue
_conservation_ref(...) -> a contractor conservation row/year
_summary_ref(...)      -> conservation summary table row/year or total
```

If a new calculation mixes RiverWare and Excel, register both kinds of references together.

Do not manufacture cell addresses in `report_builder.py` if the Excel loader already provides a structured entry containing the actual source cell.

---

# Part VI - Changing an Existing Calculation

## 21. Find the current calculation first

Do not search only for the report wording. Identify the actual arithmetic and all associated Source references.

Examples in the current project:

- yearly state and basin totals are in `_build_year_section()`;
- Arizona reduction narrative is in `_az_reductions_text()`;
- California conservation roll-up is in `_build_california()`;
- Nevada conservation decomposition is in `_build_nevada()`;
- ICS totals are in `_build_ics_table()`.

When changing a formula, review all of these elements together:

1. input retrieval;
2. arithmetic;
3. sign convention;
4. unit conversion;
5. display rounding;
6. sentence wording;
7. Source calculation label;
8. Source input references;
9. year/scenario condition; and
10. automated tests.

---

## 22. Worked example: modify a calculation using existing slots

The current basin total is conceptually calculated from:

```text
California + Arizona + Nevada + Mexico
```

The implementation in `_build_year_section()` computes the state totals and Mexico schedule, then calculates total use in maf.

If an approved future rule changed that total, do **not** modify only the numeric expression. Also update the Source metadata that describes the formula and the test expectations.

A safe procedure is:

1. Identify every old input.
2. Identify every new input.
3. Check whether all new inputs are already in `required_slots` or `optional_slots`.
4. If not, register them in policy first.
5. Retrieve all values in acre-feet.
6. Calculate in acre-feet whenever practical.
7. Convert to maf only for the final display.
8. Update `_calc_ref(...)` so its locator/formula description matches the new arithmetic.
9. Add `_rw_ref(...)` or Excel references for every new source input.
10. Remove Source references for inputs no longer used.
11. Update the regression test to prove the new exact result.
12. Test at least one representative real model.

A mismatch between the formula and the Source list is a QA defect even if the displayed number is correct.

---

## 23. Avoid double rounding

Prefer this pattern:

```python
combined_af = a_af + b_af
text = _fmt_dec(combined_af / 1000, 1)
```

Avoid this pattern unless the business rule explicitly requires component-level rounding:

```python
combined_kaf = round(a_af / 1000, 1) + round(b_af / 1000, 1)
```

Those two methods can produce different displayed totals.

The report should round at the point required by the approved reporting rule, not merely where it is convenient in code.

---

## 24. Be explicit about signs

Some workbook/model values represent reductions or water left in Mead using a negative sign, while the report may display the magnitude as a positive conservation/reduction volume.

Do not automatically apply `abs(...)` to a new value.

Instead, document the intended sign rule. Example:

```python
reduction_af = -raw_value
```

should only be used when the source convention is known to encode the report concept with the opposite sign.

Tests should include a representative sign case.

---

# Part VII - Adding a New Calculated Bullet from Two New Slots

## 25. Full step-by-step recipe

Suppose a future approved report requirement says:

> Add a California bullet called "Example combined conservation" equal to RiverWare Slot A + RiverWare Slot B, displayed in kaf to one decimal place. Do not show the bullet when the combined value is zero.

This is a hypothetical example.

### Step 1 - identify exact RiverWare names

For example:

```text
ExampleObject.AnnualA
ExampleObject.AnnualB
```

Verify annual timestep, three target years, and units in a representative model.

### Step 2 - classify each source as required or optional

If the report is invalid without these values, add both to `required_slots`.

### Step 3 - edit policy

Add the exact strings to `config/report_policy.json`.

Increment `policy_version`.

### Step 4 - validate JSON

```bat
.venv\Scripts\python.exe -m json.tool config\report_policy.json > NUL
```

### Step 5 - edit the correct builder

Because this is a California bullet, edit:

```text
report_builder.py -> _build_california()
```

### Step 6 - retrieve inputs

```python
slot_a = "ExampleObject.AnnualA"
slot_b = "ExampleObject.AnnualB"

a_af = _slot(annual, slot_a, year)
b_af = _slot(annual, slot_b, year)
```

### Step 7 - validate availability in code

For required slots, preflight should already have guaranteed valid values. Still, explicit checks make the calculation safer and easier to understand:

```python
if a_af is not None and b_af is not None:
```

For optional inputs, decide the approved missing-data behavior before writing the formula. Do not silently substitute zero unless that is the intended rule.

### Step 8 - calculate before formatting

```python
combined_af = a_af + b_af
```

### Step 9 - apply display condition

```python
if combined_af > 0:
```

Use `!= 0`, `> 0`, or another condition only according to the reporting rule.

### Step 10 - format for the report

```python
display_kaf = _fmt_dec(combined_af / 1000, 1)
```

### Step 11 - register calculation and raw sources

```python
source_id = _register(
    registry,
    [
        _calc_ref(
            "AnnualA + AnnualB",
            display_kaf,
            "kaf",
            "Example combined conservation calculation.",
        ),
        _rw_ref(annual, slot_a, year),
        _rw_ref(annual, slot_b, year),
    ],
)
```

### Step 12 - append the bullet at the desired location

```python
section.bullets.append(
    Bullet(
        f"Example combined conservation of {display_kaf} kaf",
        source_id=source_id,
    )
)
```

Put this block at the exact point in `_build_california()` where the bullet should appear relative to the existing bullets.

### Step 13 - do not edit HTML or DOCX rendering for a normal bullet

The canonical model handles it automatically.

### Step 14 - add tests

See Part IX below.

### Step 15 - run with a real model and inspect Source

Confirm that the Source panel shows:

- the calculation description;
- Slot A;
- Slot B;
- the correct report year;
- converted acre-foot values; and
- the original model unit information.

### Step 16 - visually inspect the Word report

Confirm bullet order, rounding, wording, indentation, and page flow.

---

# Part VIII - Making a New Calculation Configurable by Rule

## 26. When to add a JSON rule in addition to Python

Do not hard-code a value in Python if it is likely to change as policy.

Good candidates for `report_policy.json` include:

- effective start year;
- scenario applicability;
- approved display label;
- approved numeric constant or override;
- whether a bullet is shown;
- source-selection option among behaviors already supported by code.

The Python should implement the mechanism; JSON should state the active policy.

---

## 27. Example policy structure for a future bullet

A possible future configuration might look like:

```json
"example_combined_conservation": {
  "effective_from_year": 2027,
  "scenarios": ["Most", "Min"],
  "label": "Example combined conservation"
}
```

Adding this JSON alone does nothing. The application must have Python code that reads and applies it.

A maintainable approach is:

1. add the configuration under `report.rules`;
2. add a small helper in `policy.py` if the logic is reused or nontrivial;
3. use that helper from `report_builder.py`; and
4. add policy tests for the effective year/scenario behavior.

Avoid scattering raw JSON dictionary lookups throughout many builder functions when a named policy method would communicate the rule more clearly.

---

## 28. Scenario-specific bullets

If a new bullet should appear only for selected scenarios, store the scenario list in policy rather than inferring it from a filename.

The scenario key is one of:

```text
Most
Min
Max
```

The user-facing labels are configured separately.

Within builder code that has access to `state_use`, the selected scenario is available from:

```python
state_use.scenario
```

For functions that do not currently receive scenario/state-use information, change the function interface deliberately rather than using global state.

Add tests for every scenario whose behavior differs.

---

## 29. Effective-year rules

The current policy already uses `effective_from_year` for several rules.

Follow the same general model when a rule changes at a calendar-year boundary. Keep the year threshold in JSON and the interpretation in `policy.py` / `report_builder.py`.

Test at least:

- the last year before the rule takes effect; and
- the first year in which it applies.

For an effective-from-2027 rule, test both 2026 and 2027.

---

# Part IX - Automated Tests for New Slots and Calculations

## 30. Why tests are mandatory for numeric changes

The application can produce a plausible-looking Word document even when a source mapping or arithmetic rule is wrong. Therefore a successful document generation is not enough.

Every material numeric change should have a regression test that states the expected result from controlled inputs.

---

## 31. Use `tests/test_report_builder.py`

The main controlled report-calculation tests are in:

```text
tests/test_report_builder.py
```

The helper `make_annual()` builds a synthetic `AnnualExtractionResult`.

For every slot currently in `policy.required_slots`, it initially creates a `SlotResult` with zero values for 2026, 2027, and 2028. Therefore, if you add a new **required** slot to the JSON, the test fixture automatically creates the slot, but it will contain zeros until you assign test values.

Use the fixture's existing `set_values(...)` pattern to give your new slot controlled values.

Example:

```python
set_values("ExampleObject.AnnualA", [100_000, 110_000, 120_000])
set_values("ExampleObject.AnnualB", [25_000, 30_000, 35_000])
```

If you add an **optional** slot, `make_annual()` will not automatically create it because it initializes required slots. Add a `SlotResult` explicitly in the test fixture or in the individual test.

---

## 32. Test the bullet text

For a California bullet, a test can retrieve the California section and assert the expected text.

Illustrative pattern:

```python
ctx = build_report_context(
    make_annual(),
    make_state_use("Most"),
    "Most",
    "August 2026",
    policy,
)

california = next(
    s for s in ctx.years[0].states
    if s.heading == "California"
)

assert any(
    b.text == "Example combined conservation of 125.0 kaf"
    for b in california.bullets
)
```

Use exact expected values for calculation tests whenever possible.

---

## 33. Test Source tracing too

Do not test only the visible sentence.

Find the new bullet, read its `source_id`, then inspect `ctx.source_map`.

Illustrative pattern:

```python
bullet = next(
    b for b in california.bullets
    if b.text.startswith("Example combined conservation")
)

refs = ctx.source_map[bullet.source_id]
locators = {ref["locator"] for ref in refs}

assert "ExampleObject.AnnualA" in locators
assert "ExampleObject.AnnualB" in locators
assert "AnnualA + AnnualB" in locators
```

This guards against a future refactor displaying the right number while tracing it to the wrong source.

---

## 34. Test conditions that suppress the bullet

If the bullet should not appear when the result is zero, write a zero-value test.

If it is scenario-specific, test included and excluded scenarios.

If it is year-specific, test both sides of the year boundary.

If it is optional, test behavior when the slot is absent.

---

## 35. Add policy tests when policy behavior changes

Use:

```text
tests/test_policy.py
```

Examples of useful assertions:

- the new required slot is present;
- the new optional slot is present;
- a scenario list contains the intended scenarios;
- an effective-year method returns the expected rule before and after the boundary.

You do not need to assert every JSON string. Test items that encode important report behavior.

---

## 36. Run the standard tests

From Windows:

```bat
run_tests.bat
```

or directly:

```bat
.venv\Scripts\python.exe -m pytest -q
```

Do not release a calculation change with failing tests.

---

# Part X - Validate the New Slot Against a Real RiverWare Model

## 37. Required-slot integration test

The project contains:

```text
tests/test_real_model_optional.py
```

Despite its filename, this is an optional-to-run integration test whose current assertion validates **all policy-required slots** against a locally supplied model.

### Command Prompt

```bat
set WU_TEST_MODEL=C:\Path\To\RepresentativeModel.mdl.gz
.venv\Scripts\python.exe -m pytest -q -m integration
```

### PowerShell

```powershell
$env:WU_TEST_MODEL = "C:\Path\To\RepresentativeModel.mdl.gz"
.\.venv\Scripts\python.exe -m pytest -q -m integration
```

If the new slot is required, this test should fail if the representative model cannot supply it correctly.

### Important limitation for optional slots

The current real-model integration test calls the extractor with the required-slot list. Therefore a newly added **optional** slot is not fully proven by this existing assertion. For an optional slot, also run the app with a representative model and verify preflight/Source behavior, or extend the integration test to inspect that optional slot explicitly.

---

## 38. What a RiverWare extraction failure means

For a required slot, preflight can fail because:

- the object or slot name is wrong;
- the slot does not exist in that scenario/model version;
- the series is not 1 YEAR;
- one of the target-year values is missing;
- the unit is missing;
- the unit is not accepted by policy;
- the unit is accepted by policy text but not actually supported by the conversion function; or
- the slot uses a RiverWare serialization form the parser does not currently handle.

Do not resolve a legitimate failure by simply moving an essential source to `optional_slots`.

---

## 39. Units: policy acceptance versus actual conversion support

The policy currently lists accepted volume-unit labels such as:

```text
m3
acre-ft
ac-ft
acreft
ft3
```

`riverware_annual.py` then normalizes unit names and performs the actual conversion.

The converter currently understands acre-foot variants, cubic meters, and cubic feet. Adding a string to `accepted_volume_units` does **not** create a new conversion formula.

If a future slot uses a genuinely new unit, update and test `volume_to_acre_ft()` in `riverware_annual.py`.

Never approve a unit by JSON alone unless the converter already understands it.

---

# Part XI - End-to-End Application Validation

## 40. Restart after policy edits

The Flask application loads the policy when the process starts.

After changing `report_policy.json`:

1. stop the app;
2. restart with `run_app.bat`; and
3. click **Rules** to verify the new policy version and slot list are active.

Editing JSON while the app is running does not update the already loaded policy.

---

## 41. Run preflight with representative inputs

Use a model that should contain the new slot and the current Projected State Use workbook.

Check:

- preflight passes;
- the required-slot count changed as expected if you added a required slot;
- no unexpected unit/timestep warning appears;
- selected scenario is correct; and
- study years are correct.

For a change that affects multiple scenarios, test each affected scenario. Do not assume a slot present in Most also exists in Min and Max.

---

## 42. Generate the preview and use Source

For the new bullet or changed calculation:

1. locate the value in the preview;
2. click **Source**;
3. confirm the fully qualified RiverWare slot name(s);
4. confirm the year;
5. confirm the extracted numeric values;
6. confirm source model units/conversion notes;
7. confirm the calculation formula/description if calculated; and
8. confirm no unrelated source is listed.

The Source check is one of the fastest ways to catch accidental use of a similar-looking slot.

---

## 43. Inspect the generated Word document

Check the DOCX independently of the HTML preview.

For a new bullet verify:

- wording;
- numeric value;
- AF/kaf/maf label;
- decimal places;
- bullet level;
- order within the section;
- page flow; and
- presence/absence for the correct years and scenarios.

Remember: Source metadata is intentionally not embedded in the Word document.

---

# Part XII - Detailed Change Scenarios

## 44. Scenario A: change only the arithmetic of an existing bullet

Use this checklist:

1. Find the bullet in `report_builder.py`.
2. Identify every input currently used.
3. Write the old and new formulas side by side.
4. Confirm whether the new formula needs any new source slot.
5. If yes, add the slot to policy.
6. Change the arithmetic.
7. Update display formatting only if the reporting requirement changed.
8. Update `_calc_ref(...)` to state the new formula.
9. Update registered raw Source references.
10. Update/add regression tests with controlled values.
11. Run standard tests.
12. Run real-model integration if RiverWare source dependencies changed.
13. Restart the app.
14. Preflight affected scenarios.
15. Inspect Source.
16. Inspect Word output.
17. Update `CHANGELOG.md`.
18. Package/archive the new version and retain rollback copy.

---

## 45. Scenario B: add one direct RiverWare bullet

Use this checklist:

1. Confirm exact `Object.Slot` name.
2. Confirm annual timestep.
3. Confirm volume units.
4. Confirm expected years.
5. Decide required versus optional.
6. Add slot to `report_policy.json`.
7. Increment `policy_version`.
8. Validate JSON.
9. Open the appropriate `_build_*()` function.
10. Retrieve with `_slot()`.
11. Preserve `None` if missing has a distinct meaning.
12. Apply sign rule.
13. Apply show/suppress condition.
14. Convert to AF/kaf/maf for display.
15. Format with `_fmt_int()` or `_fmt_dec()`.
16. Append a `Bullet(...)` at the intended order position.
17. Register `_rw_ref(...)` as its Source.
18. Add a regression test.
19. Run standard tests.
20. Run real-model validation.
21. Restart and preflight.
22. Verify Source.
23. Inspect Word output.
24. Update changelog/release archive.

---

## 46. Scenario C: add a calculated bullet from several RiverWare slots

Follow Scenario B for every new slot, then additionally:

1. retrieve each input independently;
2. explicitly define missing-data behavior;
3. perform arithmetic in acre-feet;
4. convert only for display;
5. add `_calc_ref(...)` describing the formula;
6. register `_rw_ref(...)` for every RiverWare input;
7. assert the calculated value in a unit test;
8. assert the Source locator set in the same test; and
9. test zero/negative/missing behavior where relevant.

---

## 47. Scenario D: add a bullet that varies by year or scenario

In addition to the steps above:

1. put changeable year/scenario parameters in `report_policy.json`;
2. add a descriptive helper to `policy.py` when practical;
3. avoid inferring scenario from the model filename for report policy;
4. test each distinct branch;
5. inspect all affected scenarios in batch mode; and
6. verify the Rules panel displays the new configuration.

---

# Part XIII - When a New Bullet Requires More Than `report_builder.py`

## 48. Existing state bullet: no renderer changes

If the output is simply another bullet under California, Arizona, Nevada, or Mexico, use the existing `Bullet` model and renderer paths.

Normally modify only:

```text
config/report_policy.json   (if a new slot or configurable policy is needed)
policy.py                   (if a new policy helper is useful)
report_builder.py
tests/...
CHANGELOG.md
```

---

## 49. Entirely new section or new table

A genuinely new content type requires broader work.

Typical files are:

```text
report_models.py
report_builder.py
templates/report_preview.html
docx_writer.py
static/app.css              (if new preview styling is needed)
tests/...
```

Examples:

- a new standalone "Operations" section;
- a new table with different columns;
- a paragraph with special formatting not represented by current models;
- a new chart or graphic.

Design the canonical data structure first in `report_models.py`, then make both HTML and Word consume that structure.

Do not calculate a value separately inside each renderer.

---

# Part XIV - Parser Changes for Unsupported RiverWare Sources

## 50. When simply adding a slot name is not enough

Edit `riverware_annual.py` only if the slot is not representable through the existing extraction path.

Examples include:

- new RiverWare serialization syntax;
- a table/periodic source that needs special extraction;
- a nonannual series that requires aggregation;
- a new aggregate-member encoding;
- a new physical unit conversion; or
- a model-format change after a RiverWare upgrade.

Parser changes are high risk because they can affect every report source.

Add focused parser tests and rerun the production/sanitized-model integration test after any parser change.

---

# Part XV - Recommended Coding Practices for Future Changes

## 51. Use named slot variables

Prefer:

```python
example_slot = "ExampleObject.ExampleAnnualVolume"
value = _slot(annual, example_slot, year)
```

over repeating the same literal in multiple places.

This reduces copy/paste errors between retrieval and Source tracing.

---

## 52. Keep source retrieval, calculation, and presentation distinct

A clear block generally has three stages:

```text
1. Retrieve source data
2. Calculate business value
3. Format and append report content
```

Do not bury complex arithmetic inside a long f-string.

---

## 53. Prefer explicit missing-data behavior

For a required slot, preflight should prevent missing values from reaching report generation.

For optional slots, choose one explicit behavior:

- omit bullet;
- display a prescribed "not available" statement; or
- substitute a value only if policy explicitly authorizes it.

Do not silently treat unknown as zero by default.

---

## 54. Keep policy constants out of formula code

A number that changes because of operating policy should normally live in JSON.

A constant that is intrinsic to unit conversion or software mechanics belongs in Python.

Examples:

```text
760 kaf policy display override -> JSON
1,000 AF per kaf               -> code/display conversion
43,560 ft3 per acre-ft         -> unit-conversion code
```

---

## 55. Keep the Source explanation synchronized

Whenever you change a formula, ask:

> If someone clicks Source six months from now, will the Source panel accurately explain the number they see?

If not, the change is incomplete.

---

# Part XVI - Release and Review Checklist

## 56. Numeric/report-logic change checklist

Before releasing a new calculation or slot change, confirm all of the following:

- [ ] Business rule is written in plain language.
- [ ] Exact RiverWare slot name(s) verified.
- [ ] Required/optional classification approved.
- [ ] Units verified.
- [ ] Sign convention verified.
- [ ] Year applicability verified.
- [ ] Scenario applicability verified.
- [ ] `report_policy.json` updated if required.
- [ ] `policy_version` incremented.
- [ ] JSON syntax validated.
- [ ] Calculation implemented in `report_builder.py`.
- [ ] No duplicate calculation added to HTML/Word renderers.
- [ ] Source references updated.
- [ ] Automated regression test added/updated.
- [ ] Standard tests pass.
- [ ] Real-model integration test passes when required slots changed.
- [ ] Representative application preflight passes.
- [ ] Source panel verified.
- [ ] Most/Min/Max branches tested as applicable.
- [ ] Generated Word file visually inspected.
- [ ] `CHANGELOG.md` updated.
- [ ] Previous known-good release retained for rollback.

---

# Part XVII - Troubleshooting

## 57. Preflight says `Slot not found`

Check:

1. exact object name;
2. exact slot name;
3. capitalization and spaces;
4. whether the slot exists in that scenario model;
5. whether the configured name is an aggregate member that the parser can resolve; and
6. whether the RiverWare model version renamed or moved the slot.

Do not change calculation code until extraction works.

---

## 58. Preflight says expected 1-year series

The current report extractor expects time-series inputs to be annual with a 1 YEAR timestep.

If the business source is monthly/daily and must be aggregated, that is a new calculation/parser requirement. Define the aggregation rule explicitly before coding it.

---

## 59. Preflight says unit is not accepted

First verify the actual RiverWare unit.

Then distinguish two cases:

### Converter already supports the unit, but policy does not allow its label

Update `accepted_volume_units` after confirming normalization behavior.

### Converter does not support the unit

Implement and test the physical conversion in `riverware_annual.py` first. Do not merely add the label to JSON.

---

## 60. Bullet is missing even though slot extracts correctly

Check:

- `if` condition around the bullet;
- whether value is zero or negative;
- effective-year condition;
- scenario condition;
- whether the code was added to the correct `_build_*()` function;
- whether Flask was restarted after policy changes; and
- whether you are running the project copy you actually edited.

---

## 61. Bullet value is correct but Source is wrong

Check that:

- the same slot-name variable is used for `_slot()` and `_rw_ref()`;
- every calculation input is registered;
- stale inputs were removed from `_register(...)`; and
- `_calc_ref(...)` describes the current formula.

Add/update a Source-map assertion in `test_report_builder.py`.

---

## 62. Preview is correct but Word is wrong

For a normal bullet this should be unusual because both outputs use `ReportContext`.

If text differs:

1. verify the bullet exists correctly in `ReportContext`;
2. inspect `templates/report_preview.html` and `docx_writer.py` only after confirming the model;
3. check Word styling/list behavior; and
4. ensure the issue is presentation, not duplicate calculation logic.

---

# Part XVIII - Example Change Record

## 63. Document the change in `CHANGELOG.md`

A useful entry for a new numeric source might say:

```markdown
### Changed
- Added RiverWare slot `ExampleObject.ExampleAnnualVolume` as a required report source.
- Added California bullet "Example activity" using the annual slot value in kaf.
- Added Source tracing for the new slot.
- Added regression coverage for 2026-2028 and Most/Min scenario behavior.
```

For a formula change, state both the previous concept and new concept clearly enough for a future reviewer to understand why historical outputs differ.

---

# Part XIX - Quick Reference

## 64. Add a new direct RiverWare bullet

```text
1. Verify exact annual slot + unit in RiverWare.
2. Add it to required_slots or optional_slots in report_policy.json.
3. Increment policy_version.
4. Validate JSON.
5. Retrieve it with _slot() in the appropriate _build_*() function.
6. Calculate/sign-adjust if required.
7. Format AF/kaf/maf.
8. Append Bullet(...).
9. Register _rw_ref(...) for Source.
10. Add regression + Source tests.
11. Run standard tests.
12. Run real-model test if required.
13. Restart app, preflight, inspect Source.
14. Generate and visually inspect Word.
15. Update changelog and archive release.
```

## 65. Change a calculation

```text
1. Find current arithmetic in report_builder.py.
2. Write old/new formulas.
3. Register any new sources in policy.
4. Retrieve raw values in acre-ft.
5. Apply approved missing/sign rules.
6. Calculate before display rounding.
7. Update _calc_ref(...) formula description.
8. Register every underlying source.
9. Update exact regression expectations.
10. Test year/scenario branches.
11. Run real-model QA and inspect Source.
12. Inspect Word output.
```

---

# Part XX - Final Architectural Rule

## 66. Where future logic belongs

Use `report_policy.json` for **declarative policy**: values and choices that may change operationally while using an already implemented behavior.

Use `report_builder.py` for **report mathematics and narrative assembly**.

Use `riverware_annual.py` for **RiverWare parsing and unit conversion mechanics**.

Use `report_models.py` to define **canonical report content structures**.

Use `templates/report_preview.html` and `docx_writer.py` only for **presentation**.

Use `tests/` to permanently record **what the approved rule is expected to produce**.

Maintaining this separation is what allows the app to remain auditable as report rules change over time.

---

## Guided planner utility

For a normal new RiverWare-derived bullet or calculation change, maintainers can run `run_change_planner.bat` (or `report_change_planner.py`). The utility asks for the slots, formula, placement, conditions, wording, units/rounding, required-vs-optional behavior, and regression-test values, then generates a detailed Markdown implementation plan. It does **not** automatically modify the application. See `REPORT_CHANGE_PLANNER_GUIDE.md`.
