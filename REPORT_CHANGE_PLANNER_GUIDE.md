# Report Change Planner

`report_change_planner.py` is a maintainer utility for planning a new RiverWare-derived report bullet or a calculation change in the WU Assumptions Generator.

It is intentionally **non-destructive**. It does not edit the application. It asks a structured set of questions, inspects the current `config/report_policy.json` when available, and writes a detailed Markdown implementation plan containing the policy registration, `report_builder.py` code, Source registration, condition handling, tests, validation steps, and the files that should or should not change.

## Windows quick start

From an extracted application folder, double-click:

```text
run_change_planner.bat
```

The utility uses the application's `.venv` when it exists. Because the planner itself uses only the Python standard library, it can also run with a normal Python 3 installation.

To run it from a command prompt:

```bat
python report_change_planner.py --project-root .
```

## Information the planner asks for

The wizard collects:

1. a short change name and proposed policy version;
2. one or more fully qualified RiverWare slot names;
3. a short Python alias for each slot;
4. whether each new slot is required or optional;
5. whether each source is a normal annual volume slot supported by the current parser;
6. the calculation, written using the aliases (RiverWare volume inputs are in acre-ft after extraction);
7. the plain-language business meaning of the calculation;
8. the report section: California, Arizona, Nevada, or Mexico;
9. where the new bullet should appear relative to existing bullets;
10. applicable scenarios;
11. applicable years;
12. the value condition for displaying the bullet;
13. exact bullet wording;
14. displayed units and decimal precision;
15. bullet nesting level;
16. optional Source-panel explanation; and
17. deterministic test values for the current three-year test fixture.

The formula field accepts arithmetic using the aliases plus numeric constants, parentheses, `abs()`, `min()`, and `max()`. The planner validates the expression but does not execute arbitrary Python entered by the user.

## Output

By default, the generated plan is written under:

```text
change_plans\<rule_key>_implementation_plan.md
```

The plan includes:

- slots that must be added to `riverware.required_slots` or `riverware.optional_slots`;
- a proposed `report.rules.<rule_key>` configuration;
- scenario propagation changes when the target v2.1 builder needs them;
- the `report_builder.py` extraction/calculation/condition/Bullet code;
- `_rw_ref`, `_calc_ref`, and `_register` Source-tracing code;
- deterministic `tests/test_report_builder.py` fixture/test snippets;
- required/optional missing-slot tests;
- JSON validation commands;
- real-model integration-test commands;
- parser/unit warnings when a source is not a normal annual volume slot; and
- a release/review checklist.

## Reusing an approved questionnaire

The planner can save and reload the structured answers. To save answers while running interactively:

```bat
python report_change_planner.py --project-root . --save-answers change_plans\my_change_answers.json
```

To regenerate a plan later without answering the questions again:

```bat
python report_change_planner.py --project-root . ^
  --answers change_plans\my_change_answers.json ^
  --output change_plans\my_change_implementation_plan.md
```

This is useful for peer review because the requested rule can be reviewed separately from the generated instructions.

## Important design boundary

The planner follows the application's architecture:

- **JSON** stores source registration, applicability, thresholds, labels, display settings, and other editable policy.
- **Python** implements executable calculations.
- The new item is registered as a canonical `Bullet`, so the same value/wording feeds the HTML preview and Word document.
- Source tracing is registered in the same builder code so the preview can explain the RiverWare inputs and calculation.

The planner does **not** turn `report_policy.json` into a formula engine.

## Cases requiring extra developer work

The generated normal-bullet pattern assumes each new source can be obtained by `_slot(...)` as an annual volume normalized to acre-ft. If a source is monthly, periodic/table based, non-volume, uses an unsupported RiverWare serialization, or uses a new physical unit, review and test `riverware_annual.py` before applying the generated report-builder code.

Likewise, a brand-new report section/table, a change to state/basin totals, or a Word-layout change may require additional model/renderer changes beyond a normal existing-section bullet.

## Safety/review recommendation

Treat the generated Markdown as a proposed implementation, not an automatic patch. For calculation changes, have another reviewer verify slot meaning, sign convention, formula, units, rounding, scenario/year boundaries, and the final Word output before operational use.
