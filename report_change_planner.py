#!/usr/bin/env python3
"""Interactive change planner for the WU Assumptions Generator.

This utility does NOT modify the application. It interviews a maintainer about a
new RiverWare-derived report bullet (or a calculation using one or more slots)
and generates a Markdown implementation plan containing:

* report_policy.json registration/configuration changes;
* report_builder.py code for extraction, calculation, conditions, bullet text,
  and Source tracing;
* any scenario-parameter propagation required by the current v2.1 architecture;
* policy.py guidance when applicable;
* regression-test updates; and
* validation/release steps.

The script uses only the Python standard library and is intended to be run from
(or pointed at) the WU Assumptions Generator project root.
"""
from __future__ import annotations

import argparse
import ast
import json
import keyword
import math
import operator
import re
import sys
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable


APP_NAME = "WU Assumptions Generator"
KNOWN_SCENARIOS = ("Most", "Min", "Max")
SECTIONS = {
    "California": "_build_california",
    "Arizona": "_build_arizona",
    "Nevada": "_build_nevada",
    "Mexico": "_build_mexico",
}
FORMAT_UNITS = ("af", "kaf", "maf")
ALLOWED_FORMULA_FUNCS = {"abs", "min", "max"}
ALLOWED_CUSTOM_NAMES = {"computed_af", "year", "scenario"}


@dataclass
class SlotSpec:
    alias: str
    slot_name: str
    registration: str  # existing-required, existing-optional, required, optional
    description: str = ""
    annual_volume_confirmed: bool = True
    test_values_af: list[float] = field(default_factory=list)

    @property
    def is_new(self) -> bool:
        return self.registration in {"required", "optional"}

    @property
    def is_optional(self) -> bool:
        return self.registration in {"optional", "existing-optional"}


@dataclass
class ChangeSpec:
    title: str
    rule_key: str
    policy_version_current: str
    policy_version_proposed: str
    slots: list[SlotSpec]
    formula: str
    formula_description: str
    missing_optional_behavior: str
    section: str
    placement: str
    placement_anchor: str
    scenarios: list[str]
    year_mode: str
    year_value: str
    display_condition: str
    threshold_af: float | None
    custom_condition: str
    bullet_text: str
    bullet_level: int
    display_units: str
    decimals: int
    source_note: str
    test_years: list[int]
    notes: str = ""


class PlannerError(ValueError):
    pass


def find_project_root(start: Path) -> Path | None:
    """Find a project root containing the files this planner understands."""
    candidates = [start, *start.parents]
    if start.is_file():
        candidates = [start.parent, *start.parent.parents]
    for candidate in candidates:
        if (candidate / "report_builder.py").is_file() and (candidate / "config" / "report_policy.json").is_file():
            return candidate
    return None


def load_project_policy(project_root: Path | None) -> dict[str, Any]:
    if project_root is None:
        return {}
    path = project_root / "config" / "report_policy.json"
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise PlannerError(f"Could not read {path}: {exc}") from exc


def current_slot_status(policy: dict[str, Any], slot_name: str) -> str | None:
    rw = policy.get("riverware", {})
    if slot_name in rw.get("required_slots", []):
        return "existing-required"
    if slot_name in rw.get("optional_slots", []):
        return "existing-optional"
    return None


def slugify(text: str) -> str:
    value = re.sub(r"[^A-Za-z0-9]+", "_", text.strip()).strip("_").lower()
    return value or "report_change"


def safe_identifier(value: str) -> bool:
    return bool(value) and value.isidentifier() and not keyword.iskeyword(value)


def suggest_alias(slot_name: str, used: set[str]) -> str:
    tail = slot_name.split(".")[-1]
    candidate = slugify(tail)
    if not candidate or candidate[0].isdigit():
        candidate = f"slot_{candidate}"
    base = candidate
    n = 2
    while candidate in used or not safe_identifier(candidate):
        candidate = f"{base}_{n}"
        n += 1
    return candidate


def bump_policy_version(version: str) -> str:
    parts = version.split(".")
    if parts and parts[-1].isdigit():
        parts[-1] = str(int(parts[-1]) + 1)
        return ".".join(parts)
    stamp = datetime.now().strftime("%Y.%m")
    return f"{stamp}.1"


def ask(prompt: str, default: str | None = None, allow_blank: bool = False) -> str:
    suffix = f" [{default}]" if default is not None else ""
    while True:
        try:
            value = input(f"{prompt}{suffix}: ").strip()
        except EOFError as exc:
            raise PlannerError("Input ended before the questionnaire was complete.") from exc
        if not value and default is not None:
            return default
        if value or allow_blank:
            return value
        print("  A value is required.")


def ask_yes_no(prompt: str, default: bool = True) -> bool:
    default_text = "Y/n" if default else "y/N"
    while True:
        value = ask(f"{prompt} ({default_text})", "y" if default else "n").lower()
        if value in {"y", "yes"}:
            return True
        if value in {"n", "no"}:
            return False
        print("  Enter y or n.")


def ask_choice(prompt: str, choices: Iterable[str], default: str | None = None) -> str:
    choices = list(choices)
    display = ", ".join(f"{i + 1}={c}" for i, c in enumerate(choices))
    while True:
        raw = ask(f"{prompt} ({display})", default)
        if raw.isdigit() and 1 <= int(raw) <= len(choices):
            return choices[int(raw) - 1]
        for choice in choices:
            if raw.lower() == choice.lower():
                return choice
        print("  Choose one of the listed values or enter its number.")


def ask_int(prompt: str, default: int | None = None, min_value: int | None = None, max_value: int | None = None) -> int:
    while True:
        raw = ask(prompt, str(default) if default is not None else None)
        try:
            value = int(raw)
        except ValueError:
            print("  Enter a whole number.")
            continue
        if min_value is not None and value < min_value:
            print(f"  Value must be at least {min_value}.")
            continue
        if max_value is not None and value > max_value:
            print(f"  Value must be no more than {max_value}.")
            continue
        return value


def ask_float_list(prompt: str, count: int) -> list[float]:
    while True:
        raw = ask(prompt)
        parts = [p.strip().replace(",", "") for p in raw.split(";") if p.strip()]
        if len(parts) != count:
            print(f"  Enter exactly {count} values separated by semicolons, e.g. 1000; 2000; 3000.")
            continue
        try:
            return [float(p) for p in parts]
        except ValueError:
            print("  Every test value must be numeric.")


def validate_formula(expr: str, aliases: set[str]) -> None:
    """Validate a non-executed arithmetic expression for code-generation safety."""
    try:
        tree = ast.parse(expr, mode="eval")
    except SyntaxError as exc:
        raise PlannerError(f"Formula is not valid Python arithmetic: {exc.msg}") from exc

    allowed_nodes = (
        ast.Expression,
        ast.BinOp,
        ast.UnaryOp,
        ast.Add,
        ast.Sub,
        ast.Mult,
        ast.Div,
        ast.FloorDiv,
        ast.Mod,
        ast.Pow,
        ast.UAdd,
        ast.USub,
        ast.Name,
        ast.Constant,
        ast.Call,
        ast.Load,
    )
    for node in ast.walk(tree):
        if not isinstance(node, allowed_nodes):
            raise PlannerError(f"Formula uses unsupported syntax: {type(node).__name__}")
        if isinstance(node, ast.Name) and node.id not in aliases and node.id not in ALLOWED_FORMULA_FUNCS:
            raise PlannerError(f"Unknown formula name '{node.id}'. Use only aliases: {', '.join(sorted(aliases))}.")
        if isinstance(node, ast.Call):
            if not isinstance(node.func, ast.Name) or node.func.id not in ALLOWED_FORMULA_FUNCS:
                raise PlannerError("Formula functions are limited to abs(), min(), and max().")
            if node.keywords:
                raise PlannerError("Keyword arguments are not supported in formulas.")


def validate_condition(expr: str, aliases: set[str]) -> None:
    if not expr.strip():
        return
    try:
        tree = ast.parse(expr, mode="eval")
    except SyntaxError as exc:
        raise PlannerError(f"Custom condition is not valid Python: {exc.msg}") from exc
    allowed_nodes = (
        ast.Expression,
        ast.BoolOp,
        ast.And,
        ast.Or,
        ast.UnaryOp,
        ast.Not,
        ast.USub,
        ast.UAdd,
        ast.BinOp,
        ast.Add,
        ast.Sub,
        ast.Mult,
        ast.Div,
        ast.Mod,
        ast.Name,
        ast.Constant,
        ast.Compare,
        ast.Eq,
        ast.NotEq,
        ast.Lt,
        ast.LtE,
        ast.Gt,
        ast.GtE,
        ast.In,
        ast.NotIn,
        ast.List,
        ast.Tuple,
        ast.Set,
        ast.Call,
        ast.Load,
    )
    allowed_names = aliases | ALLOWED_CUSTOM_NAMES | ALLOWED_FORMULA_FUNCS
    for node in ast.walk(tree):
        if not isinstance(node, allowed_nodes):
            raise PlannerError(f"Custom condition uses unsupported syntax: {type(node).__name__}")
        if isinstance(node, ast.Name) and node.id not in allowed_names:
            raise PlannerError(f"Unknown name '{node.id}' in custom condition.")
        if isinstance(node, ast.Call):
            if not isinstance(node.func, ast.Name) or node.func.id not in ALLOWED_FORMULA_FUNCS:
                raise PlannerError("Condition functions are limited to abs(), min(), and max().")
            if node.keywords:
                raise PlannerError("Keyword arguments are not supported in conditions.")


def _safe_eval_expr(expr: str, env: dict[str, float]) -> float:
    """Evaluate only the restricted arithmetic grammar accepted by validate_formula."""
    tree = ast.parse(expr, mode="eval")
    binops = {
        ast.Add: operator.add,
        ast.Sub: operator.sub,
        ast.Mult: operator.mul,
        ast.Div: operator.truediv,
        ast.FloorDiv: operator.floordiv,
        ast.Mod: operator.mod,
        ast.Pow: operator.pow,
    }
    unary = {ast.UAdd: operator.pos, ast.USub: operator.neg}
    funcs = {"abs": abs, "min": min, "max": max}

    def walk(node: ast.AST) -> float:
        if isinstance(node, ast.Expression):
            return walk(node.body)
        if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)):
            return float(node.value)
        if isinstance(node, ast.Name):
            return float(env[node.id])
        if isinstance(node, ast.BinOp) and type(node.op) in binops:
            return float(binops[type(node.op)](walk(node.left), walk(node.right)))
        if isinstance(node, ast.UnaryOp) and type(node.op) in unary:
            return float(unary[type(node.op)](walk(node.operand)))
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Name) and node.func.id in funcs:
            return float(funcs[node.func.id](*(walk(arg) for arg in node.args)))
        raise PlannerError(f"Unsupported expression node: {type(node).__name__}")

    result = walk(tree)
    if not math.isfinite(result):
        raise PlannerError("Test formula evaluated to a non-finite value.")
    return result


def validate_bullet_template(template: str) -> None:
    # Only these placeholders are supported; doubled braces are allowed by str.format.
    try:
        fields = [field for _, field, _, _ in __import__("string").Formatter().parse(template) if field]
    except ValueError as exc:
        raise PlannerError(f"Bullet wording has unmatched braces: {exc}") from exc
    allowed = {"value", "year", "scenario"}
    bad = [field for field in fields if field not in allowed]
    if bad:
        raise PlannerError(f"Unsupported bullet placeholder(s): {', '.join(bad)}. Allowed: {{value}}, {{year}}, {{scenario}}.")


def format_display(value_af: float, units: str, decimals: int) -> str:
    if units == "af":
        if decimals == 0:
            return f"{round(value_af):,.0f}"
        return f"{value_af:,.{decimals}f}"
    if units == "kaf":
        return f"{value_af / 1000:,.{decimals}f}"
    if units == "maf":
        return f"{value_af / 1_000_000:,.{decimals}f}"
    raise PlannerError(f"Unsupported display units: {units}")


def year_condition_code(spec: ChangeSpec, rule_var: str = "rule_cfg") -> str:
    if spec.year_mode == "All years":
        return "year_ok = True"
    if spec.year_mode == "From year onward":
        return f"year_ok = year >= int({rule_var}.get(\"effective_from_year\", {int(spec.year_value)}))"
    if spec.year_mode == "Through year":
        return f"year_ok = year <= int({rule_var}.get(\"through_year\", {int(spec.year_value)}))"
    if spec.year_mode == "Exact years":
        years = [int(x.strip()) for x in spec.year_value.split(",") if x.strip()]
        return f"year_ok = year in {{int(v) for v in {rule_var}.get(\"years\", {years!r})}}"
    raise PlannerError(f"Unknown year mode: {spec.year_mode}")


def display_condition_code(spec: ChangeSpec) -> list[str]:
    """Generate condition code; ordinary modes remain editable through JSON."""
    default_key = display_condition_key(spec)
    lines = [f'display_mode = str(rule_cfg.get("display_condition", {default_key!r})).lower()']
    if spec.display_condition == "Custom Python condition":
        lines.append(f"display_ok = bool({spec.custom_condition})")
        return lines
    lines.extend([
        'if display_mode == "always":',
        '    display_ok = True',
        'elif display_mode == "positive":',
        '    display_ok = computed_af > 0',
        'elif display_mode == "nonzero":',
        '    display_ok = computed_af != 0',
        'elif display_mode == "negative":',
        '    display_ok = computed_af < 0',
        'elif display_mode == "abs_above_threshold":',
        f'    display_ok = abs(computed_af) > float(rule_cfg.get("minimum_abs_af", {spec.threshold_af if spec.threshold_af is not None else 0!r}))',
        'else:',
        f'    raise ValueError(f"Unsupported display_condition for {spec.rule_key}: {{display_mode}}")',
    ])
    return lines


def policy_rule_object(spec: ChangeSpec) -> dict[str, Any]:
    data: dict[str, Any] = {
        "description": spec.formula_description or spec.title,
        "enabled_scenarios": spec.scenarios,
        "display_condition": display_condition_key(spec),
    }
    if spec.year_mode == "From year onward":
        data["effective_from_year"] = int(spec.year_value)
    elif spec.year_mode == "Through year":
        data["through_year"] = int(spec.year_value)
    elif spec.year_mode == "Exact years":
        data["years"] = [int(x.strip()) for x in spec.year_value.split(",") if x.strip()]
    if spec.display_condition == "Absolute value above threshold":
        data["minimum_abs_af"] = spec.threshold_af
    if spec.display_condition == "Custom Python condition":
        data["custom_condition_note"] = spec.custom_condition
    data["bullet_template"] = spec.bullet_text
    data["display_units"] = spec.display_units
    data["display_decimals"] = spec.decimals
    return data


def scenario_needed(spec: ChangeSpec) -> bool:
    """Whether the target v2.1 builder needs an explicit scenario parameter added.

    California and Arizona already receive StateUseData, whose scenario field is
    the selected scenario. Nevada and Mexico do not, so we propagate scenario
    explicitly to keep enabled_scenarios editable through JSON.
    """
    return spec.section in {"Nevada", "Mexico"}


def scenario_assignment_code(spec: ChangeSpec) -> str:
    if spec.section in {"California", "Arizona"}:
        return "scenario = state_use.scenario"
    return "# `scenario` is passed explicitly into this builder; see the signature changes below."


def display_condition_key(spec: ChangeSpec) -> str:
    return {
        "Always": "always",
        "Positive (> 0)": "positive",
        "Non-zero (!= 0)": "nonzero",
        "Negative (< 0)": "negative",
        "Absolute value above threshold": "abs_above_threshold",
        "Custom Python condition": "custom_code",
    }[spec.display_condition]


def output_format_code(spec: ChangeSpec) -> list[str]:
    """Generate formatting code that keeps units/precision editable in JSON."""
    return [
        f'display_units = str(rule_cfg.get("display_units", {spec.display_units!r})).lower()',
        f'display_decimals = int(rule_cfg.get("display_decimals", {spec.decimals}))',
        'if display_units == "af":',
        '    display_value = _fmt_int(computed_af) if display_decimals == 0 else _fmt_dec(computed_af, display_decimals)',
        'elif display_units == "kaf":',
        '    display_value = _fmt_int(round(computed_af / 1000, 0)) if display_decimals == 0 else _fmt_dec(computed_af / 1000, display_decimals)',
        'elif display_units == "maf":',
        '    display_value = _fmt_int(round(computed_af / 1_000_000, 0)) if display_decimals == 0 else _fmt_dec(computed_af / 1_000_000, display_decimals)',
        'else:',
        f'    raise ValueError(f"Unsupported display_units for {spec.rule_key}: {{display_units}}")',
    ]


def insertion_target(spec: ChangeSpec) -> tuple[str, str]:
    function = SECTIONS[spec.section]
    collection = "bullets" if spec.section == "Mexico" else "section.bullets"
    return function, collection


def locate_function_line(project_root: Path | None, function_name: str) -> int | None:
    if not project_root:
        return None
    path = project_root / "report_builder.py"
    try:
        lines = path.read_text(encoding="utf-8").splitlines()
    except OSError:
        return None
    needle = f"def {function_name}("
    for i, line in enumerate(lines, 1):
        if line.startswith(needle):
            return i
    return None


def locate_anchor_lines(project_root: Path | None, anchor: str) -> list[int]:
    if not project_root or not anchor:
        return []
    try:
        lines = (project_root / "report_builder.py").read_text(encoding="utf-8").splitlines()
    except OSError:
        return []
    a = anchor.lower()
    return [i for i, line in enumerate(lines, 1) if a in line.lower()][:10]


def render_builder_block(spec: ChangeSpec) -> str:
    _, collection = insertion_target(spec)
    lines: list[str] = []
    lines.append(f"    # {spec.title}")
    lines.append(f"    rule_cfg = policy.report.get(\"rules\", {{}}).get({spec.rule_key!r}, {{}})")
    lines.append(f"    {scenario_assignment_code(spec)}")
    lines.append(
        "    scenario_ok = scenario in set(rule_cfg.get(\"enabled_scenarios\", policy.scenario_keys))"
    )
    lines.append(f"    {year_condition_code(spec)}")
    lines.append("")

    for slot in spec.slots:
        lines.append(f"    {slot.alias}_slot = {slot.slot_name!r}")
        lines.append(f"    {slot.alias} = _slot(annual, {slot.alias}_slot, year)")
    lines.append("")

    aliases = [s.alias for s in spec.slots]
    optional_aliases = [s.alias for s in spec.slots if s.is_optional]
    if optional_aliases and spec.missing_optional_behavior == "Suppress bullet if any optional input is missing":
        checks = " and ".join(f"{a} is not None" for a in aliases)
        lines.append(f"    inputs_available = {checks}")
    else:
        required_checks = [f"{s.alias} is not None" for s in spec.slots if not s.is_optional]
        lines.append(f"    inputs_available = {' and '.join(required_checks) if required_checks else 'True'}")
        if optional_aliases and spec.missing_optional_behavior == "Treat missing optional inputs as zero":
            for alias in optional_aliases:
                lines.append(f"    {alias} = 0.0 if {alias} is None else {alias}")
    lines.append("")
    lines.append("    if scenario_ok and year_ok and inputs_available:")
    lines.append(f"        computed_af = {spec.formula}")
    for condition_line in display_condition_code(spec):
        lines.append(f"        {condition_line}")
    lines.append("        if display_ok:")
    for format_line in output_format_code(spec):
        lines.append(f"            {format_line}")
    template_literal = repr(spec.bullet_text)
    lines.append(f"            bullet_template = str(rule_cfg.get(\"bullet_template\", {template_literal}))")
    format_args = ["value=display_value", "year=year", "scenario=scenario"]
    lines.append(f"            bullet_text = bullet_template.format({', '.join(format_args)})")
    lines.append("            source_refs = [")
    if spec.formula.strip() not in {s.alias for s in spec.slots}:
        note = spec.source_note or f"Calculated as: {spec.formula}. RiverWare source values are normalized to acre-ft before this calculation."
        lines.append("                _calc_ref(")
        lines.append(f"                    {spec.formula!r},")
        lines.append("                    display_value,")
        lines.append("                    display_units,")
        lines.append(f"                    {note!r},")
        lines.append("                ),")
    for slot in spec.slots:
        lines.append(f"                _rw_ref(annual, {slot.alias}_slot, year),")
    lines.append("            ]")
    lines.append(f"            {collection}.append(")
    lines.append("                Bullet(")
    lines.append("                    bullet_text,")
    if spec.bullet_level:
        lines.append(f"                    level={spec.bullet_level},")
    lines.append("                    source_id=_register(registry, source_refs),")
    lines.append("                )")
    lines.append("            )")
    return "\n".join(lines)


def scenario_propagation_snippets(spec: ChangeSpec) -> str:
    if not scenario_needed(spec):
        return (
            "No function-signature change is required for this section. "
            "`_build_california()` and `_build_arizona()` already receive `state_use`; "
            "the generated block sets `scenario = state_use.scenario`, so "
            "`enabled_scenarios` remains editable in JSON."
        )

    function = SECTIONS[spec.section]
    if spec.section in {"California", "Arizona"}:
        old_sig_hint = "annual, state_use, year, policy, registry"
        new_sig_hint = "annual, state_use, year, scenario, policy, registry"
    else:
        old_sig_hint = "annual, year, policy, registry"
        new_sig_hint = "annual, year, scenario, policy, registry"

    target_call_old = {
        "California": "_build_california(annual, state_use, year, policy, registry)",
        "Arizona": "_build_arizona(annual, state_use, year, policy, registry)",
        "Nevada": "_build_nevada(annual, year, policy, registry)",
        "Mexico": "_build_mexico(annual, year, policy, registry)",
    }[spec.section]
    target_call_new = target_call_old.replace("year, policy", "year, scenario, policy")

    return f"""The v2.1 `{spec.section}` builder does not currently receive the selected scenario. Propagate `scenario` explicitly so this bullet's `enabled_scenarios` setting remains editable in `report_policy.json`.

1. Change `_build_year_section(...)` to accept `scenario: str` before `policy`.
2. In `build_report_context(...)`, change:

```python
_build_year_section(annual, state_use, year, idx, policy, registry)
```

to:

```python
_build_year_section(annual, state_use, year, idx, scenario, policy, registry)
```

3. Change `{function}(...)` so its effective parameter sequence changes from:

```text
{old_sig_hint}
```

to:

```text
{new_sig_hint}
```

4. In `_build_year_section(...)`, change:

```python
{target_call_old}
```

to:

```python
{target_call_new}
```

Do not derive the scenario from the model filename. The selected scenario is already an explicit input to `build_report_context(...)` and should remain the authoritative scenario identifier."""


def json_registration_snippet(spec: ChangeSpec) -> str:
    required = [s.slot_name for s in spec.slots if s.registration == "required"]
    optional = [s.slot_name for s in spec.slots if s.registration == "optional"]
    lines = []
    if required:
        lines.append("Add these names once to `riverware.required_slots`:")
        lines.append("```json")
        lines.append(json.dumps(required, indent=2))
        lines.append("```")
    if optional:
        lines.append("Add these names once to `riverware.optional_slots`:")
        lines.append("```json")
        lines.append(json.dumps(optional, indent=2))
        lines.append("```")
    if not required and not optional:
        lines.append("No new RiverWare slot registration is needed; every requested slot is already registered in the current policy.")
    return "\n\n".join(lines)


def test_fixture_code(spec: ChangeSpec) -> str:
    lines: list[str] = []
    optional_new_or_existing = [s for s in spec.slots if s.is_optional]
    if optional_new_or_existing:
        lines.append("# Optional slots are not created automatically by the existing make_annual() fixture.")
        for slot in optional_new_or_existing:
            vals = slot.test_values_af or [0, 0, 0]
            lines.append(f"slots[{slot.slot_name!r}] = SlotResult(")
            lines.append(f"    slot_name={slot.slot_name!r},")
            lines.append("    unit_name=\"acre-ft\",")
            lines.append(
                "    values={y: Decimal(str(v)) for y, v in zip(years, " + repr(vals) + ")},"
            )
            lines.append("    source_unit_name=\"acre-ft\",")
            lines.append(")")
    for slot in spec.slots:
        if not slot.is_optional:
            vals = slot.test_values_af or [0, 0, 0]
            lines.append(f"set_values({slot.slot_name!r}, {vals!r})")
    return "\n".join(lines) or "# No fixture changes required."


def choose_expected_test(spec: ChangeSpec) -> tuple[int, str, float, str] | None:
    if not spec.test_years or any(len(s.test_values_af) != len(spec.test_years) for s in spec.slots):
        return None
    for i, year in enumerate(spec.test_years):
        # Check only straightforward year applicability. Custom conditions are not evaluated here.
        year_ok = True
        if spec.year_mode == "From year onward":
            year_ok = year >= int(spec.year_value)
        elif spec.year_mode == "Through year":
            year_ok = year <= int(spec.year_value)
        elif spec.year_mode == "Exact years":
            allowed = {int(x.strip()) for x in spec.year_value.split(",") if x.strip()}
            year_ok = year in allowed
        if not year_ok:
            continue
        env = {slot.alias: slot.test_values_af[i] for slot in spec.slots}
        try:
            computed = _safe_eval_expr(spec.formula, env)
        except Exception:
            return None
        condition_ok = {
            "Always": True,
            "Positive (> 0)": computed > 0,
            "Non-zero (!= 0)": computed != 0,
            "Negative (< 0)": computed < 0,
            "Absolute value above threshold": abs(computed) > float(spec.threshold_af or 0),
        }.get(spec.display_condition)
        if condition_ok is False:
            continue
        if condition_ok is None:  # custom condition: cannot safely predict
            return None
        display = format_display(computed, spec.display_units, spec.decimals)
        scenario = spec.scenarios[0]
        expected = spec.bullet_text.format(value=display, year=year, scenario=scenario)
        return year, scenario, computed, expected
    return None


def test_code(spec: ChangeSpec) -> str:
    expected = choose_expected_test(spec)
    section_lookup = {
        "California": 'next(s for s in ctx.years[year_index].states if s.heading == "California")',
        "Arizona": 'next(s for s in ctx.years[year_index].states if s.heading == "Arizona")',
        "Nevada": 'next(s for s in ctx.years[year_index].states if s.heading == "Nevada")',
        "Mexico": "ctx.years[year_index]",
    }
    if expected is None:
        return f'''def test_{spec.rule_key}_bullet_and_sources():
    policy = load_policy()
    annual = make_annual()
    state_use = make_state_use({spec.scenarios[0]!r})

    # TODO: populate the requested slot values in make_annual() or directly here.
    ctx = build_report_context(annual, state_use, {spec.scenarios[0]!r}, "Test Month 2026", policy)

    # TODO: choose the report year that satisfies the configured conditions.
    year_index = 0
    target = {section_lookup[spec.section]}
    bullets = target.mexico_bullets if {spec.section!r} == "Mexico" else target.bullets

    # Assert the exact approved wording, not only a substring.
    expected_text = "TODO: exact expected bullet text"
    bullet = next(b for b in bullets if b.text == expected_text)
    refs = ctx.source_map[bullet.source_id]
    assert {{r["locator"] for r in refs if r["kind"] == "RiverWare"}} == {set(s.slot_name for s in spec.slots)!r}
'''

    year, scenario, _computed, expected_text = expected
    year_index = spec.test_years.index(year)
    test_fn = re.sub(r"[^a-z0-9_]+", "_", spec.rule_key.lower())
    rw_set = {s.slot_name for s in spec.slots}
    return f'''def test_{test_fn}_bullet_and_sources():
    policy = load_policy()
    annual = make_annual()
    state_use = make_state_use({scenario!r})
    ctx = build_report_context(annual, state_use, {scenario!r}, "Test Month 2026", policy)

    year_index = {year_index}  # {year}
    target = {section_lookup[spec.section]}
    bullets = target.mexico_bullets if {spec.section!r} == "Mexico" else target.bullets
    expected_text = {expected_text!r}
    bullet = next(b for b in bullets if b.text == expected_text)

    refs = ctx.source_map[bullet.source_id]
    rw_locators = {{r["locator"] for r in refs if r["kind"] == "RiverWare"}}
    assert rw_locators == {rw_set!r}
'''


def inventory_files(project_root: Path | None) -> str:
    if not project_root:
        return "Project files were not available when this plan was generated. Verify filenames against the target release before implementing."
    files = [
        "config/report_policy.json",
        "policy.py",
        "report_builder.py",
        "riverware_annual.py",
        "report_models.py",
        "templates/report_preview.html",
        "docx_writer.py",
        "tests/test_report_builder.py",
        "tests/test_policy.py",
        "tests/test_real_model_optional.py",
        "run_tests.bat",
    ]
    found = [f for f in files if (project_root / f).exists()]
    return "Detected in project: " + ", ".join(f"`{f}`" for f in found) + "."


def render_plan(spec: ChangeSpec, project_root: Path | None) -> str:
    function, _collection = insertion_target(spec)
    function_line = locate_function_line(project_root, function)
    anchor_lines = locate_anchor_lines(project_root, spec.placement_anchor)
    needs_scenario = scenario_needed(spec)
    new_slots = [s for s in spec.slots if s.is_new]
    nonannual = [s for s in spec.slots if not s.annual_volume_confirmed]

    placement_detail = f"`{spec.placement}`"
    if spec.placement_anchor:
        placement_detail += f" relative to the existing bullet/code containing `{spec.placement_anchor}`"
        if anchor_lines:
            placement_detail += f" (current candidate line(s): {', '.join(map(str, anchor_lines))})"
    elif function_line:
        placement_detail += f" inside `{function}()` (currently begins near line {function_line})"

    warnings: list[str] = []
    if nonannual:
        warnings.append(
            "At least one source was not confirmed as a normal annual volume slot. Do not implement only the generated registration/builder code until `riverware_annual.py` is reviewed for the source representation/timestep/unit."
        )
    if spec.display_condition == "Custom Python condition":
        warnings.append(
            "The display condition is custom Python. The planner validates its syntax/names but does not prove the business meaning. Add branch-specific tests."
        )
    if spec.formula.strip() not in {s.alias for s in spec.slots}:
        warnings.append(
            "This is a calculation change. Have a second reviewer verify formula signs, units, and rounding against the approved business rule."
        )
    if any(s.is_optional for s in spec.slots) and spec.missing_optional_behavior == "Treat missing optional inputs as zero":
        warnings.append(
            "One or more optional sources are treated as zero when absent. Confirm that 'missing means zero' is an approved interpretation; otherwise suppress the bullet instead."
        )

    rule_json = json.dumps({spec.rule_key: policy_rule_object(spec)}, indent=2)
    builder_block = render_builder_block(spec)
    fixture = test_fixture_code(spec)
    regression = test_code(spec)

    slot_rows = "\n".join(
        f"| `{s.alias}` | `{s.slot_name}` | {s.registration} | {'Yes' if s.annual_volume_confirmed else '**No / verify parser**'} | {s.description or '-'} |"
        for s in spec.slots
    )
    warning_text = "\n".join(f"- **Warning:** {w}" for w in warnings) if warnings else "- No special planner warnings beyond the normal review/test requirements."

    return fr"""# Implementation Plan - {spec.title}

**Generated by:** `{Path(__file__).name}`  
**Generated:** {datetime.now().isoformat(timespec='seconds')}  
**Target application:** {APP_NAME} v2.1 architecture  
**Current policy version detected:** `{spec.policy_version_current}`  
**Proposed policy version:** `{spec.policy_version_proposed}`

> This is a **change plan, not an automatic patch**. Review every generated snippet against the approved business rule and a representative RiverWare model before modifying the operational application.

## 1. Requested change

- **Purpose:** {spec.formula_description or spec.title}
- **Formula (inputs are acre-ft):** `{spec.formula}`
- **Target section:** {spec.section}
- **Builder function:** `{function}()`
- **Placement:** {placement_detail}
- **Scenarios:** {', '.join(spec.scenarios)}
- **Year condition:** {spec.year_mode}{(': ' + spec.year_value) if spec.year_value else ''}
- **Display condition:** {spec.display_condition}{(' (' + str(spec.threshold_af) + ' AF)') if spec.threshold_af is not None else ''}
- **Bullet level:** {spec.bullet_level}
- **Display format:** {spec.decimals} decimal place(s), {spec.display_units}
- **Bullet template:** `{spec.bullet_text}`
- **Optional-source missing behavior:** {spec.missing_optional_behavior}
- **Maintainer notes:** {spec.notes or '-'}

### RiverWare inputs

| Alias used in code | Fully qualified slot | Registration | Confirmed normal annual volume? | Purpose |
|---|---|---|---|---|
{slot_rows}

## 2. Planner warnings / review items

{warning_text}

## 3. Files that should change

**Expected changes:**

1. `config/report_policy.json` - register any new RiverWare slot(s), add the bullet applicability/configuration rule, and increment `policy_version`.
2. `report_builder.py` - retrieve the source slot(s), perform the calculation, apply conditions, create the bullet, and register Source references.
3. `tests/test_report_builder.py` - add deterministic slot values and assertions for wording, applicability, and Source tracing.
{('4. `report_builder.py` call signatures/calls - propagate the selected scenario into the target builder because this rule is scenario-specific.\n5.' if needs_scenario else '4.')} `CHANGELOG.md` - document the approved report-rule/calculation change.

**Usually unchanged for this type of change:**

- `riverware_annual.py` - unchanged **only if** every new source is a supported annual volume slot.
- `report_models.py` - no change for a normal bullet in an existing section.
- `templates/report_preview.html` - no change; it renders canonical `Bullet` objects.
- `docx_writer.py` - no change; it renders the same canonical `Bullet` objects.
- `reference_report.docx` - no change unless Word styling itself is being changed.

{inventory_files(project_root)}

## 4. Update `config/report_policy.json`

### 4.1 Increment the policy version

Change:

```json
"policy_version": "{spec.policy_version_current}"
```

to:

```json
"policy_version": "{spec.policy_version_proposed}"
```

Use your organization's versioning convention if a different release number is required.

### 4.2 Register new RiverWare slot(s)

{json_registration_snippet(spec)}

Do **not** put a slot in both `required_slots` and `optional_slots`.

### 4.3 Add a rule describing applicability/display policy

Add the following object inside `report.rules` (remember the comma separating it from adjacent JSON members):

```json
{rule_json}
```

This rule intentionally stores **applicability and presentation policy**, not executable arithmetic. The formula stays in Python.

Validate the JSON before running the app:

```bat
.venv\Scripts\python.exe -m json.tool config\report_policy.json > NUL
```

## 5. Scenario propagation, if required

{scenario_propagation_snippets(spec)}

## 6. Add the calculation and bullet in `report_builder.py`

Insert the following block in `{function}()` at the approved location described in Section 1. The source values returned by `_slot(...)` are already normalized to **acre-ft**.

```python
{builder_block}
```

### Why the Source registration is written this way

- `_rw_ref(...)` records each underlying RiverWare slot/year/model source.
- `_calc_ref(...)` is included when the displayed value is a true arithmetic calculation rather than a direct slot display.
- `_register(...)` creates the `source_id` used by the preview's **Source** control.
- The resulting `Bullet` is part of the canonical report model, so both HTML and Word use the same content.

## 7. Update `tests/test_report_builder.py`

### 7.1 Add deterministic source values to the fixture

Use the following pattern in `make_annual()` after `slots` and `set_values(...)` are available:

```python
{fixture}
```

The values above come from the test values entered into the planner. They are not production assumptions.

### 7.2 Add a regression test for the bullet and Source trace

```python
{regression}
```

Also add separate tests when applicable for:

- a scenario where the bullet must **not** appear;
- a year before/after the effective-year boundary;
- zero/negative values if those affect display behavior;
- a missing optional slot if the slot is optional;
- rounding at a boundary that could change the displayed number.

## 8. Optional-slot test

If any new source is optional, add a test that removes it from `annual.slots` and verifies the approved missing behavior. For the recommended "suppress bullet" behavior:

```python
annual = make_annual()
annual.slots.pop("OPTIONAL_SLOT_NAME", None)
ctx = build_report_context(annual, make_state_use("Most"), "Most", "Test Month 2026", policy)
# Assert that the new bullet is absent and report generation still succeeds.
```

If a source is **required**, test missing-source behavior at the service/parser validation level instead of silently treating it as zero.

## 9. Validate with the automated suite

From the project folder:

```bat
run_tests.bat
```

Or directly:

```bat
.venv\Scripts\python.exe -m pytest -q
```

Do not accept an existing test failure merely because the new bullet appears correctly.

## 10. Validate against a real RiverWare model

For a newly registered slot, use a representative model that is supposed to contain the slot.

```bat
set WU_TEST_MODEL=C:\Path\To\RepresentativeModel.mdl.gz
.venv\Scripts\python.exe -m pytest -q -m integration
```

Then run the Flask app and verify:

1. preflight finds all required sources;
2. no unexpected unit/timestep warning appears;
3. the bullet appears only for the configured scenario/year/value conditions;
4. the displayed value agrees with a manual calculation from RiverWare;
5. **Source** lists all `{len(spec.slots)}` RiverWare source slot(s) and the calculation reference when applicable; and
6. the Word report contains exactly the same bullet text as the preview.

## 11. If a source is not a supported annual volume slot

The generated builder code assumes `_slot(...)` can retrieve the value in acre-ft. If the source is monthly, table/periodic, scalar with non-volume meaning, uses an unsupported RiverWare serialization, or uses a new physical unit, stop here and review `riverware_annual.py`.

That change normally requires:

1. a parser/conversion enhancement in `riverware_annual.py`;
2. parser-level tests for the new representation/unit;
3. policy `accepted_volume_units` changes only if the parser actually implements the conversion;
4. real-model validation; and
5. then the report-builder changes above.

Never add a unit name to JSON merely to bypass validation when no conversion has been implemented.

## 12. Changelog entry

Add an entry similar to:

```markdown
- Policy {spec.policy_version_proposed}: {spec.title}. Added/used RiverWare source(s) {', '.join(f'`{s.slot_name}`' for s in spec.slots)}; calculation `{spec.formula}`; bullet applies to {', '.join(spec.scenarios)} in the {spec.section} section.
```

## 13. Final operational review checklist

- [ ] Business rule/formula independently reviewed.
- [ ] Slot names verified directly in RiverWare.
- [ ] Required vs optional classification approved.
- [ ] All input units/timesteps confirmed.
- [ ] `report_policy.json` is valid JSON.
- [ ] `policy_version` incremented.
- [ ] Formula operates on acre-ft values and converts only for display.
- [ ] Sign convention verified (especially reductions/conservation values).
- [ ] Rounding verified against the approved report convention.
- [ ] Scenario condition tested both true and false.
- [ ] Year condition tested at its boundary.
- [ ] Value/display condition tested.
- [ ] Source panel shows all contributing sources.
- [ ] Automated tests pass.
- [ ] Real-model integration test passes.
- [ ] HTML preview reviewed.
- [ ] Generated Word report reviewed.
- [ ] `CHANGELOG.md` updated.
- [ ] Previous known-good release retained for rollback/reproducibility.

## 14. Important architecture boundary

Use JSON to describe **changeable policy**: source registration, applicability, labels, thresholds, and configuration. Keep **executable calculations** in Python. Do not turn `report_policy.json` into an expression engine and do not duplicate the calculation in the HTML or Word renderer.

"""


def collect_interactive(policy: dict[str, Any]) -> ChangeSpec:
    print("\nWU Assumptions Generator - Report Change Planner")
    print("=" * 52)
    print("This utility creates an implementation plan. It does not edit the app.\n")

    current_version = str(policy.get("policy_version", "unknown"))
    proposed_default = bump_policy_version(current_version) if current_version != "unknown" else datetime.now().strftime("%Y.%m.1")
    title = ask("Short name for this report change (e.g. Add Yuma conservation bullet)")
    rule_default = slugify(title)
    rule_key = ask("Policy rule key (letters/numbers/underscore)", rule_default)
    if not re.fullmatch(r"[a-z][a-z0-9_]*", rule_key):
        raise PlannerError("Policy rule key must start with a lowercase letter and contain only lowercase letters, digits, and underscores.")
    proposed_version = ask("Proposed policy version", proposed_default)

    print("\nRiverWare inputs")
    count = ask_int("How many RiverWare slots feed this value?", 1, 1, 20)
    slots: list[SlotSpec] = []
    used_aliases: set[str] = set()
    for i in range(count):
        print(f"\nSlot {i + 1} of {count}")
        slot_name = ask("Fully qualified RiverWare slot name (Object.Slot)")
        detected = current_slot_status(policy, slot_name)
        if detected:
            print(f"  Detected in current policy as: {detected.replace('-', ' ')}")
            registration = detected
        else:
            registration = ask_choice(
                "This slot is new. Should it be required or optional?",
                ["required", "optional"],
                "required",
            )
        alias_default = suggest_alias(slot_name, used_aliases)
        alias = ask("Short Python alias for the slot", alias_default)
        if not safe_identifier(alias):
            raise PlannerError(f"'{alias}' is not a valid Python identifier.")
        if alias in used_aliases:
            raise PlannerError(f"Duplicate alias '{alias}'.")
        used_aliases.add(alias)
        description = ask("What does this slot represent?", allow_blank=True)
        annual = ask_yes_no("Is this a normal annual volume slot supported by the current parser?", True)
        slots.append(SlotSpec(alias, slot_name, registration, description, annual))

    aliases = {s.alias for s in slots}
    default_formula = slots[0].alias if len(slots) == 1 else " + ".join(s.alias for s in slots)
    print("\nCalculation")
    print("Use the aliases above. Inputs are acre-ft after RiverWare extraction.")
    print("Allowed arithmetic: + - * / // % **, parentheses, abs(), min(), max().")
    formula = ask("Formula that returns the report value in acre-ft", default_formula)
    validate_formula(formula, aliases)
    formula_description = ask("Plain-language description of the calculation/business meaning")

    optional_slots = [s for s in slots if s.is_optional]
    if optional_slots:
        missing_behavior = ask_choice(
            "If an optional input is absent",
            ["Suppress bullet if any optional input is missing", "Treat missing optional inputs as zero"],
            "Suppress bullet if any optional input is missing",
        )
    else:
        missing_behavior = "Not applicable - all inputs are required"

    print("\nPlacement")
    section = ask_choice("Report section", SECTIONS.keys(), "California")
    placement = ask_choice(
        "Where in that section should the bullet appear?",
        ["End of section", "Beginning of section", "After existing bullet", "Before existing bullet"],
        "End of section",
    )
    anchor = ""
    if placement in {"After existing bullet", "Before existing bullet"}:
        anchor = ask("Unique text from the existing bullet to use as the placement anchor")

    print("\nApplicability")
    scenario_raw = ask("Scenarios (comma-separated Most, Min, Max; enter ALL for all)", "ALL")
    if scenario_raw.strip().upper() == "ALL":
        scenarios = list(KNOWN_SCENARIOS)
    else:
        entered = [x.strip().title() for x in scenario_raw.split(",") if x.strip()]
        bad = [x for x in entered if x not in KNOWN_SCENARIOS]
        if bad or not entered:
            raise PlannerError("Scenarios must be Most, Min, Max, or ALL.")
        scenarios = list(dict.fromkeys(entered))

    year_mode = ask_choice(
        "Year applicability",
        ["All years", "From year onward", "Through year", "Exact years"],
        "All years",
    )
    year_value = ""
    if year_mode in {"From year onward", "Through year"}:
        year_value = str(ask_int("Boundary year", datetime.now().year, 1900, 2200))
    elif year_mode == "Exact years":
        raw_years = ask("Comma-separated calendar years (e.g. 2026,2027)")
        try:
            years = [int(x.strip()) for x in raw_years.split(",") if x.strip()]
        except ValueError as exc:
            raise PlannerError("Exact years must be comma-separated integers.") from exc
        if not years:
            raise PlannerError("Enter at least one exact year.")
        year_value = ",".join(str(y) for y in years)

    display_condition = ask_choice(
        "When should the bullet display?",
        ["Always", "Positive (> 0)", "Non-zero (!= 0)", "Negative (< 0)", "Absolute value above threshold", "Custom Python condition"],
        "Positive (> 0)",
    )
    threshold: float | None = None
    custom_condition = ""
    if display_condition == "Absolute value above threshold":
        threshold = float(ask("Threshold in acre-ft", "0").replace(",", ""))
    elif display_condition == "Custom Python condition":
        custom_condition = ask(
            "Python boolean expression using computed_af, year, scenario, and/or slot aliases"
        )
        validate_condition(custom_condition, aliases)

    print("\nBullet wording / display")
    display_units = ask_choice("Displayed units", FORMAT_UNITS, "kaf")
    decimals = ask_int("Number of decimal places", 1 if display_units != "af" else 0, 0, 6)
    print("Use {value} where the formatted number should appear. Optional placeholders: {year}, {scenario}.")
    bullet_text = ask("Exact bullet wording", f"{title} of {{value}} {display_units}")
    validate_bullet_template(bullet_text)
    if "{value}" not in bullet_text:
        if not ask_yes_no("The wording does not contain {value}. Continue anyway?", False):
            raise PlannerError("Bullet wording was not accepted.")
    bullet_level = ask_choice("Bullet level", ["0", "1"], "0")
    source_note = ask(
        "Optional Source-panel calculation note (blank = generated description)",
        allow_blank=True,
    )

    print("\nRegression-test inputs")
    years_raw = ask("Three test years, comma-separated", "2026,2027,2028")
    try:
        test_years = [int(x.strip()) for x in years_raw.split(",") if x.strip()]
    except ValueError as exc:
        raise PlannerError("Test years must be comma-separated integers.") from exc
    if len(test_years) != 3:
        raise PlannerError("Enter exactly three test years to match the current report fixture.")
    for slot in slots:
        slot.test_values_af = ask_float_list(
            f"Test values for {slot.alias} in AF for {', '.join(map(str, test_years))} (semicolon-separated)",
            len(test_years),
        )

    notes = ask("Any additional implementation/reviewer notes?", allow_blank=True)

    spec = ChangeSpec(
        title=title,
        rule_key=rule_key,
        policy_version_current=current_version,
        policy_version_proposed=proposed_version,
        slots=slots,
        formula=formula,
        formula_description=formula_description,
        missing_optional_behavior=missing_behavior,
        section=section,
        placement=placement,
        placement_anchor=anchor,
        scenarios=scenarios,
        year_mode=year_mode,
        year_value=year_value,
        display_condition=display_condition,
        threshold_af=threshold,
        custom_condition=custom_condition,
        bullet_text=bullet_text,
        bullet_level=int(bullet_level),
        display_units=display_units,
        decimals=decimals,
        source_note=source_note,
        test_years=test_years,
        notes=notes,
    )
    return spec


def spec_from_json(data: dict[str, Any]) -> ChangeSpec:
    data = dict(data)
    data["slots"] = [SlotSpec(**item) for item in data.get("slots", [])]
    spec = ChangeSpec(**data)
    aliases = {s.alias for s in spec.slots}
    validate_formula(spec.formula, aliases)
    validate_condition(spec.custom_condition, aliases)
    validate_bullet_template(spec.bullet_text)
    return spec


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Interview a maintainer and generate a detailed WU Assumptions Generator report-change implementation plan."
    )
    parser.add_argument(
        "--project-root",
        type=Path,
        help="Path to the WU Assumptions Generator project root. If omitted, the script searches upward from the current directory and its own location.",
    )
    parser.add_argument(
        "--answers",
        type=Path,
        help="Load a previously saved questionnaire JSON instead of prompting interactively.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        help="Markdown plan output path. Default: change_plans/<rule_key>_implementation_plan.md under the project root or current directory.",
    )
    parser.add_argument(
        "--save-answers",
        type=Path,
        help="Optional path for saving the completed questionnaire as JSON for reuse/review.",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    try:
        if args.project_root:
            project_root = args.project_root.resolve()
            if not (project_root / "report_builder.py").exists():
                raise PlannerError(f"{project_root} does not look like the application project root.")
        else:
            project_root = find_project_root(Path.cwd()) or find_project_root(Path(__file__).resolve())

        policy = load_project_policy(project_root)
        if project_root:
            print(f"Detected project root: {project_root}")
            if policy.get("policy_version"):
                print(f"Detected policy version: {policy['policy_version']}")
        else:
            print("WARNING: Could not detect the application project. The plan will use v2.1 architecture assumptions.")

        if args.answers:
            try:
                data = json.loads(args.answers.read_text(encoding="utf-8"))
            except (OSError, json.JSONDecodeError) as exc:
                raise PlannerError(f"Could not read answers file {args.answers}: {exc}") from exc
            spec = spec_from_json(data)
        else:
            spec = collect_interactive(policy)

        # Reconcile loaded answers with the current project's registration state.
        if policy:
            for slot in spec.slots:
                detected = current_slot_status(policy, slot.slot_name)
                if detected:
                    slot.registration = detected

        base = project_root or Path.cwd()
        default_output = base / "change_plans" / f"{spec.rule_key}_implementation_plan.md"
        output = (args.output or default_output).resolve()
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(render_plan(spec, project_root), encoding="utf-8")

        answers_path = args.save_answers
        if answers_path:
            answers_path = answers_path.resolve()
            answers_path.parent.mkdir(parents=True, exist_ok=True)
            answers_path.write_text(json.dumps(asdict(spec), indent=2), encoding="utf-8")

        print("\nPlan generated successfully.")
        print(f"Markdown plan: {output}")
        if answers_path:
            print(f"Questionnaire JSON: {answers_path}")
        print("\nReview the generated plan before editing the operational application.")
        return 0
    except PlannerError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2
    except KeyboardInterrupt:
        print("\nCancelled.", file=sys.stderr)
        return 130


if __name__ == "__main__":
    raise SystemExit(main())
