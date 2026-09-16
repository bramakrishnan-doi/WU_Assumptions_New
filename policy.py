"""Versioned report-policy loader.

Policy that changes with operating guidance, scenario conventions, workbook
layout, or report wording belongs in config/report_policy.json. Processing,
parsing, and rendering mechanics stay in Python.
"""
from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Mapping

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_POLICY_PATH = BASE_DIR / "config" / "report_policy.json"


class PolicyError(RuntimeError):
    pass


@dataclass(frozen=True)
class ScenarioPolicy:
    key: str
    display_name: str
    label_column: str
    value_column: str
    notes_column: str


class ReportPolicy:
    def __init__(self, raw: Mapping[str, Any], path: Path):
        self.raw = dict(raw)
        self.path = path
        self.version = str(raw.get("policy_version", "")).strip()
        if not self.version:
            raise PolicyError("Policy file is missing policy_version.")
        scenarios = raw.get("scenarios")
        if not isinstance(scenarios, dict) or not scenarios:
            raise PolicyError("Policy file must define scenarios.")
        self._scenarios: Dict[str, ScenarioPolicy] = {}
        for key, cfg in scenarios.items():
            excel = cfg.get("excel", {})
            try:
                self._scenarios[key] = ScenarioPolicy(
                    key=key,
                    display_name=str(cfg["display_name"]),
                    label_column=str(excel["label_column"]),
                    value_column=str(excel["value_column"]),
                    notes_column=str(excel["notes_column"]),
                )
            except KeyError as exc:
                raise PolicyError(f"Scenario '{key}' is missing {exc.args[0]} in policy.") from exc

        riverware = raw.get("riverware", {})
        self.required_slots = tuple(str(v).strip() for v in riverware.get("required_slots", []) if str(v).strip())
        optional = tuple(str(v).strip() for v in riverware.get("optional_slots", []) if str(v).strip())
        self.optional_slots = frozenset(optional)
        self.accepted_volume_units = tuple(str(v).strip() for v in riverware.get("accepted_volume_units", []) if str(v).strip())
        if not self.required_slots:
            raise PolicyError("Policy file does not define RiverWare required_slots.")
        if len(set(self.required_slots)) != len(self.required_slots):
            raise PolicyError("Policy file contains duplicate RiverWare required_slots.")
        overlap = set(self.required_slots) & self.optional_slots
        if overlap:
            raise PolicyError(
                "RiverWare slots cannot be both required and optional: " + ", ".join(sorted(overlap))
            )
        if not self.accepted_volume_units:
            raise PolicyError("Policy file must define at least one accepted_volume_units value.")
        self.target_year_count = int(riverware.get("target_year_count", 3))
        if self.target_year_count != 3:
            raise PolicyError("This application currently requires target_year_count = 3.")

    @property
    def scenario_keys(self) -> tuple[str, ...]:
        return tuple(self._scenarios.keys())

    @property
    def all_slots(self) -> tuple[str, ...]:
        """Slots to parse, preserving required-slot order and then optional slots."""
        return (*self.required_slots, *(s for s in sorted(self.optional_slots) if s not in self.required_slots))

    def scenario(self, key: str) -> ScenarioPolicy:
        try:
            return self._scenarios[key]
        except KeyError as exc:
            raise PolicyError(f"Unknown scenario '{key}'.") from exc

    @property
    def report(self) -> Mapping[str, Any]:
        return self.raw["report"]

    @property
    def excel(self) -> Mapping[str, Any]:
        return self.raw["excel"]

    def disclaimer_for_year(self, year: int) -> str:
        rules = sorted(
            self.report.get("disclaimers", []),
            key=lambda item: int(item.get("effective_from_year", 0)),
            reverse=True,
        )
        for item in rules:
            if year >= int(item.get("effective_from_year", 0)):
                return str(item.get("text", ""))
        return ""

    def mexico_shortage_label(self, year: int) -> str:
        rules = sorted(
            self.report.get("rules", {}).get("mexico_shortage_wording", []),
            key=lambda item: int(item.get("effective_from_year", 0)),
            reverse=True,
        )
        for item in rules:
            if year >= int(item.get("effective_from_year", 0)):
                return str(item.get("label", "Shortage volume"))
        return "Shortage volume"


    def nevada_shortage_label(self, year: int) -> str:
        rules = sorted(
            self.report.get("rules", {}).get("nevada_shortage_wording", []),
            key=lambda item: int(item.get("effective_from_year", 0)),
            reverse=True,
        )
        for item in rules:
            if year >= int(item.get("effective_from_year", 0)):
                return str(item.get("label", "Shortage volume"))
        return "Shortage volume"

    def nevada_system_conservation_rule(self, year: int) -> Mapping[str, Any]:
        rules = sorted(
            self.report.get("rules", {}).get("nevada_system_conservation_wording", []),
            key=lambda item: int(item.get("effective_from_year", 0)),
            reverse=True,
        )
        for item in rules:
            if year >= int(item.get("effective_from_year", 0)):
                return item
        return {"label": "Total System Conservation of", "decompose": True}

    def az_reduction_override_kaf(self, year: int) -> float | None:
        cfg = self.report.get("rules", {}).get("az_water_use_reduction_override")
        if not cfg:
            return None
        if year >= int(cfg.get("effective_from_year", 9999)):
            return float(cfg["display_kaf"])
        return None

    def shows_conservation_summary(self, scenario: str) -> bool:
        return scenario in set(self.report.get("conservation_summary_scenarios", []))

    @property
    def default_enabled_scenarios(self) -> tuple[str, ...]:
        configured = self.report.get("default_enabled_scenarios", ["Most", "Min"])
        result = tuple(str(v) for v in configured if str(v) in self._scenarios)
        return result or (self.scenario_keys[0],)

    def shows_powell_release_subtitle(self, scenario: str) -> bool:
        configured = self.report.get("powell_release_subtitle_scenarios", self.scenario_keys)
        return scenario in {str(v) for v in configured}

    def rolls_california_conservation_into_total(self, note: str | None) -> bool:
        if not note:
            return False
        cfg = self.report.get("rules", {}).get("california_system_conservation_rollup", {})
        labels = cfg.get("year_sheet_note_labels", [])
        normalized_note = " ".join(str(note).lower().split())
        return any(" ".join(str(label).lower().split()) in normalized_note for label in labels if str(label).strip())


def load_policy(path: str | Path | None = None) -> ReportPolicy:
    policy_path = Path(path) if path else DEFAULT_POLICY_PATH
    try:
        with policy_path.open("r", encoding="utf-8") as f:
            raw = json.load(f)
    except (OSError, json.JSONDecodeError) as exc:
        raise PolicyError(f"Could not load report policy from {policy_path}: {exc}") from exc
    return ReportPolicy(raw, policy_path)
