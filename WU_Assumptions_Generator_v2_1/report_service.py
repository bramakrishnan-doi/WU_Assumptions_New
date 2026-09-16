"""Application service layer for validating inputs and preparing report batches."""
from __future__ import annotations

import re
import time
import uuid
from dataclasses import dataclass, field
from typing import BinaryIO, Dict, Iterable, Mapping

from excel_source import ExcelValidationError, StateUseData, load_state_use_workbook
from policy import ReportPolicy
from riverware_annual import AnnualExtractionResult, RiverwareValidationError, load_annual_slots


class BatchValidationError(Exception):
    def __init__(self, errors: Iterable[str], warnings: Iterable[str] = ()):
        self.errors = list(errors)
        self.warnings = list(warnings)
        super().__init__("; ".join(self.errors))


@dataclass
class PreparedBatch:
    token: str
    created_at: float
    policy_version: str
    workbook_filename: str
    start_year: int
    scenarios: tuple[str, ...]
    annual_by_scenario: Dict[str, AnnualExtractionResult]
    state_use_by_scenario: Dict[str, StateUseData]
    warnings: list[str] = field(default_factory=list)

    @property
    def years(self) -> tuple[int, ...]:
        has_third = all(self.state_use_by_scenario[s].has_year3_sheet for s in self.scenarios)
        return tuple(range(self.start_year, self.start_year + (3 if has_third else 2)))

    def preflight_dict(self, policy: ReportPolicy) -> dict:
        scenario_rows = []
        for scenario in self.scenarios:
            annual = self.annual_by_scenario[scenario]
            state = self.state_use_by_scenario[scenario]
            scenario_rows.append(
                {
                    "scenario": scenario,
                    "scenario_label": policy.scenario(scenario).display_name,
                    "model_filename": annual.model_filename,
                    "model_run_start": annual.run_start.strftime("%Y-%m-%d %H:%M"),
                    "model_run_finish": annual.run_finish.strftime("%Y-%m-%d %H:%M"),
                    "report_years": list(self.years),
                    "required_slots_found": len(policy.required_slots),
                    "required_slots_total": len(policy.required_slots),
                    "warnings": list(dict.fromkeys([*annual.warnings, *state.warnings])),
                    "conservation_summary_included": policy.shows_conservation_summary(scenario),
                }
            )
        return {
            "token": self.token,
            "policy_version": self.policy_version,
            "workbook_filename": self.workbook_filename,
            "start_year": self.start_year,
            "report_years": list(self.years),
            "scenarios": scenario_rows,
            "warnings": list(dict.fromkeys(self.warnings)),
        }


_SCENARIO_TOKEN = re.compile(r"(?:^|[\s_.\-])(MOST|MIN|MAX)(?:[\s_.\-]|$)", re.IGNORECASE)


def _filename_scenario_issue(filename: str, assigned: str) -> tuple[str | None, str | None]:
    """Return (error, warning) for a clearly identifiable filename scenario."""
    detected = {m.group(1).title() for m in _SCENARIO_TOKEN.finditer(filename)}
    if len(detected) > 1:
        return (
            f"{assigned}: model filename '{filename}' contains multiple scenario labels "
            f"({', '.join(sorted(detected))}). Rename or select the correct model.",
            None,
        )
    if detected:
        actual = next(iter(detected))
        if actual != assigned:
            return (
                f"{assigned}: model filename '{filename}' appears to be a {actual} scenario model.",
                None,
            )
        return None, None
    return (
        None,
        f"{assigned}: could not infer a scenario label from model filename '{filename}'. "
        "The model will still be used for the selected scenario; verify the assignment before generating.",
    )


def prepare_batch(
    model_uploads: Mapping[str, tuple[str, bytes | BinaryIO]],
    workbook_filename: str,
    workbook_bytes: bytes,
    policy: ReportPolicy,
) -> PreparedBatch:
    """Parse and validate all selected scenarios and the workbook exactly once."""
    if not model_uploads:
        raise BatchValidationError(["Select and upload at least one RiverWare scenario model."])
    if not workbook_filename or not workbook_bytes:
        raise BatchValidationError(["Upload the current Projected State Use .xlsx workbook."])

    errors: list[str] = []
    warnings: list[str] = []
    annual_by_scenario: Dict[str, AnnualExtractionResult] = {}

    ordered_scenarios = tuple(s for s in policy.scenario_keys if s in model_uploads)
    unknown = [s for s in model_uploads if s not in policy.scenario_keys]
    if unknown:
        errors.append("Unknown scenario(s): " + ", ".join(unknown))

    for scenario in ordered_scenarios:
        filename, model_source = model_uploads[scenario]
        model_bytes = model_source if isinstance(model_source, bytes) else model_source.read()
        err, warn = _filename_scenario_issue(filename, scenario)
        if err:
            errors.append(err)
        if warn:
            warnings.append(warn)
        try:
            annual_by_scenario[scenario] = load_annual_slots(
                model_bytes,
                filename,
                policy.all_slots,
                policy.optional_slots,
                policy.accepted_volume_units,
            )
        except RiverwareValidationError as exc:
            errors.extend(f"{scenario}: {message}" for message in exc.errors)

    if errors:
        raise BatchValidationError(errors, warnings)

    starts = {a.start_year for a in annual_by_scenario.values()}
    if len(starts) != 1:
        details = ", ".join(f"{s}={a.start_year}" for s, a in annual_by_scenario.items())
        raise BatchValidationError(
            [f"Selected scenario models do not have the same study start year ({details})."],
            warnings,
        )
    start_year = next(iter(starts))

    try:
        state_use_by_scenario = load_state_use_workbook(
            workbook_bytes,
            workbook_filename,
            start_year,
            ordered_scenarios,
            policy,
        )
    except ExcelValidationError as exc:
        raise BatchValidationError(exc.errors, warnings) from exc

    for scenario in ordered_scenarios:
        annual = annual_by_scenario[scenario]
        state = state_use_by_scenario[scenario]
        warnings.extend(annual.warnings)
        warnings.extend(state.warnings)

    return PreparedBatch(
        token=uuid.uuid4().hex,
        created_at=time.time(),
        policy_version=policy.version,
        workbook_filename=workbook_filename,
        start_year=start_year,
        scenarios=ordered_scenarios,
        annual_by_scenario=annual_by_scenario,
        state_use_by_scenario=state_use_by_scenario,
        warnings=list(dict.fromkeys(warnings)),
    )


class PreparedBatchCache:
    """Small in-memory cache so generation does not re-parse large uploads."""

    def __init__(self, ttl_seconds: int = 30 * 60, max_entries: int = 5):
        self.ttl_seconds = ttl_seconds
        self.max_entries = max_entries
        self._items: Dict[str, PreparedBatch] = {}

    def _purge(self) -> None:
        now = time.time()
        expired = [k for k, v in self._items.items() if now - v.created_at > self.ttl_seconds]
        for key in expired:
            self._items.pop(key, None)
        while len(self._items) > self.max_entries:
            oldest = min(self._items.items(), key=lambda kv: kv[1].created_at)[0]
            self._items.pop(oldest, None)

    def put(self, prepared: PreparedBatch) -> None:
        self._purge()
        self._items[prepared.token] = prepared
        self._purge()

    def get(self, token: str) -> PreparedBatch | None:
        self._purge()
        return self._items.get(token)

    def clear(self, token: str) -> None:
        self._items.pop(token, None)
