"""Canonical report data model shared by HTML preview and Word rendering."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Optional


@dataclass(frozen=True)
class SourceRef:
    kind: str
    source_name: str
    locator: str
    year: Optional[int] = None
    raw_value: Optional[str] = None
    units: Optional[str] = None
    note: Optional[str] = None

    def to_dict(self) -> dict:
        return {
            "kind": self.kind,
            "source_name": self.source_name,
            "locator": self.locator,
            "year": self.year,
            "raw_value": self.raw_value,
            "units": self.units,
            "note": self.note,
        }


class SourceRegistry:
    def __init__(self) -> None:
        self._items: Dict[str, List[SourceRef]] = {}
        self._counter = 0

    def register(self, refs: List[SourceRef] | tuple[SourceRef, ...]) -> str:
        refs = [r for r in refs if r is not None]
        if not refs:
            return ""
        self._counter += 1
        key = f"src-{self._counter:04d}"
        self._items[key] = list(refs)
        return key

    def to_dict(self) -> Dict[str, List[dict]]:
        return {key: [r.to_dict() for r in refs] for key, refs in self._items.items()}


@dataclass
class Bullet:
    text: str
    level: int = 0
    source_id: str = ""


@dataclass
class StateSection:
    heading: str
    maf_value: str
    maf_source_id: str = ""
    bullets: List[Bullet] = field(default_factory=list)


@dataclass
class YearSection:
    year: int
    disclaimer_text: str
    total_use_maf: str
    us_contractors_maf: str
    total_use_source_id: str = ""
    us_contractors_source_id: str = ""
    states: List[StateSection] = field(default_factory=list)
    mexico_heading: str = ""
    mexico_source_id: str = ""
    mexico_bullets: List[Bullet] = field(default_factory=list)


@dataclass
class IcsTableRow:
    state: str
    values: List[str]
    source_ids: List[str]
    is_total: bool = False


@dataclass
class IcsTable:
    year_labels: List[str]
    rows: List[IcsTableRow]
    final_year_total_maf: str
    final_year_label: str
    final_total_source_id: str = ""


@dataclass
class ConservationTableRow:
    state: str
    values: List[str]
    source_ids: List[str]
    total: str
    total_source_id: str = ""
    bold: bool = False


@dataclass
class ConservationSummaryTable:
    rows: List[ConservationTableRow]
    year_labels: List[str]


@dataclass
class ReportContext:
    title_years: str
    scenario_key: str
    scenario_label: str
    mon_year: str
    years: List[YearSection]
    ics_table: IcsTable
    conservation_summary: ConservationSummaryTable
    show_conservation_summary: bool
    conservation_disclaimer: str
    powell_release_subtitle: str
    policy_version: str
    source_map: Dict[str, List[dict]]
    warnings: List[str] = field(default_factory=list)
