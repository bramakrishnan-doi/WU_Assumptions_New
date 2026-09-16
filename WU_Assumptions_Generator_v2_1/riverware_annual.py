"""RiverWare .mdl/.mdl.gz annual-slot extractor.

This module reuses the low-level RiverWare text parsing approach from the
original hourly Davis/Parker schedule adapter (object/slot declarations,
AggSeriesSlot member resolution, DSeries/CSeries decompression, RiverWare
timestamp parsing) but is aimed at a different report: pulling a fixed list
of *annual* slots for the model's start year and the two following years,
converting volumes from cubic meters to acre-feet.

RiverWare stores annual (and other end-of-timestep) values on a "24:00"
convention -- e.g. a value for calendar year 2026 is stamped
12-31-2026 24:00:00, which naive date math rolls forward into
2027-01-01 00:00:00. This module subtracts one minute before reading the
year off a timestamp so annual values map back to the year that closed,
matching the convention confirmed against real model data.
"""

from __future__ import annotations

import gzip
import io
import re
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from decimal import Decimal, InvalidOperation
from typing import Dict, List, Mapping, Sequence, Tuple

MAX_DECOMPRESSED_BYTES = 512 * 1024 * 1024

# 1 acre-ft = 43,560 ft^3; 1 ft = 0.3048 m exactly.
ACRE_FT_PER_M3 = Decimal(1) / (Decimal(43560) * (Decimal("0.3048") ** 3))

# The report-specific slot list is intentionally not hard-coded here.
# It is supplied by the versioned policy layer so this parser stays generic.


class RiverwareValidationError(Exception):
    def __init__(self, errors: List[str]):
        super().__init__("; ".join(errors))
        self.errors = errors


@dataclass(frozen=True)
class UnitInfo:
    unit_type: str
    scale: Decimal
    name: str


@dataclass(frozen=True)
class RunInfo:
    start: datetime
    finish: datetime
    step_size: int
    step_unit: str


@dataclass(frozen=True)
class SlotDeclaration:
    slot_type: str
    name: str
    local_index: int


@dataclass(frozen=True)
class ModelObject:
    name: str
    slots: Tuple[SlotDeclaration, ...]
    block: str


@dataclass(frozen=True)
class SeriesPoint:
    timestamp: datetime
    value: Decimal | None


@dataclass(frozen=True)
class SeriesData:
    full_slot_name: str
    rows: Tuple[SeriesPoint, ...]
    unit_name: str | None
    step_size: int | None
    step_unit: str | None
    scalar_value: Decimal | None = None


@dataclass(frozen=True)
class _SlotDetail:
    unit_code: int | None
    unit_name: str | None
    series_start: str | None
    series_end: str | None
    step_size: int | None
    step_unit: str | None
    raw_series: str | None
    scalar_value: str | None
    is_table: bool


@dataclass
class SlotResult:
    slot_name: str
    unit_name: str | None
    values: Dict[int, Decimal | None] = field(default_factory=dict)  # year -> acre-ft value
    source_unit_name: str | None = None


@dataclass
class AnnualExtractionResult:
    model_filename: str
    start_year: int
    target_years: Tuple[int, int, int]
    run_start: datetime
    run_finish: datetime
    slots: Dict[str, SlotResult]
    warnings: List[str] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Byte decoding
# ---------------------------------------------------------------------------

def _decode_model_bytes(data: bytes, filename: str) -> str:
    lower_name = filename.lower()
    is_gzip = lower_name.endswith(".gz") or data[:2] == b"\x1f\x8b"
    payload = data
    if is_gzip:
        try:
            with gzip.GzipFile(fileobj=io.BytesIO(data), mode="rb") as stream:
                payload = stream.read(MAX_DECOMPRESSED_BYTES + 1)
        except (OSError, EOFError) as exc:
            raise RiverwareValidationError(
                ["The uploaded .mdl.gz file is not a valid gzip archive."]
            ) from exc
        if len(payload) > MAX_DECOMPRESSED_BYTES:
            raise RiverwareValidationError(
                ["The decompressed RiverWare model exceeds the 512 MB safety limit."]
            )

    try:
        return payload.decode("utf-8-sig")
    except UnicodeDecodeError:
        try:
            return payload.decode("cp1252")
        except UnicodeDecodeError as exc:
            raise RiverwareValidationError(
                ["The RiverWare model text could not be decoded as UTF-8 or Windows-1252."]
            ) from exc


def _parse_decimal_token(token: str) -> Decimal | None:
    cleaned = token.strip()
    if not cleaned or cleaned == "\\" or cleaned.lower() in {
        "nan", "+nan", "-nan", "1.#qnan", "-1.#qnan", "1.#ind", "-1.#ind",
    }:
        return None
    try:
        value = Decimal(cleaned)
    except InvalidOperation:
        return None
    return value if value.is_finite() else None


def _parse_unit_manager(text: str) -> Dict[int, UnitInfo]:
    table: Dict[int, UnitInfo] = {}
    pattern = re.compile(
        r"\$unitMgr\s+scaledUnit\s+(\d+)\s+\{([^}]*)\}\s+([+\-\d.eE]+)\s+\{([^}]*)\}"
    )
    for match in pattern.finditer(text):
        try:
            scale = Decimal(match.group(3))
        except InvalidOperation:
            scale = Decimal(1)
        table[int(match.group(1))] = UnitInfo(
            unit_type=match.group(2).strip(),
            scale=scale,
            name=match.group(4).strip(),
        )
    return table


def _parse_rw_date(value: str) -> datetime | None:
    """Parse RiverWare MM-DD-YYYY HH:MM:SS, including the 24:00 convention."""

    match = re.fullmatch(
        r"\s*(\d{2})-(\d{2})-(\d{4})\s+(\d{2}):(\d{2}):(\d{2})\s*",
        value,
    )
    if not match:
        return None
    month, day, year, hour, minute, second = (int(part) for part in match.groups())
    if hour > 24 or minute > 59 or second > 59:
        return None
    if hour == 24 and (minute != 0 or second != 0):
        return None
    try:
        base = datetime(year, month, day)
    except ValueError:
        return None
    return base + timedelta(hours=hour, minutes=minute, seconds=second)


def _parse_run_info(text: str) -> RunInfo | None:
    match = re.search(
        r"\$ws\.runInfo\s+runParam\s+\{([^}]*)\}\s+\{([^}]*)\}\s+(\d+)\s+(\w+)",
        text,
    )
    if not match:
        return None
    start = _parse_rw_date(match.group(1))
    finish = _parse_rw_date(match.group(2))
    if start is None or finish is None:
        return None
    return RunInfo(
        start=start,
        finish=finish,
        step_size=int(match.group(3)),
        step_unit=match.group(4).upper(),
    )


def _parse_model_objects(text: str) -> Tuple[ModelObject, ...]:
    objects: List[ModelObject] = []
    object_matches = list(
        re.finditer(r"^set\s+obj\s+\{([^}]*)\}\s*\r?$", text, re.MULTILINE)
    )
    slot_pattern = re.compile(r'"\$o"\s+\{(\w+Slot)\}\s+\{([^}]*)\}')

    for index, match in enumerate(object_matches):
        start = match.start()
        end = object_matches[index + 1].start() if index + 1 < len(object_matches) else len(text)
        block = text[start:end]
        slots = tuple(
            SlotDeclaration(
                slot_type=slot_match.group(1),
                name=slot_match.group(2),
                local_index=slot_match.start(),
            )
            for slot_match in slot_pattern.finditer(block)
        )
        if slots:
            objects.append(ModelObject(name=match.group(1), slots=slots, block=block))
    return tuple(objects)


def _decompress_series(raw_series: str) -> List[str]:
    tokens = [token for token in raw_series.strip().split() if token != "\\"]
    values: List[str] = []
    index = 0
    while index < len(tokens):
        token = tokens[index]
        if index + 2 < len(tokens) and tokens[index + 1] == "@":
            try:
                count = int(tokens[index + 2])
            except ValueError:
                values.append(token)
                index += 1
                continue
            if count < 0:
                values.append(token)
                index += 3
                continue
            values.extend([token] * count)
            index += 3
        else:
            values.append(token)
            index += 1
    return values


def _add_months_like_javascript(value: datetime, months: int) -> datetime:
    month_index = value.year * 12 + (value.month - 1) + months
    target_year, zero_month = divmod(month_index, 12)
    first = value.replace(year=target_year, month=zero_month + 1, day=1)
    return first + timedelta(days=value.day - 1)


def _add_timestep(value: datetime, size: int, unit: str) -> datetime:
    normalized = unit.upper()
    if normalized == "HOUR":
        return value + timedelta(hours=size)
    if normalized == "DAY":
        return value + timedelta(days=size)
    if normalized == "MINUTE":
        return value + timedelta(minutes=size)
    if normalized == "MONTH":
        return _add_months_like_javascript(value, size)
    if normalized == "YEAR":
        return _add_months_like_javascript(value, size * 12)
    raise ValueError(f"unsupported RiverWare timestep unit {unit!r}")


def _slot_subsection(model_object: ModelObject, slot: SlotDeclaration) -> str:
    after_declaration = model_object.block[slot.local_index:]
    next_slot = re.search(
        r'\n\s*"\$o"\s+\{\w+Slot\}\s+\{[^}]*\}',
        after_declaration[1:],
    )
    if not next_slot:
        return after_declaration
    return after_declaration[: 1 + next_slot.start()]


def _parse_slot_subsection(subsection: str, slot_type: str) -> _SlotDetail:
    is_table = slot_type in {"TableSlot", "PeriodicSlot", "TableSeriesSlot"}
    unit_code_match = re.search(r'"\$s"\s+unit\s+(\d+)\b', subsection)
    unit_code = int(unit_code_match.group(1)) if unit_code_match else None

    if slot_type == "ScalarSlot":
        scalar_match = re.search(r'"\$s"\s+value\s+([^\s{]+)\s+\{([^}]*)\}', subsection)
        return _SlotDetail(
            unit_code=unit_code,
            unit_name=scalar_match.group(2).strip() if scalar_match else None,
            series_start=None, series_end=None, step_size=None, step_unit=None,
            raw_series=None,
            scalar_value=scalar_match.group(1) if scalar_match else None,
            is_table=False,
        )

    if is_table:
        return _SlotDetail(
            unit_code=unit_code, unit_name=None, series_start=None, series_end=None,
            step_size=None, step_unit=None, raw_series=None, scalar_value=None, is_table=True,
        )

    series_terminator = r"(?=\r?\n\s*\"\$s\"|\Z)"
    d_series_match = re.search(
        r'"\$s"\s+setDSeries\s+\{([^}]*)\}\s+\{([^}]*)\}\s+\{([^}]*)\}\s+'
        r"(\d+)\s+(\w+)\s+-?\d+\s+([\s\S]*?)" + series_terminator,
        subsection,
    )
    if d_series_match:
        return _SlotDetail(
            unit_code=unit_code,
            unit_name=d_series_match.group(1).strip() or None,
            series_start=d_series_match.group(2),
            series_end=d_series_match.group(3),
            step_size=int(d_series_match.group(4)),
            step_unit=d_series_match.group(5).upper(),
            raw_series=d_series_match.group(6),
            scalar_value=None,
            is_table=False,
        )

    c_series_match = re.search(
        r'"\$s"\s+setCSeries\s+\{([^}]*)\}\s+(\d+)\s+(\d+)\s+(\w+)\s+-?\d+\s+'
        r"([\s\S]*?)" + series_terminator,
        subsection,
    )
    if c_series_match:
        return _SlotDetail(
            unit_code=unit_code,
            unit_name=None,
            series_start=c_series_match.group(1),
            series_end=None,
            step_size=int(c_series_match.group(2)),
            step_unit=c_series_match.group(4).upper(),
            raw_series=c_series_match.group(5),
            scalar_value=None,
            is_table=False,
        )

    return _SlotDetail(
        unit_code=unit_code, unit_name=None, series_start=None, series_end=None,
        step_size=None, step_unit=None, raw_series=None, scalar_value=None, is_table=False,
    )


def _extract_slot_detail(model_object: ModelObject, slot: SlotDeclaration) -> _SlotDetail:
    subsection = _slot_subsection(model_object, slot)
    return _parse_slot_subsection(subsection, slot.slot_type)


def _extract_aggregate_member_detail(
    model_object: ModelObject,
    aggregate_slot: SlotDeclaration,
    member_name: str,
) -> _SlotDetail | None:
    if aggregate_slot.slot_type != "AggSeriesSlot":
        return None

    aggregate_block = _slot_subsection(model_object, aggregate_slot)
    root_label_match = re.search(r'"\$s"\s+setAggLabel\s+\{([^}]*)\}', aggregate_block)
    root_label = root_label_match.group(1).strip() if root_label_match else None

    child_selector = re.compile(
        r'^\s*set\s+s\s+"\$o\.' + re.escape(aggregate_slot.name) + r'\.([^"]+)"\s*$',
        re.MULTILINE,
    )
    child_matches = list(child_selector.finditer(aggregate_block))

    if root_label == member_name:
        end = child_matches[0].start() if child_matches else len(aggregate_block)
        member_block = aggregate_block[:end]
        return _parse_slot_subsection(member_block, "SeriesSlot")

    for index, match in enumerate(child_matches):
        if match.group(1).strip() != member_name:
            continue
        end = (
            child_matches[index + 1].start()
            if index + 1 < len(child_matches)
            else len(aggregate_block)
        )
        member_block = aggregate_block[match.start():end]
        return _parse_slot_subsection(member_block, "SeriesSlot")

    return None


def _extract_series(
    objects: Sequence[ModelObject],
    unit_table: Mapping[int, UnitInfo],
    full_slot_name: str,
) -> SeriesData | None:
    if "." not in full_slot_name:
        return None
    object_name, slot_name = full_slot_name.split(".", 1)
    model_object = next((item for item in objects if item.name == object_name), None)
    if model_object is None:
        return None

    slot = next((item for item in model_object.slots if item.name == slot_name), None)
    if slot is not None:
        detail = _extract_slot_detail(model_object, slot)
        effective_slot_type = slot.slot_type
    else:
        if "." not in slot_name:
            return None
        aggregate_name, member_name = slot_name.rsplit(".", 1)
        aggregate_slot = next(
            (
                item for item in model_object.slots
                if item.name == aggregate_name and item.slot_type == "AggSeriesSlot"
            ),
            None,
        )
        if aggregate_slot is None:
            return None
        detail = _extract_aggregate_member_detail(model_object, aggregate_slot, member_name)
        if detail is None:
            return None
        effective_slot_type = "SeriesSlot"

    unit_info = unit_table.get(detail.unit_code) if detail.unit_code is not None else None
    unit_name = detail.unit_name or (unit_info.name if unit_info else None)

    if detail.is_table:
        return SeriesData(full_slot_name, (), unit_name, None, None)

    if effective_slot_type == "ScalarSlot":
        scalar = _parse_decimal_token(detail.scalar_value or "")
        return SeriesData(full_slot_name, (), unit_name, None, None, scalar_value=scalar)

    if not detail.raw_series or not detail.series_start or detail.step_size is None or not detail.step_unit:
        return SeriesData(full_slot_name, (), unit_name, detail.step_size, detail.step_unit)

    start = _parse_rw_date(detail.series_start)
    if start is None:
        return SeriesData(full_slot_name, (), unit_name, detail.step_size, detail.step_unit)

    rows: List[SeriesPoint] = []
    current = start
    for token in _decompress_series(detail.raw_series):
        rows.append(SeriesPoint(timestamp=current, value=_parse_decimal_token(token)))
        try:
            current = _add_timestep(current, detail.step_size, detail.step_unit)
        except ValueError:
            break

    return SeriesData(full_slot_name, tuple(rows), unit_name, detail.step_size, detail.step_unit)


# ---------------------------------------------------------------------------
# Unit conversion
# ---------------------------------------------------------------------------

def _normalize_unit_name(value: str | None) -> str:
    return (value or "").strip().lower().replace("\u00b3", "3").replace("^", "").replace(" ", "").replace("_", "").replace("-", "")


def volume_to_acre_ft(value: Decimal, unit_name: str | None) -> Decimal:
    unit = _normalize_unit_name(unit_name)
    if unit in {"", "none"}:
        raise ValueError("has missing/unitless volume units; expected an explicit acre-ft, m3, or ft3 unit")
    if unit in {"acreft", "acft", "acrefeet"}:
        return value
    if unit in {"m3", "cm", "cubicmeter", "cubicmeters", "cum"}:
        return value * ACRE_FT_PER_M3
    if unit in {"ft3", "cf", "cubicfoot", "cubicfeet"}:
        return value / Decimal(43560)
    raise ValueError(f"uses unsupported volume units {unit_name or '(missing)'}; expected acre-ft, m3, or ft3")


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def load_annual_slots(
    data: bytes,
    filename: str,
    slot_names: Sequence[str],
    optional_slot_names: Sequence[str] = (),
    accepted_volume_units: Sequence[str] = (),
) -> AnnualExtractionResult:
    """Parse a RiverWare .mdl/.mdl.gz upload and extract the report's annual slots.

    Values are returned in acre-feet for the model's start year and the
    following two calendar years.
    """

    if not filename.lower().endswith((".mdl", ".mdl.gz")):
        raise RiverwareValidationError(
            ["The RiverWare source must have a .mdl or .mdl.gz extension."]
        )

    text = _decode_model_bytes(data, filename)
    run_info = _parse_run_info(text)
    if run_info is None:
        raise RiverwareValidationError(
            ["The model Start Timestep could not be read from $ws.runInfo runParam."]
        )

    unit_table = _parse_unit_manager(text)
    objects = _parse_model_objects(text)
    if not objects:
        raise RiverwareValidationError(
            ["No RiverWare objects and slots could be parsed from the uploaded model."]
        )

    start_year = run_info.start.year
    target_years = (start_year, start_year + 1, start_year + 2)

    errors: List[str] = []
    warnings: List[str] = []
    slots: Dict[str, SlotResult] = {}

    def _report_issue(slot_name: str, message: str) -> None:
        if slot_name in optional_slots:
            warnings.append(f"(optional slot skipped) {message}")
        else:
            errors.append(message)

    optional_slots = frozenset(optional_slot_names)
    accepted_units = {_normalize_unit_name(v) for v in accepted_volume_units if str(v).strip()}

    for slot_name in slot_names:
        series = _extract_series(objects, unit_table, slot_name)

        if series is None:
            slots[slot_name] = SlotResult(slot_name, None, {y: None for y in target_years}, None)
            _report_issue(slot_name, f"Slot not found: {slot_name}")
            continue

        if accepted_units and _normalize_unit_name(series.unit_name) not in accepted_units:
            slots[slot_name] = SlotResult(slot_name, series.unit_name, {y: None for y in target_years}, series.unit_name)
            _report_issue(
                slot_name,
                f"{slot_name}: volume unit {series.unit_name or '(missing)'} is not accepted by the active report policy.",
            )
            continue

        if series.scalar_value is not None and not series.rows:
            try:
                converted = volume_to_acre_ft(series.scalar_value, series.unit_name)
            except ValueError as exc:
                _report_issue(slot_name, f"{slot_name}: {exc}")
                slots[slot_name] = SlotResult(slot_name, series.unit_name, {y: None for y in target_years}, series.unit_name)
                continue
            slots[slot_name] = SlotResult(
                slot_name, "acre-ft", {y: converted for y in target_years}, series.unit_name
            )
            continue

        if series.step_unit != "YEAR" or series.step_size != 1:
            _report_issue(
                slot_name,
                f"{slot_name}: expected a 1-year series, found "
                f"{series.step_size or '?'} {series.step_unit or 'unknown'} timestep.",
            )
            slots[slot_name] = SlotResult(slot_name, series.unit_name, {y: None for y in target_years}, series.unit_name)
            continue

        # RiverWare stores annual values on a 24:00 end-of-timestep
        # convention (e.g. 12-31-2026 24:00:00), which rolls into
        # 2027-01-01 00:00:00 under naive date math. Subtract one minute
        # so the point maps back to the year that just closed.
        by_year: Dict[int, Decimal | None] = {}
        for row in series.rows:
            effective = row.timestamp - timedelta(minutes=1)
            by_year[effective.year] = row.value

        values: Dict[int, Decimal | None] = {}
        for year in target_years:
            raw = by_year.get(year)
            if raw is None:
                values[year] = None
                if year in by_year:
                    _report_issue(slot_name, f"{slot_name}: value for {year} could not be parsed as a number.")
                else:
                    _report_issue(slot_name, f"{slot_name}: no data point stamped for year {year}.")
                continue
            try:
                values[year] = volume_to_acre_ft(raw, series.unit_name)
            except ValueError as exc:
                _report_issue(slot_name, f"{slot_name}: {exc}")
                values[year] = None

        slots[slot_name] = SlotResult(slot_name, "acre-ft", values, series.unit_name)

    if errors:
        raise RiverwareValidationError(errors[:50])

    return AnnualExtractionResult(
        model_filename=filename,
        start_year=start_year,
        target_years=target_years,
        run_start=run_info.start,
        run_finish=run_info.finish,
        slots=slots,
        warnings=warnings[:50],
    )
