"""Build the canonical report content model.

All report-specific policy (scenario Excel columns, wording effective dates,
section visibility, and policy overrides) is read from report_policy.json.
This module applies those rules to RiverWare and Excel source data. The same
ReportContext is consumed by both the HTML preview and Word renderer.
"""
from __future__ import annotations

from decimal import Decimal
from typing import Iterable, List, Optional, Sequence

from excel_source import ConservationRow, ExcelValue, StateUseData, SummaryRow
from policy import ReportPolicy
from report_models import (
    Bullet,
    ConservationSummaryTable,
    ConservationTableRow,
    IcsTable,
    IcsTableRow,
    ReportContext,
    SourceRef,
    SourceRegistry,
    StateSection,
    YearSection,
)
from riverware_annual import AnnualExtractionResult


def _f(value) -> float:
    if value is None:
        return 0.0
    if isinstance(value, Decimal):
        return float(value)
    return float(value)


def _fmt_int(value: float) -> str:
    # RiverWare values converted from m3 can land microscopically below an
    # exact acre-foot boundary (for example 478549.4999996). Snap conversion
    # noise before applying the report's integer rounding.
    return f"{round(round(float(value), 6)):,}"


def _fmt_dec(value: float, digits: int) -> str:
    return f"{value:,.{digits}f}"


def _slot(annual: AnnualExtractionResult, name: str, year: int) -> Optional[float]:
    result = annual.slots.get(name)
    if result is None:
        return None
    value = result.values.get(year)
    return _f(value) if value is not None else None


def _rw_ref(annual: AnnualExtractionResult, slot_name: str, year: int, note: str | None = None) -> SourceRef:
    result = annual.slots.get(slot_name)
    value = result.values.get(year) if result else None
    source_unit = result.source_unit_name if result else None
    detail_parts = []
    if source_unit:
        detail_parts.append(f"RiverWare model unit: {source_unit}; converted to acre-ft by the application.")
    if note:
        detail_parts.append(note)
    return SourceRef(
        kind="RiverWare",
        source_name=annual.model_filename,
        locator=slot_name,
        year=year,
        raw_value=f"{_f(value):,.3f}" if value is not None else None,
        units="acre-ft" if value is not None else None,
        note=" ".join(detail_parts) or None,
    )


def _excel_ref(state_use: StateUseData, entry: ExcelValue, note: str | None = None) -> SourceRef:
    notes = []
    if entry.note:
        notes.append(f"Workbook note: {entry.note}")
    if note:
        notes.append(note)
    return SourceRef(
        kind="Excel",
        source_name=state_use.workbook_filename,
        locator=f"{entry.sheet}!{entry.value_cell}",
        raw_value=f"{entry.value:,.3f}" if entry.value is not None else None,
        units="acre-ft" if entry.value is not None else None,
        note=(f"Label: {entry.label}. " + " ".join(notes)).strip(),
    )


def _conservation_ref(state_use: StateUseData, row: ConservationRow, year: int) -> Optional[SourceRef]:
    cell = row.cells_by_year.get(year)
    value = row.values_by_year.get(year)
    if cell is None:
        return None
    return SourceRef(
        kind="Excel",
        source_name=state_use.workbook_filename,
        locator=f"{row.sheet_name}!{cell}",
        year=year,
        raw_value=f"{value:,.3f}" if value is not None else None,
        units="acre-ft",
        note=f"System Conservation contractor: {row.contractor}.",
    )


def _summary_ref(state_use: StateUseData, row: SummaryRow, year: int | None) -> Optional[SourceRef]:
    if year is None:
        cell = row.total_cell
        value = row.total
        note = f"Conservation summary row: {row.state}; Total column."
    else:
        cell = row.cells_by_year.get(year)
        value = row.values_by_year.get(year)
        note = f"Conservation summary row: {row.state}; year {year}."
    if not cell:
        return None
    return SourceRef(
        kind="Excel",
        source_name=state_use.workbook_filename,
        locator=f"{row.sheet_name}!{cell}",
        year=year,
        raw_value=f"{value:,.3f}" if value is not None else None,
        units="acre-ft" if value is not None else None,
        note=note,
    )


def _calc_ref(locator: str, display_value: str, units: str, note: str) -> SourceRef:
    return SourceRef(
        kind="Calculated",
        source_name="Report calculation",
        locator=locator,
        raw_value=display_value,
        units=units,
        note=note,
    )


def _policy_ref(policy: ReportPolicy, locator: str, value: str, note: str | None = None) -> SourceRef:
    return SourceRef(
        kind="Policy",
        source_name=f"report_policy.json (version {policy.version})",
        locator=locator,
        raw_value=value,
        note=note,
    )


def _register(registry: SourceRegistry, refs: Iterable[Optional[SourceRef]]) -> str:
    return registry.register([r for r in refs if r is not None])


def _needles_bhc_entry(year_data, name: str) -> tuple[float, ExcelValue | None]:
    entry = year_data.find_entry(name) if year_data else None
    return (-_f(entry.value) if entry and entry.value is not None else 0.0, entry)


def _mwd_ics_text(annual: AnnualExtractionResult, year: int) -> tuple[str, List[SourceRef]]:
    creation = _slot(annual, "ICSProjectionData.AnnualCreationEC_MWD_Default", year) or 0.0
    delivery = _slot(annual, "ICSProjectionData.AnnualICSDelivery_MWD_Default", year) or 0.0
    if creation != 0:
        return f"creation of {_fmt_dec(creation / 1000, 1)}", [
            _rw_ref(annual, "ICSProjectionData.AnnualCreationEC_MWD_Default", year)
        ]
    if delivery != 0:
        return f"delivery of {_fmt_dec(delivery / 1000, 1)}", [
            _rw_ref(annual, "ICSProjectionData.AnnualICSDelivery_MWD_Default", year)
        ]
    return "", []


def _az_reductions_text(
    annual: AnnualExtractionResult,
    year: int,
    policy: ReportPolicy,
) -> tuple[str, List[SourceRef]]:
    shortage_slot = "Shortage.CAP Annual Shortage Volume"
    dcp_slot = "ICS Credits.AnnualDCPContribution_AZ"
    delivery_slot = "ICS Credits.AnnualDeliveryEC_AZ"
    x = _slot(annual, shortage_slot, year) or 0.0
    y = _slot(annual, dcp_slot, year) or 0.0
    z = _slot(annual, delivery_slot, year) or 0.0

    phrases: List[str] = []
    refs: List[SourceRef] = []
    if x > 0:
        override = policy.az_reduction_override_kaf(year)
        if override is not None:
            phrases.append(f"water use reduction of {_fmt_int(override)} kaf")
            refs.extend([
                _rw_ref(
                    annual,
                    shortage_slot,
                    year,
                    note="The displayed reduction is overridden by report policy for this year.",
                ),
                _policy_ref(
                    policy,
                    "report.rules.az_water_use_reduction_override",
                    f"{override:g} kaf",
                    "Effective-year display override.",
                ),
            ])
        else:
            phrases.append(f"Shortage volume of {_fmt_int(x / 1000)} kaf")
            refs.append(_rw_ref(annual, shortage_slot, year))
    if y > 0:
        phrases.append(f"DCP contribution of {_fmt_int(y / 1000)} kaf by CAWCD")
        refs.append(_rw_ref(annual, dcp_slot, year))
    if z > 0:
        phrases.append(f"ICS delivery of {_fmt_dec(z / 1000, 1)} kaf")
        refs.append(_rw_ref(annual, delivery_slot, year))

    if not phrases:
        return "", refs
    if len(phrases) == 1:
        return phrases[0], refs
    if len(phrases) == 2:
        return f"{phrases[0]}, and {phrases[1]}", refs
    return f"{phrases[0]}, {phrases[1]}, and {phrases[2]}", refs


def _build_california(
    annual: AnnualExtractionResult,
    state_use: StateUseData,
    year: int,
    policy: ReportPolicy,
    registry: SourceRegistry,
) -> StateSection:
    year_data = state_use.year_sheets.get(year)
    section = StateSection(heading="California", maf_value="")

    mwd_slot = "ForecastUse.MWDDiversionAnnualFC"
    mwd_diversion = _slot(annual, mwd_slot, year) or 0.0
    section.bullets.append(
        Bullet(
            f"MWD annual diversion of {_fmt_int(mwd_diversion / 1000)} kaf",
            source_id=_register(registry, [_rw_ref(annual, mwd_slot, year)]),
        )
    )

    vacated_slot = "ICS Credits.VacatedECICSToBeDelivered_MWD"
    vacated = _slot(annual, vacated_slot, year) or 0.0
    if vacated > 0:
        section.bullets.append(
            Bullet(
                f"Projected diversion includes {_fmt_dec(round(vacated / 1000, 1), 1)} kaf of ICS vacated to system water.",
                level=1,
                source_id=_register(registry, [_rw_ref(annual, vacated_slot, year)]),
            )
        )

    mwd_ics_text, mwd_ics_refs = _mwd_ics_text(annual, year)
    coach_slot = "ICS Credits.AnnualDCPContribution_Coachella"
    mwd_dcp_slot = "ICS Credits.AnnualDCPContribution_MWD"
    coachella_dcp = _slot(annual, coach_slot, year) or 0.0
    mwd_dcp = _slot(annual, mwd_dcp_slot, year) or 0.0
    post_2026 = int(policy.report.get("rules", {}).get("post_2026_start_year", 2027))

    if year < post_2026:
        line2 = ""
        line_refs: List[SourceRef] = list(mwd_ics_refs)
        if mwd_ics_text:
            line2 += f"Projected diversion includes the {mwd_ics_text} kaf of ICS"
        if coachella_dcp > 0:
            coach_kaf = round(coachella_dcp / 1000, 1)
            mwd_kaf = round(mwd_dcp / 1000, 1)
            line2 += f" and {_fmt_dec(coach_kaf, 1)} kaf of water stored for CVWD's share of California's DCP contribution"
            line_refs.extend([_rw_ref(annual, coach_slot, year), _rw_ref(annual, mwd_dcp_slot, year)])
            if line2:
                section.bullets.append(
                    Bullet(line2, level=1, source_id=_register(registry, line_refs))
                )
            section.bullets.append(
                Bullet(
                    f"DCP contribution of {_fmt_dec(mwd_kaf + coach_kaf, 1)} kaf through EC ICS conversion",
                    source_id=_register(
                        registry,
                        [_rw_ref(annual, coach_slot, year), _rw_ref(annual, mwd_dcp_slot, year)],
                    ),
                )
            )
        elif line2:
            section.bullets.append(Bullet(line2, level=1, source_id=_register(registry, line_refs)))
    else:
        ca_ics_slot = "ICSProjectionData.AnnualICSDelivery_MWD_Default"
        ca_ics_delivery = _slot(annual, ca_ics_slot, year) or 0.0
        if ca_ics_delivery > 0:
            section.bullets.append(
                Bullet(
                    f"California ICS delivery of {_fmt_int(round(ca_ics_delivery / 1000, 0))} kaf",
                    source_id=_register(registry, [_rw_ref(annual, ca_ics_slot, year)]),
                )
            )

    ca_cons_refs = [
        _conservation_ref(state_use, row, year)
        for row in state_use.conservation.ca
        if (row.values_by_year.get(year) or 0) != 0
    ]
    total_ca_cons = sum((row.values_by_year.get(year, 0.0) or 0.0) for row in state_use.conservation.ca)

    # Some monthly workbooks identify the California water-left-in-Mead value
    # as MWD System Conservation. Policy controls whether that value is rolled
    # into the displayed California conservation total or printed separately.
    mead_entry = year_data.find_entry("california- water left in mead") if year_data else None
    mead_value = 0.0
    roll_mead_into_total = False
    if mead_entry and mead_entry.value is not None:
        mead_value = -mead_entry.value
        if mead_value > 0 and policy.rolls_california_conservation_into_total(mead_entry.note):
            roll_mead_into_total = True
            total_ca_cons += mead_value
            ca_cons_refs.append(
                _excel_ref(
                    state_use,
                    mead_entry,
                    note=(
                        "Included in Total California System Conservation by report policy; "
                        "the workbook sign is inverted to a positive conservation volume."
                    ),
                )
            )

    if total_ca_cons:
        section.bullets.append(
            Bullet(
                f"Total California System Conservation of {_fmt_dec(total_ca_cons / 1000, 1)} kaf",
                source_id=_register(registry, ca_cons_refs),
            )
        )

    if year >= post_2026:
        ca_overrun_slot = "AnnualWaterUse.CaOverrun"
        ca_overrun = _slot(annual, ca_overrun_slot, year) or 0.0
        reduction = ca_overrun * -1
        if reduction > 0:
            section.bullets.append(
                Bullet(
                    f"California water use reduction of {_fmt_int(round(reduction / 1000, 0))} kaf",
                    source_id=_register(registry, [_rw_ref(annual, ca_overrun_slot, year)]),
                )
            )

    if year_data:
        if mead_entry and mead_value > 0 and not roll_mead_into_total:
            mead_note = mead_entry.note or "System Conservation"
            section.bullets.append(
                Bullet(
                    f"{mead_note} of {_fmt_dec(mead_value / 1000, 1)} kaf",
                    source_id=_register(registry, [_excel_ref(state_use, mead_entry)]),
                )
            )

        needles, needles_entry = _needles_bhc_entry(year_data, "needles")
        if needles_entry is not None:
            section.bullets.append(
                Bullet(
                    f"Needles PSCP volume of {_fmt_int(needles)} af",
                    source_id=_register(registry, [_excel_ref(state_use, needles_entry)]),
                )
            )

    bics_slot = "ICS Credits.AnnualCreationBiNat_MWD"
    bics_mwd = _slot(annual, bics_slot, year) or 0.0
    if bics_mwd > 0:
        section.bullets.append(
            Bullet(
                f"Binational ICS creation of {_fmt_dec(round(bics_mwd / 1000, 1), 1)} kaf by MWD and IID",
                source_id=_register(registry, [_rw_ref(annual, bics_slot, year)]),
            )
        )
    return section


def _build_arizona(
    annual: AnnualExtractionResult,
    state_use: StateUseData,
    year: int,
    policy: ReportPolicy,
    registry: SourceRegistry,
) -> StateSection:
    year_data = state_use.year_sheets.get(year)
    section = StateSection(heading="Arizona", maf_value="")

    cap_slot = "ForecastUse.CAPAnnualFC"
    cap_diversion = _slot(annual, cap_slot, year) or 0.0
    section.bullets.append(
        Bullet(
            f"CAP annual diversion of {_fmt_int(cap_diversion / 1000)} kaf",
            source_id=_register(registry, [_rw_ref(annual, cap_slot, year)]),
        )
    )

    reductions_text, reduction_refs = _az_reductions_text(annual, year, policy)
    if reductions_text:
        phrase = "Projected diversions include a" if year >= int(policy.report.get("rules", {}).get("post_2026_start_year", 2027)) else "Projected diversion includes a"
        section.bullets.append(
            Bullet(
                f"{phrase} {reductions_text}",
                level=1,
                source_id=_register(registry, reduction_refs),
            )
        )

    post_2026 = int(policy.report.get("rules", {}).get("post_2026_start_year", 2027))
    if year < post_2026:
        ics_slot = "ICS Credits.AnnualCreationDCP_CAWCD"
        sys_slot = "ICS Credits.AnnualSysWaterforDCP_CAWCD"
        az_dcp_ics = _slot(annual, ics_slot, year) or 0.0
        az_dcp_sys = _slot(annual, sys_slot, year) or 0.0
        ics_kaf = az_dcp_ics / 1000
        sys_kaf = az_dcp_sys / 1000
        ics_disp = _fmt_dec(round(ics_kaf, 1), 1) if ics_kaf % 1 != 0 else _fmt_int(round(ics_kaf, 0))
        sys_disp = _fmt_dec(round(sys_kaf, 1), 1) if sys_kaf % 1 != 0 else _fmt_int(round(sys_kaf, 0))
        if az_dcp_ics > 0:
            dcp_text = f"{ics_disp} kaf of ICS and {sys_disp} kaf of non-ICS water"
        else:
            dcp_text = f"{sys_disp} kaf of non-ICS water"
        section.bullets.append(
            Bullet(
                f"DCP contribution will be made by creating {dcp_text}",
                source_id=_register(registry, [_rw_ref(annual, ics_slot, year), _rw_ref(annual, sys_slot, year)]),
            )
        )

    conservation_rows = list(state_use.conservation.non_cawcd) + list(state_use.conservation.cawcd)
    total_az_conservation = sum((row.values_by_year.get(year, 0.0) or 0.0) for row in conservation_rows)
    if total_az_conservation:
        section.bullets.append(
            Bullet(
                f"Arizona System Conservation of {_fmt_dec(total_az_conservation / 1000, 1)} kaf",
                source_id=_register(
                    registry,
                    [_conservation_ref(state_use, row, year) for row in conservation_rows if (row.values_by_year.get(year) or 0) != 0],
                ),
            )
        )

    if year_data:
        bhc, bhc_entry = _needles_bhc_entry(year_data, "bhc")
        if bhc > 0 and bhc_entry is not None:
            section.bullets.append(
                Bullet(
                    f"Bullhead City PSCP volume of {_fmt_int(bhc)} af",
                    source_id=_register(registry, [_excel_ref(state_use, bhc_entry)]),
                )
            )

    wellfield_row = next(
        (r for r in state_use.conservation.non_cawcd_all if "242 wellfield" in r.contractor.lower()),
        None,
    )
    if wellfield_row:
        value = wellfield_row.values_by_year.get(year)
        if value and value > 0:
            section.bullets.append(
                Bullet(
                    f"System water created by the 242 Well Field Expansion Project of {_fmt_dec(round(value / 1000, 1), 1)} kaf",
                    source_id=_register(registry, [_conservation_ref(state_use, wellfield_row, year)]),
                )
            )

    vac_slot = "ICS Credits.VacatedDCPICSToSysWater_CAWCD"
    vac = _slot(annual, vac_slot, year) or 0.0
    if vac > 0:
        section.bullets.append(
            Bullet(
                f"Vacated ICS to system water of {_fmt_dec(round(vac / 1000, 1), 1)} kaf",
                source_id=_register(registry, [_rw_ref(annual, vac_slot, year)]),
            )
        )

    bics_slot = "ICS Credits.AnnualCreationBiNat_CAWCD"
    bics = _slot(annual, bics_slot, year) or 0.0
    if bics > 0:
        section.bullets.append(
            Bullet(
                f"Binational ICS creation of {_fmt_dec(round(bics / 1000, 1), 1)} kaf by CAWCD",
                source_id=_register(registry, [_rw_ref(annual, bics_slot, year)]),
            )
        )
    return section


def _build_nevada(
    annual: AnnualExtractionResult,
    year: int,
    policy: ReportPolicy,
    registry: SourceRegistry,
) -> StateSection:
    section = StateSection(heading="Nevada", maf_value="")
    snwp_slot = "ForecastUse.SNWPAnnualFC"
    snwp = _slot(annual, snwp_slot, year) or 0.0
    section.bullets.append(
        Bullet(
            f"SNWA annual use of {_fmt_int(round(snwp / 1000, 0))} kaf. Projected diversion includes:",
            source_id=_register(registry, [_rw_ref(annual, snwp_slot, year)]),
        )
    )

    shortage_slot = "Shortage.SNWP Annual Shortage Volume"
    shortage = _slot(annual, shortage_slot, year) or 0.0
    if shortage > 0:
        section.bullets.append(
            Bullet(
                f"{policy.nevada_shortage_label(year)} of {_fmt_int(shortage / 1000)} kaf",
                level=1,
                source_id=_register(
                    registry,
                    [
                        _rw_ref(annual, shortage_slot, year),
                        _policy_ref(policy, "report.rules.nevada_shortage_wording", policy.nevada_shortage_label(year)),
                    ],
                ),
            )
        )

    ec_creation_slot = "ICS Credits.AnnualCreationEC_NV"
    ec_creation = _slot(annual, ec_creation_slot, year) or 0.0
    if ec_creation > 0:
        section.bullets.append(
            Bullet(
                f"Creation of {_fmt_dec(round(ec_creation / 1000, 1), 1)} kaf of EC ICS",
                level=1,
                source_id=_register(registry, [_rw_ref(annual, ec_creation_slot, year)]),
            )
        )

    ec_del_slot = "ICS Credits.VacatedECICSToBeDelivered_SNWA"
    ec_del = _slot(annual, ec_del_slot, year) or 0.0
    if ec_del > 0:
        section.bullets.append(
            Bullet(
                f"Delivery of {_fmt_dec(round(ec_del / 1000, 1), 1)} kaf of EC ICS due to full ICS bank",
                level=1,
                source_id=_register(registry, [_rw_ref(annual, ec_del_slot, year)]),
            )
        )

    nv_dcp_slot = "ICS Credits.AnnualCreationDCP_NV"
    nv_dcp = _slot(annual, nv_dcp_slot, year) or 0.0
    if nv_dcp > 0:
        section.bullets.append(
            Bullet(
                f"DCP contribution of {_fmt_int(round(nv_dcp / 1000, 0))} kaf through EC ICS conversion",
                source_id=_register(registry, [_rw_ref(annual, nv_dcp_slot, year)]),
            )
        )

    sys_slot = "ICSProjectionData.NV_SystemConservation"
    trib_slot = "ICSProjectionData.NV_LeftinMead"
    tot_syscon = _slot(annual, sys_slot, year) or 0.0
    trib_syscon = _slot(annual, trib_slot, year) or 0.0
    nv_cons_rule = policy.nevada_system_conservation_rule(year)
    if tot_syscon > 0:
        cons_label = str(nv_cons_rule.get("label", "Total System Conservation of"))
        section.bullets.append(
            Bullet(
                f"{cons_label} {_fmt_int(tot_syscon / 1000)} kaf",
                source_id=_register(
                    registry,
                    [
                        _rw_ref(annual, sys_slot, year),
                        _policy_ref(policy, "report.rules.nevada_system_conservation_wording", cons_label),
                    ],
                ),
            )
        )
    if bool(nv_cons_rule.get("decompose", True)) and trib_syscon > 0:
        section.bullets.append(
            Bullet(
                f"Tributary conservation of {_fmt_int(trib_syscon / 1000)} kaf",
                level=1,
                source_id=_register(registry, [_rw_ref(annual, trib_slot, year)]),
            )
        )
    if bool(nv_cons_rule.get("decompose", True)) and tot_syscon > 0 and trib_syscon > 0 and tot_syscon - trib_syscon > 0:
        section.bullets.append(
            Bullet(
                f"Other system conservation of {_fmt_int((tot_syscon - trib_syscon) / 1000)} kaf",
                level=1,
                source_id=_register(
                    registry,
                    [
                        _calc_ref("NV_SystemConservation - NV_LeftinMead", _fmt_int(tot_syscon - trib_syscon), "acre-ft", "Difference of two RiverWare slots."),
                        _rw_ref(annual, sys_slot, year),
                        _rw_ref(annual, trib_slot, year),
                    ],
                ),
            )
        )

    bics_slot = "ICS Credits.AnnualCreationBiNat_NV"
    bics = _slot(annual, bics_slot, year) or 0.0
    if bics > 0:
        section.bullets.append(
            Bullet(
                f"Binational ICS creation of {_fmt_dec(round(bics / 1000, 1), 1)} kaf by SNWA",
                source_id=_register(registry, [_rw_ref(annual, bics_slot, year)]),
            )
        )

    vac_dcp_slot = "ICS Credits.VacatedDCPICSToSysWater_SNWA"
    vac_ec_slot = "ICS Credits.VacatedECICSToSysWater_SNWA"
    vac_dcp = _slot(annual, vac_dcp_slot, year) or 0.0
    vac_ec = _slot(annual, vac_ec_slot, year) or 0.0
    vac_dcp_kaf = round(vac_dcp / 1000, 1) if vac_dcp > 0 else 0.0
    vac_ec_kaf = round(vac_ec / 1000, 1) if vac_ec > 0 else 0.0
    if vac_dcp > 0 or vac_ec > 0:
        total_to_sys = vac_dcp_kaf + vac_ec_kaf if (vac_dcp > 0 and vac_ec > 0) else (vac_dcp_kaf or vac_ec_kaf)
        section.bullets.append(
            Bullet(
                f"Vacated ICS space (to system water) of {_fmt_dec(total_to_sys, 1)} kaf",
                source_id=_register(registry, [_rw_ref(annual, vac_dcp_slot, year), _rw_ref(annual, vac_ec_slot, year)]),
            )
        )
    if vac_dcp > 0:
        section.bullets.append(
            Bullet(
                f"DCP ICS of {_fmt_dec(vac_dcp_kaf, 1)} kaf converted to system water",
                level=1,
                source_id=_register(registry, [_rw_ref(annual, vac_dcp_slot, year)]),
            )
        )
    if vac_ec > 0:
        section.bullets.append(
            Bullet(
                f"EC ICS of {_fmt_dec(vac_ec_kaf, 1)} kaf converted to system water",
                level=1,
                source_id=_register(registry, [_rw_ref(annual, vac_ec_slot, year)]),
            )
        )

    trib_creation_slot = "ICS Credits.AnnualCreationTrib_NV"
    trib_creation = _slot(annual, trib_creation_slot, year) or 0.0
    if trib_creation > 0:
        section.bullets.append(
            Bullet(
                f"Tributary ICS creation of {_fmt_dec(round(trib_creation / 1000, 1), 1)} kaf",
                source_id=_register(registry, [_rw_ref(annual, trib_creation_slot, year)]),
            )
        )
    return section


def _build_mexico(
    annual: AnnualExtractionResult,
    year: int,
    policy: ReportPolicy,
    registry: SourceRegistry,
) -> List[Bullet]:
    bullets: List[Bullet] = []
    shortage_slot = "Mexico Shortage and Surplus.Mexico Annual Shortage"
    create_ws_slot = "Mexico Shortage and Surplus.AnnualCreationMXRecoverableWaterSavings"
    del_ws_slot = "Mexico Shortage and Surplus.AnnualDeliveryMXRecoverableWaterSavings"
    sys_cons_slot = "Mexico Shortage and Surplus.MX_SystemConservation"
    del_mwr_slot = "Mexico Shortage and Surplus.AnnualDeliveryMexicoWaterReserve"
    create_mwr_slot = "Mexico Shortage and Surplus.AnnualCreationMexicoWaterReserve"

    shortage = _slot(annual, shortage_slot, year) or 0.0
    create_ws = _slot(annual, create_ws_slot, year) or 0.0
    del_ws = _slot(annual, del_ws_slot, year) or 0.0
    sys_cons = _slot(annual, sys_cons_slot, year) or 0.0
    del_mwr = _slot(annual, del_mwr_slot, year) or 0.0
    create_mwr = _slot(annual, create_mwr_slot, year) or 0.0

    # Original R logic only emitted the grouped section when recoverable water
    # savings creation was positive. Current post-2026 final reports also need
    # the shortage/reduced-delivery line even when create_ws is zero.
    has_detail = any(v > 0 for v in (shortage, create_ws, del_ws, sys_cons, del_mwr, create_mwr))
    if not has_detail:
        return bullets

    bullets.append(
        Bullet(
            "Projected delivery includes:",
            source_id=_register(
                registry,
                [
                    _rw_ref(annual, shortage_slot, year),
                    _rw_ref(annual, create_ws_slot, year),
                    _rw_ref(annual, del_ws_slot, year),
                    _rw_ref(annual, sys_cons_slot, year),
                    _rw_ref(annual, del_mwr_slot, year),
                    _rw_ref(annual, create_mwr_slot, year),
                ],
            ),
        )
    )
    if shortage > 0:
        shortage_label = policy.mexico_shortage_label(year)
        bullets.append(
            Bullet(
                f"{shortage_label} of {_fmt_int(round(shortage / 1000, 0))} kaf",
                level=1,
                source_id=_register(
                    registry,
                    [
                        _rw_ref(annual, shortage_slot, year),
                        _policy_ref(policy, "report.rules.mexico_shortage_wording", shortage_label),
                    ],
                ),
            )
        )
    if create_ws > 0:
        bullets.append(
            Bullet(
                f"Recoverable Water Savings Contribution of {_fmt_int(round(create_ws / 1000, 0))} kaf",
                level=1,
                source_id=_register(registry, [_rw_ref(annual, create_ws_slot, year)]),
            )
        )
    if del_ws > 0:
        bullets.append(
            Bullet(
                f"Recoverable Water Savings Delivery of {_fmt_int(round(del_ws / 1000, 0))} kaf",
                level=1,
                source_id=_register(registry, [_rw_ref(annual, del_ws_slot, year)]),
            )
        )
    if sys_cons > 0:
        bullets.append(
            Bullet(
                f"Minute 330 System Conservation of {_fmt_dec(round(sys_cons / 1000, 1), 1)} kaf",
                level=1,
                source_id=_register(registry, [_rw_ref(annual, sys_cons_slot, year)]),
            )
        )
    if del_mwr > 0:
        bullets.append(
            Bullet(
                f"Water Reserve delivery of {_fmt_dec(round(del_mwr / 1000, 1), 1)} kaf",
                source_id=_register(registry, [_rw_ref(annual, del_mwr_slot, year)]),
            )
        )
    if create_mwr > 0:
        bullets.append(
            Bullet(
                f"Water Reserve creation of {_fmt_dec(round(create_mwr / 1000, 1), 1)} kaf",
                source_id=_register(registry, [_rw_ref(annual, create_mwr_slot, year)]),
            )
        )
    return bullets


def _build_year_section(
    annual: AnnualExtractionResult,
    state_use: StateUseData,
    year: int,
    year_index: int,
    policy: ReportPolicy,
    registry: SourceRegistry,
) -> YearSection:
    year_data = state_use.year_sheets.get(year)
    ca_slot = "AnnualWaterUse.CaTotalAnnual"
    az_slot = "AnnualWaterUse.AzTotalAnnual"
    nv_slot = "AnnualWaterUse.NvTotalAnnual"
    mx_slot = "Mexico Shortage and Surplus.MexicoAdjustedSched"

    ca_total = _slot(annual, ca_slot, year) or 0.0
    az_total = _slot(annual, az_slot, year) or 0.0
    nv_total = _slot(annual, nv_slot, year) or 0.0
    mexico_sched = _slot(annual, mx_slot, year) or 0.0
    mexico_refs: List[SourceRef] = [_rw_ref(annual, mx_slot, year)]
    if year_index == 0 and year_data:
        excel_mexico = year_data.find_entry("mexico use")
        if excel_mexico and excel_mexico.value is not None:
            mexico_sched = excel_mexico.value
            mexico_refs = [_excel_ref(state_use, excel_mexico, note="Year 1 Mexico value is sourced from the scenario-specific workbook block by policy.")]

    total_use = (ca_total + az_total + nv_total + mexico_sched) / 1_000_000
    us_contractors = (ca_total + az_total + nv_total) / 1_000_000
    total_refs = [
        _calc_ref("CA + AZ + NV + Mexico", _fmt_dec(total_use, 3), "maf", "Sum of displayed state totals and Mexico delivery."),
        _rw_ref(annual, ca_slot, year),
        _rw_ref(annual, az_slot, year),
        _rw_ref(annual, nv_slot, year),
        *mexico_refs,
    ]
    us_refs = [
        _calc_ref("CA + AZ + NV", _fmt_dec(us_contractors, 3), "maf", "Sum of U.S. contractor state totals."),
        _rw_ref(annual, ca_slot, year),
        _rw_ref(annual, az_slot, year),
        _rw_ref(annual, nv_slot, year),
    ]

    section = YearSection(
        year=year,
        disclaimer_text=policy.disclaimer_for_year(year),
        total_use_maf=_fmt_dec(total_use, 3),
        us_contractors_maf=_fmt_dec(us_contractors, 3),
        total_use_source_id=_register(registry, total_refs),
        us_contractors_source_id=_register(registry, us_refs),
    )

    ca_section = _build_california(annual, state_use, year, policy, registry)
    ca_section.maf_value = _fmt_dec(ca_total / 1_000_000, 3)
    ca_section.maf_source_id = _register(registry, [_rw_ref(annual, ca_slot, year)])
    section.states.append(ca_section)

    az_section = _build_arizona(annual, state_use, year, policy, registry)
    az_section.maf_value = _fmt_dec(az_total / 1_000_000, 3)
    az_section.maf_source_id = _register(registry, [_rw_ref(annual, az_slot, year)])
    section.states.append(az_section)

    nv_section = _build_nevada(annual, year, policy, registry)
    nv_section.maf_value = _fmt_dec(nv_total / 1_000_000, 3)
    nv_section.maf_source_id = _register(registry, [_rw_ref(annual, nv_slot, year)])
    section.states.append(nv_section)

    section.mexico_heading = f"Mexico's Scheduled Water Delivery: {_fmt_dec(mexico_sched / 1_000_000, 3)} maf"
    section.mexico_source_id = _register(registry, mexico_refs)
    section.mexico_bullets = _build_mexico(annual, year, policy, registry)
    return section


def _build_ics_table(annual: AnnualExtractionResult, years: Sequence[int], registry: SourceRegistry) -> IcsTable:
    state_slots = {
        "AZ": "ICS Credits.Bank_AZ",
        "CA": "ICS Credits.Bank_CA",
        "NV": "ICS Credits.Bank_NV",
    }
    rows: List[IcsTableRow] = []
    totals = [0.0 for _ in years]
    refs_by_year: List[List[SourceRef]] = [[] for _ in years]
    for state, slot_name in state_slots.items():
        values: List[str] = []
        source_ids: List[str] = []
        for i, year in enumerate(years):
            value = _slot(annual, slot_name, year)
            values.append(_fmt_int(value) if value is not None else "")
            source_ids.append(_register(registry, [_rw_ref(annual, slot_name, year)]))
            if value is not None:
                totals[i] += value
                refs_by_year[i].append(_rw_ref(annual, slot_name, year))
        rows.append(IcsTableRow(state=state, values=values, source_ids=source_ids))

    total_source_ids = []
    for i, year in enumerate(years):
        total_source_ids.append(
            _register(
                registry,
                [
                    _calc_ref("AZ ICS + CA ICS + NV ICS", _fmt_int(totals[i]), "acre-ft", f"ICS total for {year}."),
                    *refs_by_year[i],
                ],
            )
        )
    rows.append(
        IcsTableRow(
            state="Total",
            values=[_fmt_int(v) for v in totals],
            source_ids=total_source_ids,
            is_total=True,
        )
    )
    final_total = totals[-1]
    return IcsTable(
        year_labels=[str(y) for y in years],
        rows=rows,
        final_year_total_maf=_fmt_dec(final_total / 1_000_000, 3),
        final_year_label=str(years[-1]),
        final_total_source_id=total_source_ids[-1],
    )


def _build_conservation_summary(
    state_use: StateUseData,
    years: Sequence[int],
    registry: SourceRegistry,
) -> ConservationSummaryTable:
    rows: List[ConservationTableRow] = []
    for row in state_use.summary_rows:
        display_values: List[str] = []
        source_ids: List[str] = []
        for year in years:
            value = row.values_by_year.get(year)
            display_values.append(_fmt_int(value) if value is not None else "")
            source_ids.append(_register(registry, [_summary_ref(state_use, row, year)]))
        rows.append(
            ConservationTableRow(
                state=row.state,
                values=display_values,
                source_ids=source_ids,
                total=_fmt_int(row.total) if row.total is not None else "",
                total_source_id=_register(registry, [_summary_ref(state_use, row, None)]),
                bold=row.state.strip().lower() in {"annual total", "cumulative total"},
            )
        )
    return ConservationSummaryTable(rows=rows, year_labels=[str(y) for y in years])


def build_report_context(
    annual: AnnualExtractionResult,
    state_use: StateUseData,
    scenario: str,
    mon_year: str,
    policy: ReportPolicy,
) -> ReportContext:
    years = list(annual.target_years)
    if not state_use.has_year3_sheet:
        years = years[:2]

    registry = SourceRegistry()
    year_sections = [
        _build_year_section(annual, state_use, year, idx, policy, registry)
        for idx, year in enumerate(years)
    ]
    ics_table = _build_ics_table(annual, years, registry)
    cons_summary = _build_conservation_summary(state_use, years, registry)
    warnings = list(dict.fromkeys([*annual.warnings, *state_use.warnings]))

    return ReportContext(
        title_years=f"{years[0]} - {years[-1]}",
        scenario_key=scenario,
        scenario_label=policy.scenario(scenario).display_name,
        mon_year=mon_year.strip(),
        years=year_sections,
        ics_table=ics_table,
        conservation_summary=cons_summary,
        show_conservation_summary=policy.shows_conservation_summary(scenario),
        conservation_disclaimer=str(policy.report["notes_conservation_disclaimer"]),
        powell_release_subtitle=(str(policy.report["powell_release_subtitle"]) if policy.shows_powell_release_subtitle(scenario) else ""),
        policy_version=policy.version,
        source_map=registry.to_dict(),
        warnings=warnings,
    )
