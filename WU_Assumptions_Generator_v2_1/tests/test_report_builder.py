from datetime import datetime
from decimal import Decimal

from excel_source import (
    ConservationRow,
    ConservationTables,
    ExcelValue,
    StateUseData,
    SummaryRow,
    YearSheetData,
)
from policy import load_policy
from report_builder import build_report_context
from riverware_annual import AnnualExtractionResult, SlotResult


def make_annual():
    policy = load_policy()
    years = (2026, 2027, 2028)
    slots = {
        name: SlotResult(
            slot_name=name,
            unit_name="acre-ft",
            values={y: Decimal("0") for y in years},
            source_unit_name="m3",
        )
        for name in policy.required_slots
    }

    def set_values(name, values):
        slots[name].values = {y: Decimal(str(v)) for y, v in zip(years, values)}

    set_values("AnnualWaterUse.CaTotalAnnual", [3_787_000, 4_210_000, 4_220_000])
    set_values("AnnualWaterUse.AzTotalAnnual", [2_033_000, 2_049_000, 2_049_000])
    # Deliberately different: report totals must NOT use the apportionment value.
    set_values("AnnualWaterUse.Arizona_Apportionment", [2_033_000, 2_764_000, 2_764_000])
    set_values("AnnualWaterUse.NvTotalAnnual", [208_000, 198_000, 198_000])
    set_values("Mexico Shortage and Surplus.MexicoAdjustedSched", [1_304_000, 1_250_000, 1_250_000])
    set_values("Mexico Shortage and Surplus.Mexico Annual Shortage", [50_000, 250_000, 250_000])
    set_values("Shortage.CAP Annual Shortage Volume", [320_000, 755_150, 755_150])
    set_values("Shortage.SNWP Annual Shortage Volume", [13_000, 50_000, 50_000])
    set_values("ICSProjectionData.NV_SystemConservation", [114_000, 52_000, 52_000])
    set_values("ICSProjectionData.NV_LeftinMead", [35_000, 0, 0])
    set_values("ICS Credits.AnnualCreationTrib_NV", [0, 35_000, 35_000])
    return AnnualExtractionResult(
        model_filename="CRMMS_TEST_MOST.mdl.gz",
        start_year=2026,
        target_years=years,
        run_start=datetime(2026, 1, 1),
        run_finish=datetime(2028, 12, 31),
        slots=slots,
        warnings=[],
    )


def make_state_use(scenario="Most"):
    years = (2026, 2027, 2028)
    sheets = {}
    for year in years:
        sheets[year] = YearSheetData(
            sheet_name=str(year),
            scenario=scenario,
            entries=[
                ExcelValue(1_304_000 if year == 2026 else 1_250_000, str(year), "Mexico Use", "B2"),
                ExcelValue(-145, str(year), "PSCP (Needles)", "B3"),
                ExcelValue(-2477, str(year), "PSCP (BHC)", "B4"),
            ],
        )
    summary = [
        SummaryRow(s, {y: 0 for y in years}, {y: f"B{idx}" for y in years}, 0, f"E{idx}", "SysCon SummaryTable")
        for idx, s in enumerate(["AZ", "CA", "NV", "Annual Total", "Cumulative Total"], 2)
    ]
    return StateUseData(
        workbook_filename="Projected State Use-TEST.xlsx",
        scenario=scenario,
        year_sheets=sheets,
        conservation=ConservationTables([], [], [], [], list(years)),
        summary_rows=summary,
        has_year3_sheet=True,
        warnings=[],
    )


def test_report_uses_arizona_total_annual_not_apportionment():
    policy = load_policy()
    ctx = build_report_context(make_annual(), make_state_use(), "Most", "August 2026", policy)
    y2027 = ctx.years[1]
    assert y2027.total_use_maf == "7.707"
    assert y2027.us_contractors_maf == "6.457"
    arizona = next(s for s in y2027.states if s.heading == "Arizona")
    assert arizona.maf_value == "2.049"
    refs = ctx.source_map[arizona.maf_source_id]
    assert refs[0]["locator"] == "AnnualWaterUse.AzTotalAnnual"


def test_preview_content_rules_are_in_canonical_model():
    policy = load_policy()
    ctx = build_report_context(make_annual(), make_state_use(), "Most", "August 2026", policy)
    assert ctx.show_conservation_summary is True
    assert any("Shortage volume of 50 kaf" in b.text for b in ctx.years[0].mexico_bullets)
    assert any("Reduced delivery of 250 kaf" in b.text for b in ctx.years[1].mexico_bullets)
    assert any("water use reduction of 760 kaf" in b.text for b in next(s for s in ctx.years[1].states if s.heading == "Arizona").bullets)

    nv_2026 = next(s for s in ctx.years[0].states if s.heading == "Nevada")
    nv_2027 = next(s for s in ctx.years[1].states if s.heading == "Nevada")
    assert any("Shortage volume of 13 kaf" in b.text for b in nv_2026.bullets)
    assert any("Total System Conservation of 114 kaf" in b.text for b in nv_2026.bullets)
    assert any("Tributary conservation of 35 kaf" in b.text for b in nv_2026.bullets)
    assert any("Other system conservation of 79 kaf" in b.text for b in nv_2026.bullets)
    assert any("Water use reduction of 50 kaf" in b.text for b in nv_2027.bullets)
    assert any("Other water left in Mead is 52 kaf" in b.text for b in nv_2027.bullets)
    assert not any("Other system conservation" in b.text for b in nv_2027.bullets)

    min_ctx = build_report_context(make_annual(), make_state_use("Min"), "Min", "August 2026", policy)
    assert min_ctx.show_conservation_summary is False


def test_scenario_labels_and_powell_subtitle_visibility():
    policy = load_policy()
    most_ctx = build_report_context(make_annual(), make_state_use("Most"), "Most", "August 2026", policy)
    min_ctx = build_report_context(make_annual(), make_state_use("Min"), "Min", "August 2026", policy)
    max_ctx = build_report_context(make_annual(), make_state_use("Max"), "Max", "August 2026", policy)

    assert most_ctx.scenario_label == "Most Probable"
    assert most_ctx.powell_release_subtitle == "For the 6 maf & 7 maf Powell Release Scenarios"
    assert min_ctx.scenario_label == "Probable Minimum"
    assert min_ctx.powell_release_subtitle == ""
    assert max_ctx.scenario_label == "Probable Max"
    assert max_ctx.powell_release_subtitle == ""


def test_mwd_system_conservation_rolls_into_california_total():
    policy = load_policy()
    state_use = make_state_use("Most")
    state_use.conservation.ca = [
        ConservationRow(
            number=1,
            contractor="California contractor",
            values_by_year={2026: 448_500},
            cells_by_year={2026: "AJ5"},
            sheet_name="System Conservation",
        )
    ]
    state_use.year_sheets[2026].entries.append(
        ExcelValue(
            -185_900,
            "2026",
            "California- water left in Mead",
            "B5",
            note="MWD System Conservation",
            note_cell="D5",
        )
    )

    ctx = build_report_context(make_annual(), state_use, "Most", "August 2026", policy)
    california = next(s for s in ctx.years[0].states if s.heading == "California")
    texts = [b.text for b in california.bullets]
    total = next(b for b in california.bullets if b.text.startswith("Total California System Conservation"))

    assert total.text == "Total California System Conservation of 634.4 kaf"
    assert not any(t.startswith("MWD System Conservation") for t in texts)
    refs = ctx.source_map[total.source_id]
    assert {r["locator"] for r in refs} == {"System Conservation!AJ5", "2026!B5"}


def test_non_mwd_water_left_in_mead_stays_separate():
    policy = load_policy()
    state_use = make_state_use("Most")
    state_use.conservation.ca = [
        ConservationRow(
            number=1,
            contractor="California contractor",
            values_by_year={2026: 448_500},
            cells_by_year={2026: "AJ5"},
            sheet_name="System Conservation",
        )
    ]
    state_use.year_sheets[2026].entries.append(
        ExcelValue(
            -163_700,
            "2026",
            "California- water left in Mead",
            "B5",
            note="Other water left in Mead",
            note_cell="D5",
        )
    )

    ctx = build_report_context(make_annual(), state_use, "Most", "August 2026", policy)
    california = next(s for s in ctx.years[0].states if s.heading == "California")
    texts = [b.text for b in california.bullets]
    assert "Total California System Conservation of 448.5 kaf" in texts
    assert "Other water left in Mead of 163.7 kaf" in texts
