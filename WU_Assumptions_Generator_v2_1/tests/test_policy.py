from policy import load_policy


def test_policy_scenario_excel_blocks_and_visibility():
    policy = load_policy()
    assert policy.version == "2026.09.1"
    assert policy.scenario("Most").value_column == "B"
    assert policy.scenario("Min").value_column == "G"
    assert policy.scenario("Max").value_column == "L"
    assert policy.shows_conservation_summary("Most") is True
    assert policy.shows_conservation_summary("Min") is False
    assert policy.shows_conservation_summary("Max") is False
    assert policy.scenario("Min").display_name == "Probable Minimum"
    assert policy.default_enabled_scenarios == ("Most", "Min")
    assert policy.shows_powell_release_subtitle("Most") is True
    assert policy.shows_powell_release_subtitle("Min") is False
    assert policy.shows_powell_release_subtitle("Max") is False
    assert policy.rolls_california_conservation_into_total("MWD System Conservation") is True
    assert policy.rolls_california_conservation_into_total("Other water left in Mead") is False


def test_effective_year_policy_rules():
    policy = load_policy()
    assert policy.mexico_shortage_label(2026) == "Shortage volume"
    assert policy.mexico_shortage_label(2027) == "Reduced delivery"
    assert policy.az_reduction_override_kaf(2026) is None
    assert policy.az_reduction_override_kaf(2027) == 760
    assert policy.nevada_shortage_label(2026) == "Shortage volume"
    assert policy.nevada_shortage_label(2027) == "Water use reduction"
    assert policy.nevada_system_conservation_rule(2026)["decompose"] is True
    assert policy.nevada_system_conservation_rule(2027)["label"] == "Other water left in Mead is"


def test_all_slots_contains_required_and_optional_without_overlap():
    policy = load_policy()
    assert policy.all_slots[: len(policy.required_slots)] == policy.required_slots
    assert set(policy.required_slots).isdisjoint(policy.optional_slots)
    assert set(policy.all_slots) == set(policy.required_slots) | set(policy.optional_slots)
    assert policy.accepted_volume_units
