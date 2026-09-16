from pathlib import Path

from policy import load_policy

ROOT = Path(__file__).resolve().parents[1]


def test_index_exposes_rules_and_has_no_seis_field():
    html = (ROOT / 'templates' / 'index.html').read_text(encoding='utf-8')
    assert 'id="rulesButton"' in html
    assert 'Rules {{ policy_version }}' in html
    assert 'id="model{{ scenario.key }}"' in html
    assert 'SEIS ROD note' not in html
    assert 'id="seisText"' not in html
    # Max input must remain selectable even while its scenario toggle starts off.
    max_line = next(line for line in html.splitlines() if 'id="model{{ scenario.key }}"' in line)
    assert 'disabled' not in max_line


def test_rules_route_and_active_policy_are_wired():
    app_source = (ROOT / 'app.py').read_text(encoding='utf-8')
    js_source = (ROOT / 'static' / 'app.js').read_text(encoding='utf-8')
    policy = load_policy()

    assert '@app.get("/api/rules")' in app_source
    assert "fetch('/api/rules')" in js_source
    assert policy.version == '2026.09.1'
    assert policy.scenario('Min').display_name == 'Probable Minimum'
    assert policy.report['powell_release_subtitle_scenarios'] == ['Most']
    assert policy.report['rules']['california_system_conservation_rollup']['year_sheet_note_labels'] == ['MWD System Conservation']
