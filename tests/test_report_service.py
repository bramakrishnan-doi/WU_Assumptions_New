from report_service import _filename_scenario_issue


def test_filename_scenario_validation():
    assert _filename_scenario_issue("CRMMS_AUG2026_MOST_6MAF.mdl.gz", "Most") == (None, None)
    error, warning = _filename_scenario_issue("CRMMS_AUG2026_MAX.mdl.gz", "Min")
    assert "appears to be a Max" in error
    assert warning is None
    error, warning = _filename_scenario_issue("CRMMS_AUG2026_RUN1.mdl.gz", "Most")
    assert error is None
    assert "could not infer" in warning
