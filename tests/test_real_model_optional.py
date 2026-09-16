"""Optional integration test for a locally supplied production/sanitized model.

Run with WU_TEST_MODEL pointing to a .mdl or .mdl.gz. The model is never bundled
with the project or test artifacts.
"""
import os
from pathlib import Path

import pytest

from policy import load_policy
from riverware_annual import load_annual_slots


@pytest.mark.integration
def test_optional_real_model_extracts_all_required_slots():
    path_value = os.getenv("WU_TEST_MODEL")
    if not path_value:
        pytest.skip("Set WU_TEST_MODEL to run the RiverWare integration test.")
    path = Path(path_value)
    policy = load_policy()
    result = load_annual_slots(
        path.read_bytes(), path.name, policy.required_slots, policy.optional_slots, policy.accepted_volume_units
    )
    assert len(result.slots) == len(policy.required_slots)
    assert all(result.slots[name].values for name in policy.required_slots)
