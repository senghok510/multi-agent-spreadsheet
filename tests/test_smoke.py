"""Smoke test so CI has something to run until the feature tests are written."""

import spreadsheet_agents


def test_package_imports():
    assert spreadsheet_agents is not None
