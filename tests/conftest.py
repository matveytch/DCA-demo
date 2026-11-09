"""Shared test setup."""

import pytest


@pytest.fixture(autouse=True)
def excel_is_running(monkeypatch):
    """Tests that put a «~$» owner file next to a workbook mean «Excel has it open».
    Whether Excel really runs on the test machine must not matter: say it does.
    Tests of the leftover-lock case switch this off themselves."""
    monkeypatch.setattr("dca.utils.atomic.excel_running", lambda: True)


@pytest.fixture
def home(tmp_path, monkeypatch):
    """A throwaway data root: `$DCA_HOME` beats the dca.toml lookup, so no test can
    touch real data, wherever its deal folder lives."""
    h = tmp_path / "home"
    h.mkdir()
    monkeypatch.setenv("DCA_HOME", str(h))
    monkeypatch.setenv("DCA_NO_OPEN", "1")
    return h
