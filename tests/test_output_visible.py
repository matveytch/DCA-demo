"""output_<num> is always visible in Finder.

A build folder named with a leading dot is hidden; a Finder window open on the deal
folder during the build would remember it as hidden and keep hiding output_<num> after
the rename. So the build folder has a visible name, and the hidden flag is cleared.
"""

from __future__ import annotations

import os
import shutil
import stat

import pytest
from typer.testing import CliRunner

import dca.cli as cli
from dca.cli import app
from tests.helpers import COMPANY, FILLED, NUM

runner = CliRunner()
CAN_HIDE = hasattr(os, "chflags") and hasattr(stat, "UF_HIDDEN")


@pytest.fixture
def deal(home):
    """A deal generated once: <company>/Outgoing/1-2026-1705_1000kg."""
    assert runner.invoke(app, ["new", NUM]).exit_code == 0
    d = home / "Новые сделки" / f"{NUM}_pending"
    shutil.copy(FILLED, d / f"input_{NUM}.xlsx")
    assert runner.invoke(app, ["generate", str(d)]).exit_code == 0
    return home / COMPANY / "Outgoing" / f"{NUM}_1000kg"


def _service_folders(deal_dir):
    return [p.name for p in deal_dir.iterdir() if p.name.startswith("_сборка_")]


def test_build_folder_has_no_leading_dot(deal, monkeypatch):
    seen = []
    real = cli._publish

    def spy(building, deal_dir, output_dir_name):
        seen.append(building.name)
        return real(building, deal_dir, output_dir_name)
    monkeypatch.setattr(cli, "_publish", spy)
    result = runner.invoke(app, ["generate", str(deal)])
    assert result.exit_code == 0, result.output
    assert len(seen) == 1
    assert seen[0].startswith(f"_сборка_output_{NUM}_") and not seen[0].startswith(".")


@pytest.mark.skipif(not CAN_HIDE, reason="file flags (chflags / UF_HIDDEN) are macOS / BSD only")
def test_hidden_flag_on_the_build_folder_does_not_reach_the_output(deal, monkeypatch):
    real = cli._publish

    def hide_then_publish(building, deal_dir, output_dir_name):
        os.chflags(building, os.stat(building).st_flags | stat.UF_HIDDEN)
        return real(building, deal_dir, output_dir_name)
    monkeypatch.setattr(cli, "_publish", hide_then_publish)
    result = runner.invoke(app, ["generate", str(deal)])
    assert result.exit_code == 0, result.output
    assert not os.stat(deal / f"output_{NUM}").st_flags & stat.UF_HIDDEN


def test_leftovers_of_failed_builds_are_removed(deal):
    for name in (f"_сборка_output_{NUM}_dead0001", f"_сборка_output_{NUM}_dead0002"):
        (deal / name).mkdir()
        (deal / name / f"Agreement_{NUM}.docx").write_bytes(b"half built")
    result = runner.invoke(app, ["generate", str(deal)])
    assert result.exit_code == 0, result.output
    assert _service_folders(deal) == []


def test_no_service_folders_after_generate(deal):
    assert _service_folders(deal) == []
    assert len(list((deal / f"output_{NUM}").glob("*.docx"))) == 4
