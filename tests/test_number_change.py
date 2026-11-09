"""Agreement number changed in an existing deal; the result of generate is checked.

The operator changes the agreement number and quantity and runs generate again: the
output must be complete in the new folder, under the path the program prints.
"""

from __future__ import annotations

import shutil
import sys
from datetime import date

import pytest
from typer.testing import CliRunner

import dca.cli as cli
from dca.cli import app
from tests.helpers import COMPANY, FILLED, set_annexes, set_deal_values

runner = CliRunner()
OLD, NEW = "1-2026-1705", "2-2026-0101"
NEW_VALUES = {"agreement_num": NEW, "quantity_kg": 1200, "batch_use_kg_1": 600, "batch_use_kg_2": 600}
N_DOCX = 4    # agreement, annex 1, annex 2, cover list


def _filed_deal(home):
    """A deal generated once and filed as <company>/Outgoing/1-2026-1705_1000kg."""
    assert runner.invoke(app, ["new", OLD]).exit_code == 0
    d = home / "Новые сделки" / f"{OLD}_pending"
    shutil.copy(FILLED, d / f"input_{OLD}.xlsx")
    result = runner.invoke(app, ["generate", str(d)])
    assert result.exit_code == 0, result.output
    return home / COMPANY / "Outgoing" / f"{OLD}_1000kg"


def _change(deal_dir, values):
    xlsx = deal_dir / f"input_{OLD}.xlsx"
    set_deal_values(xlsx, values)
    set_annexes(xlsx, [("1", date(2026, 5, 17), 700), ("2", date(2026, 6, 1), 500)])


def test_number_and_quantity_change_gives_a_complete_output_in_the_new_folder(home):
    old_dir = _filed_deal(home)
    _change(old_dir, NEW_VALUES)
    result = runner.invoke(app, ["generate", str(old_dir)])
    assert result.exit_code == 0, result.output
    new_dir = home / COMPANY / "Outgoing" / f"{NEW}_1200kg"
    assert not old_dir.exists() and new_dir.is_dir()
    out = new_dir / f"output_{NEW}"
    assert f"Done. Open: {out}" in result.output
    assert len(list(out.glob("*.docx"))) == N_DOCX
    names = {p.name for p in new_dir.iterdir()}
    # The old documents do not lie next to the new ones: they are in the archive.
    assert f"output_{OLD}" not in names
    archived = [p for p in (new_dir / "_archive").iterdir() if (p / f"Package_{OLD}.zip").exists()]
    assert len(archived) == 1
    # Runners carry the new number; the empty signed_<old> is gone.
    assert {f"run_{NEW}.command", f"transfer_{NEW}.command", f"amend_{NEW}.command"} <= names
    assert not any(n.endswith(".command") and OLD in n for n in names)
    assert f"signed_{OLD}" not in names and f"signed_{NEW}" in names
    assert f"input_{OLD}.xlsx" in names                      # never renamed: Excel may hold it open


def test_renamed_runner_still_runs_the_deal(home):
    old_dir = _filed_deal(home)
    _change(old_dir, NEW_VALUES)
    assert runner.invoke(app, ["generate", str(old_dir)]).exit_code == 0
    new_dir = home / COMPANY / "Outgoing" / f"{NEW}_1200kg"
    text = (new_dir / f"run_{NEW}.command").read_text(encoding="utf-8")
    assert "command=generate" in text
    assert (new_dir / f"run_{NEW}.command").stat().st_mode & 0o111


def test_signed_files_for_the_old_number_are_kept_with_a_notice(home):
    old_dir = _filed_deal(home)
    (old_dir / f"signed_{OLD}" / "Agreement_signed.pdf").write_bytes(b"%PDF signed")
    _change(old_dir, NEW_VALUES)
    result = runner.invoke(app, ["generate", str(old_dir)])
    assert result.exit_code == 0, result.output
    new_dir = home / COMPANY / "Outgoing" / f"{NEW}_1200kg"
    assert (new_dir / f"signed_{OLD}" / "Agreement_signed.pdf").read_bytes() == b"%PDF signed"
    assert f"signed_{OLD}" in result.output and "прежнему номеру" in result.output
    assert f"signed_{OLD}" in (new_dir / f"output_{NEW}" / "❗ПРОВЕРЬТЕ.txt").read_text(encoding="utf-8")


def test_same_number_keeps_everything_as_before(home):
    deal_dir = _filed_deal(home)
    result = runner.invoke(app, ["generate", str(deal_dir)])
    assert result.exit_code == 0, result.output
    names = {p.name for p in deal_dir.iterdir()}
    assert {f"output_{OLD}", f"signed_{OLD}", f"run_{OLD}.command"} <= names


def test_missing_output_after_filing_is_a_red_error_not_done(home, monkeypatch):
    deal_dir = _filed_deal(home)
    lost = home / "nowhere"
    monkeypatch.setattr(cli, "_place_deal", lambda deal, d, root: (lost, None))
    result = runner.invoke(app, ["generate", str(deal_dir)])
    assert result.exit_code != 0
    assert "Done" not in result.output
    assert "папка с документами не найдена" in result.output and str(lost / f"output_{OLD}") in result.output


def test_every_run_writes_a_line_to_the_generate_log(home):
    old_dir = _filed_deal(home)
    _change(old_dir, NEW_VALUES)
    assert runner.invoke(app, ["generate", str(old_dir)]).exit_code == 0
    lines = (home / "Data" / "_generate.log").read_text(encoding="utf-8").splitlines()
    assert len(lines) == 2
    last = lines[-1]
    for part in (f"agreement={NEW}", f"from={old_dir}", f"to={home / COMPANY / 'Outgoing' / f'{NEW}_1200kg'}",
                 "output_exists=yes", f"docx={N_DOCX}"):
        assert part in last, part


@pytest.mark.parametrize("platform, tty, env, opened", [
    ("darwin", True, None, True),
    ("darwin", False, None, False),       # CliRunner / pipes: never open Finder
    ("darwin", True, "1", False),         # DCA_NO_OPEN=1
    ("linux", True, None, False),
])
def test_finder_opens_only_from_a_terminal_on_macos(tmp_path, monkeypatch, platform, tty, env, opened):
    calls = []
    monkeypatch.setattr(sys, "platform", platform)
    monkeypatch.setattr(sys.stdout, "isatty", lambda: tty, raising=False)
    if env is None:
        monkeypatch.delenv("DCA_NO_OPEN", raising=False)
    else:
        monkeypatch.setenv("DCA_NO_OPEN", env)
    monkeypatch.setattr(cli.subprocess, "run", lambda args, **kw: calls.append(args))
    cli._open_in_finder(tmp_path)
    assert calls == ([["open", str(tmp_path)]] if opened else [])
