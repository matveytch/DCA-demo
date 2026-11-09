"""`dca backup` / `dca restore`: encrypted versioned backups with restic (2026-10-29).

restic is replaced by a small stub with the same command line (init, backup --json,
forget, check, snapshots --json, restore); one test runs the real restic if installed.
"""

from __future__ import annotations

import json
import os
import shutil
import sys
from datetime import datetime, timedelta

import pytest
from typer.testing import CliRunner

import dca.backup as bk
from dca.cli import app
from dca.config import DATA_DIR
from dca.excel_io import load_deal
from dca.ledger import append_to_ledger
from dca.master_db import init_master
from dca.paths import code_root, resolve_data_root

runner = CliRunner()
FILLED = DATA_DIR / "input_example_FILLED.xlsx"

STUB = r'''
import fnmatch, json, os, shutil, sys
from pathlib import Path

repo = Path(os.environ["RESTIC_REPOSITORY"])
args = sys.argv[1:]
cmd = args[0]
if repo.exists():
    with open(repo / "calls.log", "a", encoding="utf-8") as f:
        f.write(json.dumps(args) + "\n")
if os.environ.get("STUB_FAIL") == cmd:
    print(f"Fatal: stub failure in {cmd}", file=sys.stderr)
    sys.exit(1)
snaps = sorted(repo.glob("snap*"), key=lambda p: int(p.name[4:])) if repo.exists() else []
if cmd == "init":
    repo.mkdir(parents=True, exist_ok=True)
    (repo / "config").write_text("stub")
elif cmd == "backup":
    src = Path(args[1])
    excl = [args[i + 1] for i, a in enumerate(args) if a == "--exclude"]
    def ignore(folder, names):
        out = set()
        for n in names:
            full = str(Path(folder) / n)
            if any(full == e or fnmatch.fnmatch(n, e) for e in excl):
                out.add(n)
        return out
    snap = repo / f"snap{len(snaps) + 1}"
    shutil.copytree(src, snap / "data" / src.relative_to("/"), ignore=ignore)
    (snap / "src").write_text(str(src))
    files = [p for p in (snap / "data").rglob("*") if p.is_file()]
    print(json.dumps({"message_type": "status", "percent_done": 1}))
    print(json.dumps({"message_type": "summary", "snapshot_id": snap.name,
                      "total_files_processed": len(files), "total_bytes_processed": sum(p.stat().st_size for p in files)}))
elif cmd == "snapshots":
    print(json.dumps([{"short_id": s.name, "time": "2026-10-29T19:00:00"} for s in snaps]))
elif cmd == "restore":
    snap = snaps[-1] if args[1] == "latest" else repo / args[1]
    target = Path(args[args.index("--target") + 1])
    shutil.copytree(snap / "data", target, dirs_exist_ok=True)
'''


@pytest.fixture
def restic(tmp_path, monkeypatch):
    """A fake `restic` first on PATH; RESTIC_PASSWORD set so the Keychain is never asked."""
    bin_dir = tmp_path / "bin"
    bin_dir.mkdir()
    stub = bin_dir / "restic"
    stub.write_text(f"#!{sys.executable}\n{STUB}", encoding="utf-8")
    stub.chmod(0o755)
    monkeypatch.setenv("PATH", f"{bin_dir}{os.pathsep}{os.environ['PATH']}")
    monkeypatch.setenv("RESTIC_PASSWORD", "test-only-password")
    monkeypatch.delenv("STUB_FAIL", raising=False)
    calls = []
    monkeypatch.setattr(bk, "store_password", lambda: calls.append("keychain"))
    return calls


@pytest.fixture
def home(tmp_path, monkeypatch):
    """A data root: ledger, master database, a deal, plus things a backup must skip."""
    h = tmp_path / "DCA"
    monkeypatch.setenv("DCA_HOME", str(h))
    append_to_ledger(load_deal(FILLED), h / "Data" / "batch_ledger.xlsx", deal_id="deal-1")
    init_master(h / "Data" / "master_db.xlsx")
    deal = h / "Компания А" / "Outgoing" / "1-2026-1705_1000kg"
    deal.mkdir(parents=True)
    (deal / "input_1-2026-1705.xlsx").write_bytes(b"deal workbook")
    (deal / "~$input_1-2026-1705.xlsx").write_bytes(b"excel lock")       # open in Excel
    (deal / "_сборка_output_1-2026-1705_ab12cd34").mkdir()                  # a crashed build
    (deal / "_сборка_output_1-2026-1705_ab12cd34" / "Agreement.docx").write_bytes(b"half built")
    (h / "Data" / "batch_ledger.xlsx.lock").write_bytes(b"")
    (h / "dca").mkdir()
    (h / "dca" / "code.py").write_text("print('code')")
    (h / ".dca.toml").write_text('code_dir = "dca"\n')
    return h


@pytest.fixture
def drive(tmp_path, monkeypatch):
    gd = tmp_path / "GoogleDrive" / "Мой диск"
    gd.mkdir(parents=True)
    monkeypatch.setattr(bk, "google_drive_dirs", lambda: [gd])
    return gd


def _setup(home, drive):
    result = runner.invoke(app, ["backup", "--setup"])
    assert result.exit_code == 0, result.output
    return drive / "DCA-backups"


def _calls(repo):
    return [json.loads(line) for line in (repo / "calls.log").read_text("utf-8").splitlines()]


# ─── setup ─────────────────────────────────────────────────────

def test_setup_creates_repo_in_google_drive_and_remembers_it(home, drive, restic):
    repo = _setup(home, drive)
    assert (repo / "config").exists()
    assert f'backup_repo = "{repo.resolve()}"' in (home / ".dca.toml").read_text("utf-8")
    assert 'code_dir = "dca"' in (home / ".dca.toml").read_text("utf-8")   # marker kept
    assert restic == ["keychain"]                                              # password asked once
    assert bk.configured_repo(resolve_data_root()) == repo.resolve()


def test_setup_twice_keeps_one_line_and_does_not_ask_again(home, drive, restic):
    _setup(home, drive)
    _setup(home, drive)
    assert (home / ".dca.toml").read_text("utf-8").count("backup_repo") == 1
    assert restic == ["keychain"]


def test_setup_without_restic_says_how_to_install(home, drive, monkeypatch):
    monkeypatch.setattr(bk, "find_restic", lambda: None)
    result = runner.invoke(app, ["backup", "--setup"])
    assert result.exit_code == 1 and "brew install restic" in result.output


def test_setup_refuses_a_repo_inside_the_data(home, drive, restic):
    result = runner.invoke(app, ["backup", "--setup", "--repo", str(home / "Data" / "backups")])
    assert result.exit_code == 1 and "не может лежать внутри папки с данными" in result.output


def test_setup_without_google_drive_asks_for_a_folder(home, restic, monkeypatch):
    monkeypatch.setattr(bk, "google_drive_dirs", lambda: [])
    result = runner.invoke(app, ["backup", "--setup"])
    assert result.exit_code == 1 and "--repo" in result.output


# ─── backup ────────────────────────────────────────────────────

def test_backup_keeps_versions_checks_and_skips_junk(home, drive, restic):
    repo = _setup(home, drive)
    result = runner.invoke(app, ["backup"])
    assert result.exit_code == 0, result.output
    assert "✅ Бэкап готов: версия snap1" in result.output
    assert "открыт в Excel" in result.output and "~$input_1-2026-1705.xlsx" in result.output

    calls = _calls(repo)
    assert [c[0] for c in calls[-3:]] == ["backup", "forget", "check"]
    assert calls[-2] == ["forget", "--tag", "dca", *bk.KEEP, "--prune"]
    copied = {p.relative_to(repo / "snap1" / "data").as_posix().split(home.name + "/", 1)[-1]
              for p in (repo / "snap1" / "data").rglob("*") if p.is_file()}
    assert "Data/batch_ledger.xlsx" in copied and "Data/master_db.xlsx" in copied
    assert "Компания А/Outgoing/1-2026-1705_1000kg/input_1-2026-1705.xlsx" in copied
    assert not any(n.startswith("dca/") or "~$" in n or "_сборка_" in n
                   or n.endswith(".lock") for n in copied), copied

    status = bk.read_status(resolve_data_root())
    assert status["ok"] and status["snapshot"] == "snap1" and status["last_success"] == status["last_attempt"]


def test_failed_backup_is_recorded_and_keeps_the_last_good_one(home, drive, restic, monkeypatch):
    _setup(home, drive)
    assert runner.invoke(app, ["backup"]).exit_code == 0
    good = bk.read_status(resolve_data_root())["last_success"]
    monkeypatch.setenv("STUB_FAIL", "check")
    result = runner.invoke(app, ["backup"])
    assert result.exit_code == 1 and "restic check: Fatal: stub failure in check" in result.output
    status = bk.read_status(resolve_data_root())
    assert status["ok"] is False and status["last_success"] == good and "check" in status["error"]


def test_backup_before_setup_says_what_to_do(home, restic):
    result = runner.invoke(app, ["backup"])
    assert result.exit_code == 1 and "dca backup --setup" in result.output


# ─── «backup is stale» notice ──────────────────────────────────

def test_stale_notice_only_when_needed(home, drive, restic):
    root = resolve_data_root()
    assert bk.stale_notice(root) is None                          # not set up → silent
    _setup(home, drive)
    assert "ни разу" in bk.stale_notice(root)
    runner.invoke(app, ["backup"])
    assert bk.stale_notice(root) is None                          # fresh → silent
    assert "устарел" in bk.stale_notice(root, now=datetime.now() + timedelta(days=bk.STALE_DAYS + 1))
    status = bk.read_status(root)
    status.update(ok=False, error="disk full")
    bk.status_path(root).write_text(json.dumps(status), encoding="utf-8")
    assert "НЕ УДАЛСЯ" in bk.stale_notice(root) and "disk full" in bk.stale_notice(root)


def test_generate_warns_about_a_stale_backup_but_not_in_the_check_file(home, drive, restic):
    from dca.checks import CHECK_FILENAME
    _setup(home, drive)
    assert runner.invoke(app, ["new", "1-2026-2000"]).exit_code == 0
    d = home / "Новые сделки" / "1-2026-2000_pending"
    shutil.copy(FILLED, d / "input_1-2026-2000.xlsx")
    from openpyxl import load_workbook
    wb = load_workbook(d / "input_1-2026-2000.xlsx")
    ws = wb["Сделка"]
    row = next(r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value == "agreement_num")
    ws.cell(row, 3, "1-2026-2000")
    wb.save(d / "input_1-2026-2000.xlsx")
    result = runner.invoke(app, ["generate", str(d)])
    assert "БЭКАП ещё ни разу не выполнен" in result.output
    out = next(home.rglob("output_1-2026-2000"))
    check = out / CHECK_FILENAME
    assert not check.exists() or "БЭКАП" not in check.read_text("utf-8")


# ─── restore ───────────────────────────────────────────────────

def test_restore_goes_to_a_new_folder_only(home, drive, restic, tmp_path):
    _setup(home, drive)
    runner.invoke(app, ["backup"])
    for bad in (home / "restored", home / "Data"):
        result = runner.invoke(app, ["restore", "latest", "--to", str(bad)])
        assert result.exit_code == 1 and "вне" in result.output
    busy = tmp_path / "busy"
    busy.mkdir()
    (busy / "x").write_text("x")
    assert "не пустая" in runner.invoke(app, ["restore", "latest", "--to", str(busy)]).output

    result = runner.invoke(app, ["restore", "snap1", "--to", str(tmp_path / "restored")])
    assert result.exit_code == 0, result.output
    restored = tmp_path.joinpath("restored", *home.parts[1:])
    assert (restored / "Data" / "batch_ledger.xlsx").read_bytes() == (home / "Data" / "batch_ledger.xlsx").read_bytes()


def test_restore_list(home, drive, restic):
    _setup(home, drive)
    runner.invoke(app, ["backup"])
    runner.invoke(app, ["backup"])
    result = runner.invoke(app, ["restore", "--list"])
    assert result.exit_code == 0 and "snap1" in result.output and "snap2" in result.output


def test_test_restore_ok_and_detects_a_damaged_copy(home, drive, restic):
    repo = _setup(home, drive)
    runner.invoke(app, ["backup"])
    result = runner.invoke(app, ["backup", "--test-restore"])
    assert result.exit_code == 0, result.output
    assert "совпадает с данными" in result.output
    assert bk.read_status(resolve_data_root())["last_restore_test"]["ok"] is True

    damaged = next((repo / "snap1").rglob("master_db.xlsx"))
    damaged.write_bytes(b"damaged")
    result = runner.invoke(app, ["backup", "--test-restore"])
    assert result.exit_code == 1
    assert "отличается: Data/master_db.xlsx" in result.output and "не открывается: master_db.xlsx" in result.output


# ─── schedule ──────────────────────────────────────────────────

@pytest.fixture
def agents(tmp_path, monkeypatch):
    d = tmp_path / "LaunchAgents"
    monkeypatch.setenv("DCA_LAUNCH_AGENTS_DIR", str(d))
    calls = []

    class Done:
        returncode, stderr = 0, ""

    monkeypatch.setattr(bk, "_launchctl", lambda *a: calls.append(a) or Done())
    return d, calls


def test_schedule_writes_and_loads_a_launch_agent(home, agents):
    d, calls = agents
    result = runner.invoke(app, ["backup", "--schedule"])
    assert result.exit_code == 0, result.output
    plist = (d / "com.dca.backup.plist").read_text("utf-8")
    assert f"<string>{code_root() / 'dca'}</string><string>backup</string>" in plist
    assert "<key>Hour</key><integer>19</integer>" in plist and "/opt/homebrew/bin" in plist
    assert calls[-1][0] == "bootstrap"
    assert bk.scheduled_launcher() == code_root() / "dca"

    result = runner.invoke(app, ["backup", "--unschedule"])
    assert result.exit_code == 0 and "снято" in result.output and not (d / "com.dca.backup.plist").exists()


def test_doctor_shows_backup_state(home, drive, restic, agents, monkeypatch):
    assert "backups are not set up" in runner.invoke(app, ["doctor"]).output
    _setup(home, drive)
    runner.invoke(app, ["backup"])
    runner.invoke(app, ["backup", "--schedule"])
    out = runner.invoke(app, ["doctor"]).output
    assert "✓ last backup:" in out and "✓ every day at 19:00" in out
    assert "test restore: never" in out


# ─── the real restic (only if installed) ───────────────────────

@pytest.mark.skipif(bk.find_restic() is None, reason="restic not installed")
def test_real_restic_backup_and_restore_check(home, tmp_path, monkeypatch):
    monkeypatch.setenv("RESTIC_PASSWORD", "test-only-password")
    root = resolve_data_root()
    repo = tmp_path / "repo"
    bk.init_repo(repo)
    bk.save_repo(root, repo)
    status = bk.run_backup(root)
    assert status["ok"] and status["files"] >= 3
    report = bk.restore_check(root)
    assert report["ok"], report
