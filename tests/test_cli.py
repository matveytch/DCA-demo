"""End-to-end CLI tests (typer's CliRunner) on the demo data: AC1–AC3, AC6, AC13–AC22."""

from __future__ import annotations

import json
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

import pytest
from openpyxl import load_workbook
from typer.testing import CliRunner

import dca.cli as cli
from dca.cli import app
from dca.log import LOG_SCHEMA_VERSION
from tests.helpers import COMPANY, FILLED, NUM, docx_text, set_archive_value, set_deal_values

runner = CliRunner()
DOCS = {f"Agreement_{NUM}.docx", f"Annex_1_{NUM}.docx", f"Annex_2_{NUM}.docx", f"CoverList_{NUM}.docx"}


def _new(home, num=NUM, command="new"):
    result = runner.invoke(app, [command, num])
    assert result.exit_code == 0, result.output
    folder = f"{num}_pending" if command == "new" else f"{num}_incoming_pending"
    return home / "Новые сделки" / folder


def _filled(home, num=NUM):
    d = _new(home, num)
    shutil.copy(FILLED, d / f"input_{num}.xlsx")
    return d


def _generated(home):
    d = _filled(home)
    result = runner.invoke(app, ["generate", str(d)])
    assert result.exit_code == 0, result.output
    return home / COMPANY / "Outgoing" / f"{NUM}_1000kg"


def _scans(home):
    """Scans of the example's batches and the company profile, where the operator keeps them."""
    scans = home / "Data" / "scans" / "П-001"
    scans.mkdir(parents=True)
    (scans / "соглашение В-17.pdf").write_bytes(b"%PDF 17")
    (scans / "приложение 1 к В-17.pdf").write_bytes(b"%PDF 17-1")
    (home / "Data" / "scans" / "соглашение В-21.pdf").write_bytes(b"%PDF 21")
    (home / COMPANY).mkdir(parents=True, exist_ok=True)
    (home / COMPANY / "company_profile.pdf").write_bytes(b"%PDF profile")


# ─── basics ────────────────────────────────────────────────────

def test_hello_and_help():
    assert "DCA" in runner.invoke(app, ["hello"]).output
    help_text = runner.invoke(app, ["--help"]).output
    for command in ("new", "new-incoming", "generate", "generate-incoming", "amend", "transfer",
                    "from-log", "check-ledger", "init-db", "refresh-db", "rewrite-runners", "backup", "doctor"):
        assert command in help_text


def test_new_creates_folder_identity_input_and_portable_runners(home):
    d = _new(home)
    assert (d / f"input_{NUM}.xlsx").exists()
    assert json.loads((d / ".dca-deal.json").read_text(encoding="utf-8"))["kind"] == "outgoing"
    for name, command in ((f"run_{NUM}.command", "generate"), (f"transfer_{NUM}.command", "transfer"),
                          (f"amend_{NUM}.command", "amend")):
        text = (d / name).read_text(encoding="utf-8")
        assert f"command={command}" in text and f'-m dca {command} "$DEAL_DIR"' in text
        assert not [ln for ln in text.splitlines() if ln.strip().startswith("cd ")]   # no project path
        assert (d / name).stat().st_mode & 0o111


def test_new_refuses_an_existing_folder_and_a_duplicate_number(home):
    _new(home)
    again = runner.invoke(app, ["new", NUM])
    assert again.exit_code == 1 and "already exists" in again.output
    other = runner.invoke(app, ["new-incoming", NUM])
    assert other.exit_code == 1 and f"agreement number {NUM} already exists" in other.output


# ─── generate ──────────────────────────────────────────────────

def test_generate_full_pipeline(home):
    _scans(home)
    deal_dir = _generated(home)
    out = deal_dir / f"output_{NUM}"
    assert {p.name for p in out.glob("*.docx")} == DOCS
    assert (out / "log.json").exists()
    assert (out / f"transfer_docs_{NUM}" / f"TransferAct_{NUM}.docx").exists()
    attachments = sorted(p.name for p in (out / f"attachments_{NUM}").iterdir())
    assert attachments == [
        "2. копия соглашения №В-17-2025 от 15.09.2025.pdf",
        "3. копия приложения №1 от 20.09.2025 к соглашению №В-17-2025 от 15.09.2025.pdf",
        "4. копия соглашения №В-21-2025 от 01.11.2025.pdf",
        "5. карточка компании.pdf",
    ]
    with zipfile.ZipFile(out / f"Package_{NUM}.zip") as z:
        members = set(z.namelist())
    assert DOCS <= members and f"attachments_{NUM}/5. карточка компании.pdf" in members
    assert not any(m.endswith("log.json") or "transfer_docs" in m for m in members)
    assert not (out / "❗ПРОВЕРЬТЕ.txt").exists()              # complete deal → nothing to check
    assert (home / "Data" / "batch_ledger.xlsx").exists()
    assert (deal_dir / f"signed_{NUM}").is_dir()


def test_missing_scans_are_notices_in_the_check_file(home):
    deal_dir = _generated(home)
    check = (deal_dir / f"output_{NUM}" / "❗ПРОВЕРЬТЕ.txt").read_text(encoding="utf-8")
    assert "Нет файла для пакета приложений: соглашение В-17.pdf" in check
    assert f"Нет файла для пакета приложений: {COMPANY}/company_profile.pdf" in check


def test_invalid_input_exits_1_with_one_russian_message_and_no_documents(home):
    d = _filled(home)
    set_deal_values(d / f"input_{NUM}.xlsx", {"place_ru": None, "quantity_kg": None})
    result = runner.invoke(app, ["generate", str(d)])
    assert result.exit_code == 1
    assert "в Excel не хватает данных или они неверные" in result.output and "Traceback" not in result.output
    assert not list(d.glob("output_*"))


def test_party_a_must_be_our_company(home):
    d = _filled(home)
    set_archive_value(d / f"input_{NUM}.xlsx", "Стороны А", "party_a_folder_name", "alpha", None)
    result = runner.invoke(app, ["generate", str(d)])
    assert result.exit_code == 1 and "has no party_a_folder_name" in result.output


def test_second_generate_archives_the_previous_output(home):
    deal_dir = _generated(home)
    first = (deal_dir / f"output_{NUM}" / "log.json").read_bytes()
    assert runner.invoke(app, ["generate", str(deal_dir)]).exit_code == 0
    archived = deal_dir / "_archive" / "v1" / "log.json"
    assert archived.read_bytes() == first


def test_failure_while_building_leaves_the_previous_output(home, monkeypatch):
    deal_dir = _generated(home)
    before = sorted(p.name for p in (deal_dir / f"output_{NUM}").iterdir())

    def boom(*a, **k):
        raise RuntimeError("disk full")
    monkeypatch.setattr(cli, "dump_log", boom)
    result = runner.invoke(app, ["generate", str(deal_dir)])
    assert result.exit_code != 0
    assert sorted(p.name for p in (deal_dir / f"output_{NUM}").iterdir()) == before
    assert not list(deal_dir.glob("_сборка_*")) and not (deal_dir / "_archive").exists()


def test_failed_publish_restores_the_previous_output(home, monkeypatch):
    deal_dir = _generated(home)
    real_rename = Path.rename

    def failing_rename(self, target):
        if self.name.startswith("_сборка_"):
            raise OSError("rename failed")
        return real_rename(self, target)
    monkeypatch.setattr(Path, "rename", failing_rename)
    result = runner.invoke(app, ["generate", str(deal_dir)])
    assert result.exit_code != 0
    assert (deal_dir / f"output_{NUM}" / f"Agreement_{NUM}.docx").exists()


def test_excel_open_keeps_the_deal_where_it_is(home):
    d = _filled(home)
    (d / f"~$input_{NUM}.xlsx").write_bytes(b"owner")
    result = runner.invoke(app, ["generate", str(d)])
    assert result.exit_code == 0, result.output
    assert "СДЕЛКА НЕ ПЕРЕНЕСЕНА" in result.output
    assert (d / f"output_{NUM}" / f"Agreement_{NUM}.docx").exists()


def test_letter_date_is_kept_from_the_previous_generation(home):
    deal_dir = _generated(home)
    log = deal_dir / f"output_{NUM}" / "log.json"
    data = json.loads(log.read_text(encoding="utf-8"))
    data["deal"]["letter_date"] = "2026-01-02"
    log.write_text(json.dumps(data), encoding="utf-8")
    assert runner.invoke(app, ["generate", str(deal_dir)]).exit_code == 0
    assert "Исх. № 7 от 02.01.2026" in docx_text(deal_dir / f"output_{NUM}" / f"CoverList_{NUM}.docx")


def test_stale_signatures_are_reported(home):
    deal_dir = _generated(home)
    (deal_dir / f"signed_{NUM}" / "Agreement_signed.pdf").write_bytes(b"%PDF")
    assert runner.invoke(app, ["generate", str(deal_dir)]).exit_code == 0     # binds the signatures
    set_deal_values(deal_dir / f"input_{NUM}.xlsx", {"place_ru": "Город Е"})
    result = runner.invoke(app, ["generate", str(deal_dir)])
    assert result.exit_code == 0 and "ПРЕДЫДУЩЕЙ версии документов" in result.output


def test_a_running_command_locks_the_deal(home):
    from filelock import FileLock

    d = _filled(home)
    with FileLock(str(d / ".dca.lock"), timeout=0):
        result = runner.invoke(app, ["generate", str(d)])
    assert result.exit_code == 1 and "Another dca command is working" in result.output


def test_renamed_folder_does_not_hide_a_duplicate_number(home):
    deal_dir = _generated(home)
    deal_dir.rename(deal_dir.with_name("renamed"))
    result = runner.invoke(app, ["new", NUM])
    assert result.exit_code == 1 and "already exists" in result.output


# ─── incoming deals ────────────────────────────────────────────

def test_new_incoming_hides_outgoing_rows_and_the_transfer_sheet(home):
    d = _new(home, command="new-incoming")
    xlsx = d / f"input_{NUM}.xlsx"
    wb = load_workbook(xlsx)
    ws = wb["Сделка"]
    hidden = {ws.cell(r, 1).value for r in range(1, ws.max_row + 1) if ws.row_dimensions[r].hidden}
    assert {"variant", "party_c_name", "letter_num", "batch_key_1", "batch_use_kg_5"} <= hidden
    assert not hidden & {"agreement_num", "quantity_kg", "party_b_key", "place_ru"}
    assert wb["Передача"].sheet_state == "hidden"
    assert "conditionalFormatting" in zipfile.ZipFile(xlsx).read(cli._deal_sheet_member(xlsx)).decode()
    assert "command=generate-incoming" in (d / f"run_{NUM}.command").read_text(encoding="utf-8")
    assert not (d / f"transfer_{NUM}.command").exists()


def test_generate_incoming_pipeline(home):
    d = _new(home, command="new-incoming")
    xlsx = d / f"input_{NUM}.xlsx"
    shutil.copy(FILLED, xlsx)
    set_deal_values(xlsx, {"party_b_key": "alpha-b", "variant": None, "batch_key_1": None,
                           "batch_use_kg_1": None, "batch_key_2": None, "batch_use_kg_2": None})
    result = runner.invoke(app, ["generate-incoming", str(d)])
    assert result.exit_code == 0, result.output
    out = home / COMPANY / "Incoming" / f"{NUM}_1000kg" / f"output_{NUM}"
    assert {p.name for p in out.glob("*.docx")} == DOCS - {f"CoverList_{NUM}.docx"}
    assert "получает от Стороны А" in docx_text(out / f"Agreement_{NUM}.docx")
    assert not (home / "Data" / "batch_ledger.xlsx").exists()


def test_the_command_must_match_the_deal_kind(home):
    d = _filled(home)
    result = runner.invoke(app, ["generate-incoming", str(d)])
    assert result.exit_code == 1 and "This is an outgoing deal" in result.output


# ─── amend / transfer / from-log ───────────────────────────────

def test_amend_writes_annexes_and_the_amendment(home):
    d = _filled(home)
    result = runner.invoke(app, ["amend", str(d)])
    assert result.exit_code == 0, result.output
    out = d / f"amendment_{NUM}"
    assert {p.name for p in out.iterdir()} == {f"Amendment_{NUM}.docx", f"Annex_1_{NUM}.docx", f"Annex_2_{NUM}.docx"}
    assert "стало 1 200 кг" in docx_text(out / f"Amendment_{NUM}.docx")
    assert runner.invoke(app, ["amend", str(d)]).exit_code == 0
    assert (d / "_archive" / "amendment_v1").is_dir()


def test_amend_refuses_an_incomplete_sheet_and_keeps_the_previous_result(home):
    d = _filled(home)
    assert runner.invoke(app, ["amend", str(d)]).exit_code == 0
    set_deal_values(d / f"input_{NUM}.xlsx", {"amendment_num": None}, sheet="Изменение")
    result = runner.invoke(app, ["amend", str(d)])
    assert result.exit_code == 1 and "amendment_num is required" in result.output
    assert not (d / "_archive").exists()


def test_amend_adds_a_missing_sheet(home, tmp_path):
    d = _filled(home)
    xlsx = d / f"input_{NUM}.xlsx"
    wb = load_workbook(xlsx)
    del wb["Изменение"]
    wb.save(xlsx)
    result = runner.invoke(app, ["amend", str(d)])
    assert result.exit_code == 1 and "Added sheet «Изменение»" in result.output
    assert "Изменение" in load_workbook(xlsx).sheetnames


def test_transfer_writes_the_act_only(home):
    d = _filled(home)
    result = runner.invoke(app, ["transfer", str(d)])
    assert result.exit_code == 0, result.output
    out = d / f"output_{NUM}"
    assert [p.name for p in out.rglob("*.docx")] == [f"TransferAct_{NUM}.docx"]


def test_transfer_date_after_due_date_is_a_warning(home):
    d = _filled(home)
    set_deal_values(d / f"input_{NUM}.xlsx", {"transfer_date": "20.08.2026"}, sheet="Передача")
    result = runner.invoke(app, ["transfer", str(d)])
    assert result.exit_code == 0 and "позже срока исполнения 17.07.2026" in result.output


def test_unfilled_transfer_sheet_is_not_a_notice(home):
    d = _filled(home)
    xlsx = d / f"input_{NUM}.xlsx"
    set_deal_values(xlsx, {"transfer_num": None, "handed_by_ru": None}, sheet="Передача")
    _scans(home)
    result = runner.invoke(app, ["generate", str(d)])
    assert result.exit_code == 0 and "not filled yet — skipped" in result.output
    assert not (home / COMPANY / "Outgoing" / f"{NUM}_1000kg" / f"output_{NUM}" / "❗ПРОВЕРЬТЕ.txt").exists()


def test_filled_transfer_sheet_with_an_error_is_a_notice(home):
    d = _filled(home)
    set_deal_values(d / f"input_{NUM}.xlsx", {"transfer_date": "01.01.2026"}, sheet="Передача")
    result = runner.invoke(app, ["generate", str(d)])
    assert result.exit_code == 0, result.output
    check = home / COMPANY / "Outgoing" / f"{NUM}_1000kg" / f"output_{NUM}" / "❗ПРОВЕРЬТЕ.txt"
    assert "Акт передачи не создан" in check.read_text(encoding="utf-8")


def test_from_log_reports_changed_templates(home):
    deal_dir = _generated(home)
    log = deal_dir / f"output_{NUM}" / "log.json"
    data = json.loads(log.read_text(encoding="utf-8"))
    data["_metadata"]["templates_sha256"]["agreement.docx"] = "0" * 64
    log.write_text(json.dumps(data), encoding="utf-8")
    result = runner.invoke(app, ["from-log", str(log)])
    assert result.exit_code == 0, result.output
    assert "Шаблоны изменились после сохранения log.json: agreement.docx" in result.output


def test_transfer_is_refused_for_an_incoming_deal(home):
    d = _new(home, command="new-incoming")
    result = runner.invoke(app, ["transfer", str(d)])
    assert result.exit_code == 1 and "outgoing deals" in result.output


def test_from_log_re_renders_and_refuses_a_newer_schema(home):
    deal_dir = _generated(home)
    log = deal_dir / f"output_{NUM}" / "log.json"
    result = runner.invoke(app, ["from-log", str(log)])
    assert result.exit_code == 0, result.output
    assert (deal_dir / "_archive" / "v1").is_dir()

    log = deal_dir / f"output_{NUM}" / "log.json"
    data = json.loads(log.read_text(encoding="utf-8"))
    data["_metadata"]["schema_version"] = LOG_SCHEMA_VERSION + 1
    log.write_text(json.dumps(data), encoding="utf-8")
    newer = runner.invoke(app, ["from-log", str(log)])
    assert newer.exit_code == 1 and "newer than this DCA supports" in newer.output


def test_from_log_bad_input(tmp_path):
    assert runner.invoke(app, ["from-log", str(tmp_path / "none.json")]).exit_code == 1
    bad = tmp_path / "log.json"
    bad.write_text("not json")
    result = runner.invoke(app, ["from-log", str(bad)])
    assert result.exit_code == 1 and "Traceback" not in result.output


# ─── runners / doctor ──────────────────────────────────────────

def test_rewrite_runners_keeps_each_deal_command(home):
    out = _new(home, "1-2026-0001")
    inc = _new(home, "1-2026-0002", command="new-incoming")
    for d in (out, inc):
        for f in d.glob("*.command"):
            f.write_text("old runner\n")
    (out / "run_1-2026-0001.command").write_text("# dca-runner v1 command=generate\n")
    (inc / "run_1-2026-0002.command").write_text("# dca-runner v1 command=generate-incoming\n")
    result = runner.invoke(app, ["rewrite-runners", "--all"])
    assert result.exit_code == 0, result.output
    assert "command=generate\n" in (out / "run_1-2026-0001.command").read_text(encoding="utf-8")
    assert "command=generate-incoming" in (inc / "run_1-2026-0002.command").read_text(encoding="utf-8")
    assert cli.RUNNER_TAG in (out / "transfer_1-2026-0001.command").read_text(encoding="utf-8")


def test_rewrite_runners_argument_errors(home, tmp_path):
    assert runner.invoke(app, ["rewrite-runners"]).exit_code == 1
    assert runner.invoke(app, ["rewrite-runners", str(tmp_path), "--all"]).exit_code == 1
    assert runner.invoke(app, ["rewrite-runners", str(tmp_path / "none")]).exit_code == 1


def test_doctor_flags_outdated_runners(home, monkeypatch):
    from dca.paths import code_root

    monkeypatch.setattr(cli, "_editable_pointers", lambda: [(Path("__editable__.dca-0.1.0.pth"), code_root() / "src")])
    old = home / "Новые сделки" / "1-2026-OLD_pending"
    old.mkdir(parents=True)
    (old / "run_1-2026-OLD.command").write_text('dca generate "$DEAL_DIR"\n')
    result = runner.invoke(app, ["doctor"])
    assert result.exit_code == 0, result.output
    assert "outdated runners: Новые сделки/1-2026-OLD_pending" in result.output
    assert "templates + input_template_EMPTY.xlsx present" in result.output


@pytest.mark.skipif(sys.platform == "win32", reason="bash runner")
def test_runner_works_in_a_moved_copy_without_dca_on_path(tmp_path):
    """Copy the code to a new path (spaces + Cyrillic), run the runner with NO `dca` on PATH:
    it finds the code by walking up from the deal folder and renders."""
    from dca.config import DATA_DIR, TEMPLATES_DIR
    from dca.paths import code_root

    moved = tmp_path / "Новая папка" / "DCA copy"
    shutil.copytree(code_root() / "src" / "dca", moved / "src" / "dca", ignore=shutil.ignore_patterns("__pycache__"))
    shutil.copytree(TEMPLATES_DIR, moved / "templates")
    (moved / "data").mkdir()
    shutil.copy(DATA_DIR / "input_template_EMPTY.xlsx", moved / "data")
    (moved / "dca.toml").write_text("")

    bin_dir = tmp_path / "bin"
    bin_dir.mkdir()
    (bin_dir / "python3").symlink_to(sys.executable)
    env = {"PATH": f"{bin_dir}:/usr/bin:/bin", "HOME": str(tmp_path)}
    subprocess.run([sys.executable, "-m", "dca", "new", NUM], cwd=moved,
                   env={**env, "PYTHONPATH": str(moved / "src")}, check=True, capture_output=True)
    deal_dir = moved / "Новые сделки" / f"{NUM}_pending"
    shutil.copy(FILLED, deal_dir / f"input_{NUM}.xlsx")

    proc = subprocess.run(["bash", str(deal_dir / f"run_{NUM}.command")], env=env, input="x",
                          capture_output=True, text=True, timeout=120)
    assert proc.returncode == 0, proc.stdout + proc.stderr
    out = moved / COMPANY / "Outgoing" / f"{NUM}_1000kg" / f"output_{NUM}"
    assert (out / f"Agreement_{NUM}.docx").exists() and (out / f"Package_{NUM}.zip").exists()
    assert (moved / "Data" / "batch_ledger.xlsx").exists()
