"""CLI entry point (typer).

Commands:
    dca new <agreement_num>            — create a new outgoing deal folder + copy empty template
    dca new-incoming <agreement_num>   — the same for an incoming deal
    dca generate <deal_dir>            — render the outgoing package from input_<num>.xlsx
    dca generate-incoming <deal_dir>   — render the incoming package (agreement + annexes)
    dca amend <deal_dir>               — render annexes + amendment (sheet «Изменение»)
    dca transfer <deal_dir>            — render the transfer act (sheet «Передача»)
    dca from-log <log_path>            — re-render from a saved log.json
    dca check-ledger                   — check the batch ledger and rebuild «Баланс»
    dca init-db                        — create the empty master database Data/master_db.xlsx
    dca upgrade-db                     — add new template fields to an existing master database
    dca refresh-db <deal_dir>          — re-copy the archive sheets of a deal from the master database
    dca rewrite-runners <deal_dir>     — refresh the *.command runners (use after upgrades)
    dca rewrite-runners --all          — refresh runners in every deal folder
    dca backup / dca restore           — encrypted versioned backup of the data (restic)
    dca doctor                         — check code/templates/data locations, print fixes
    dca hello                          — smoke test

Global option `--home <dir>` picks the data folder explicitly (see `dca.paths`).
"""

from __future__ import annotations

import datetime as _dt
import math
import re
import shutil
import subprocess
import sys
import traceback
import zipfile
from pathlib import Path

import typer
from pydantic import ValidationError

from dca import signatures
from dca.config import (
    AMENDMENT_RUN_COMMAND_FILENAME,
    EMPTY_TEMPLATE,
    INPUT_FILENAME,
    OUTPUT_DIR,
    OUTPUT_FILES,
    RUN_COMMAND_FILENAME,
    TRANSFER_DIR,
    TRANSFER_RUN_COMMAND_FILENAME,
)
from dca.deal_identity import (
    DealIdentity,
    known_number,
    new_identity,
    number_key,
    read_identity,
    write_identity,
)
from dca.excel_io import load_deal
from dca.log import dump_log
from dca.models import Deal
from dca.paths import DataRoot, code_root, resolve_data_root
from dca.render import render_all, render_transfer
from dca.utils.slug import slug
from dca.xlsx_patch import (
    AMENDMENT_SHEET,
    TRANSFER_SHEET,
    add_amendment_sheet,
    add_transfer_sheet,
    hide_sheet,
    sheet_names,
)

app = typer.Typer(
    help="DCA — Document Creation Automation. Generate document packages from one Excel file.",
    no_args_is_help=True,
)

# Set from the global `--home` option on every invocation (see `_main`).
_home_override: Path | None = None


@app.callback()
def _main(
    home: Path | None = typer.Option(  # noqa: B008
        None,
        "--home",
        help="Data folder (company folders, «Новые сделки», Data/). "
        "Overrides $DCA_HOME and the dca.toml lookup.",
    ),
) -> None:
    global _home_override
    _home_override = home


def _data_root(start: Path | None = None) -> DataRoot:
    """Data root for a command; `start` is the deal folder (None → cwd)."""
    return resolve_data_root(start, _home_override)


# ─── Commands ──────────────────────────────────────────────────────────────

@app.command()
def hello() -> None:
    """Smoke-test that the CLI is wired up."""
    typer.echo("DCA v0.1.0 — hello!")


def _create_deal_folder(agreement_num: str, kind: str) -> tuple[Path, Path, str]:
    """Shared part of `new` / `new-incoming`: folder, identity, input Excel, database snapshot."""
    if not EMPTY_TEMPLATE.exists():
        typer.echo(f"❌ Empty template not found: {EMPTY_TEMPLATE}", err=True)
        raise typer.Exit(code=1)

    num_slug = slug(agreement_num)
    folder_name = f"{num_slug}_pending" if kind == "outgoing" else f"{num_slug}_incoming_pending"
    _refuse_duplicate(agreement_num)
    deal_dir = _data_root().new_deals_dir / folder_name
    if deal_dir.exists():
        typer.echo(f"❌ Folder already exists: {deal_dir}", err=True)
        raise typer.Exit(code=1)

    deal_dir.mkdir(parents=True)
    write_identity(deal_dir, new_identity(agreement_num, kind))
    input_path = deal_dir / INPUT_FILENAME.format(num=num_slug)
    shutil.copy(EMPTY_TEMPLATE, input_path)
    _snapshot_master_db(input_path)
    return deal_dir, input_path, num_slug


@app.command()
def new(agreement_num: str) -> None:
    """Create a new OUTGOING deal folder + copy the empty input template.

    Creates `Новые сделки/<agreement_num>_pending/` with `input_<agreement_num>.xlsx`
    (the archive sheets get a snapshot of the master database) and double-click
    runners: run_<num>.command (the package), transfer_<num>.command (the transfer
    act) and amend_<num>.command (an amendment).

    After you fill the input file, run `dca generate <deal_dir>` — it files the
    deal into <company>/Outgoing/<num>_<quantity>kg.
    """
    deal_dir, input_path, num_slug = _create_deal_folder(agreement_num, "outgoing")
    _write_runners(deal_dir, num_slug)
    _write_transfer_runners(deal_dir, num_slug)
    _write_amendment_runners(deal_dir, num_slug)

    typer.echo(f"✓ Created deal folder: {deal_dir}")
    typer.echo(f"  → fill {input_path} in Excel")
    typer.echo(f"  → then run: dca generate {deal_dir}")
    typer.echo(f"  (or double-click {deal_dir / RUN_COMMAND_FILENAME.format(num=num_slug)} on macOS)")
    typer.echo(f"  → for the transfer act: fill sheet «Передача», then double-click "
               f"{TRANSFER_RUN_COMMAND_FILENAME.format(num=num_slug)}")


@app.command(name="new-incoming")
def new_incoming(agreement_num: str) -> None:
    """Create a new INCOMING deal folder (agreement + annexes only).

    Same input template as `dca new`, but the rows an incoming deal doesn't need
    (variant, party C, cover list, batches) are hidden, and so is the «Передача»
    sheet. The runner calls `dca generate-incoming`.
    """
    deal_dir, input_path, num_slug = _create_deal_folder(agreement_num, "incoming")

    # Hide outgoing-only rows on «Сделка» (zip/XML patch — keeps the archive
    # dropdowns/validation/CF that an openpyxl save would destroy).
    try:
        _hide_outgoing_rows(input_path)
    except Exception as e:  # don't lose the created folder over a cosmetic step
        typer.echo(f"  ⚠ could not hide outgoing-only rows ({e}); template left as-is", err=True)
    try:  # incoming deals have no transfer act
        hide_sheet(input_path, TRANSFER_SHEET)
    except Exception as e:
        typer.echo(f"  ⚠ could not hide sheet «{TRANSFER_SHEET}» ({e})", err=True)

    _write_runners(deal_dir, num_slug, command="generate-incoming")
    _write_amendment_runners(deal_dir, num_slug)

    typer.echo(f"✓ Created incoming deal folder: {deal_dir}")
    typer.echo(f"  → fill {input_path} in Excel")
    typer.echo(f"  → then run: dca generate-incoming {deal_dir}")
    typer.echo(f"  (or double-click {deal_dir / RUN_COMMAND_FILENAME.format(num=num_slug)} on macOS)")


def _generate(deal_dir: str, kind: str) -> None:
    """Shared body of `generate` / `generate-incoming`."""
    deal_dir_path = Path(deal_dir).resolve()
    identity = _open_deal(deal_dir_path, kind)
    root = _data_root(deal_dir_path)
    input_xlsx = _find_input_file(deal_dir_path)
    if input_xlsx is None:
        typer.echo(f"❌ No input_*.xlsx found in {deal_dir_path}", err=True)
        raise typer.Exit(code=1)

    typer.echo(f"📖 Loading {input_xlsx.name} ({kind} deal)...")
    try:
        deal = load_deal(input_xlsx, deal_kind=kind)
    except Exception as e:
        raise _exit_with_error(e, "Failed to load deal") from e

    typer.echo(f"  ✓ agreement_num={deal.agreement_num}, party B={deal.party_b_name_eng}")
    _require_company_folder(deal)
    old_num = identity.agreement_num if identity else None
    identity = _confirm_identity(deal, deal_dir_path, identity)

    excel_open = _excel_open_notice(input_xlsx)
    notices = _run_pipeline(deal, deal_dir_path, root, excel_open, input_xlsx=input_xlsx,
                            deal_id=identity.deal_id, old_agreement_num=old_num)
    _file_and_finish(deal, deal_dir_path, root, input_xlsx, bool(excel_open), notices)


@app.command()
def generate(deal_dir: str) -> None:
    """Render an OUTGOING deal from <deal_dir>/input_<num>.xlsx.

    Pipeline:
      1. Find input_*.xlsx in deal_dir, load + validate
      2. Build in a temp folder: agreement, annex×N, cover list, attachment pack,
         log.json, Package_<num>.zip; the transfer act when «Передача» is filled
      3. Publish: previous output_<num>/ → _archive/vN/, the new one in its place
      4. Update the batch ledger, write ❗ПРОВЕРЬТЕ.txt when there is something to check
      5. File the deal into <company>/Outgoing/<num>_<quantity>kg
    """
    _generate(deal_dir, "outgoing")


@app.command(name="generate-incoming")
def generate_incoming(deal_dir: str) -> None:
    """Render an INCOMING deal: agreement + annex×N only.

    No attachment pack and no batch ledger (those are outgoing-only). Otherwise
    mirrors `dca generate`: archives previous output, creates signed_<num>/,
    writes log.json, builds the ZIP and files the deal into <company>/Incoming/.
    """
    _generate(deal_dir, "incoming")


@app.command()
def amend(deal_dir: str) -> None:
    """Render ONLY the annex(es) + the amendment.

    A focused one-off: loads <deal_dir>/input_<num>.xlsx (with relaxed,
    incoming-style validation) and writes into <deal_dir>/amendment_<num>/:
      - Annex(es) — fully filled;
      - Amendment_<num>.docx — from the «Изменение» sheet (number, date,
        old/new quantity and price); totals, words and clause numbers are
        computed. An empty or incomplete sheet → nothing is written.

    A deal whose Excel has no «Изменение» sheet gets it added and is asked to fill it.
    Does NOT touch output_<num>/, the ledger, the attachment pack or the folder name.
    """
    from dca.config import AMENDMENT_DIR, AMENDMENT_OUTPUT
    from dca.render import render_amendment, render_annex_set

    deal_dir_path = Path(deal_dir).resolve()
    _open_deal(deal_dir_path, None)
    input_xlsx = _find_input_file(deal_dir_path)
    if input_xlsx is None:
        typer.echo(f"❌ No input_*.xlsx found in {deal_dir_path}", err=True)
        raise typer.Exit(code=1)

    if AMENDMENT_SHEET not in sheet_names(input_xlsx):
        if _excel_has_open(input_xlsx):
            typer.echo(f"❌ {input_xlsx.name} is open in Excel — close it and run again "
                       f"(the sheet «{AMENDMENT_SHEET}» has to be added to it).", err=True)
            raise typer.Exit(code=1)
        add_amendment_sheet(input_xlsx)
        typer.echo(f"📄 Added sheet «{AMENDMENT_SHEET}» to {input_xlsx.name}.")
        typer.echo("   Fill it (number, date, new quantity, price if it changes) and run again.")
        raise typer.Exit(code=1)

    typer.echo(f"📖 Loading {input_xlsx.name}...")
    try:
        deal = load_deal(input_xlsx, deal_kind="incoming")
    except Exception as e:
        raise _exit_with_error(e, "Failed to load deal") from e

    # Render before touching the folder: a refused sheet leaves the previous
    # amendment_<num>/ exactly as it was.
    typer.echo("🖨  Rendering annex(es) + amendment...")
    try:
        amendment_bytes = render_amendment(deal)
        rendered = render_annex_set(deal)
    except Exception as e:
        raise _exit_with_error(e, "Cannot render the amendment") from e

    num_slug = slug(deal.agreement_num)
    out_dir = deal_dir_path / AMENDMENT_DIR.format(num=num_slug)
    # Keep the previous run (and any hand edits in Word) instead of mixing old and
    # new files in one folder.
    from dca.archive import archive_previous_output
    prev = archive_previous_output(deal_dir_path, out_dir.name, prefix="amendment_v")
    if prev is not None:
        typer.echo(f"📦 Previous {out_dir.name}/ → _archive/{prev.name}/")
    out_dir.mkdir(parents=True, exist_ok=True)

    rendered[AMENDMENT_OUTPUT.format(num=num_slug)] = amendment_bytes
    for filename, content in rendered.items():
        (out_dir / filename).write_bytes(content)
        typer.echo(f"  ✓ {filename}  ({len(content) // 1024} KB)")

    typer.echo(f"\n✅ Done. Open: {out_dir}")


@app.command()
def transfer(deal_dir: str) -> None:
    """Render the transfer act from the «Передача» sheet.

    Writes output_<num>/transfer_docs_<num>/TransferAct_<num>.docx. Does NOT touch
    the rest of output_<num>/, the ledger, the ZIP or the folder name. Outgoing deals only.

    A deal whose Excel has no «Передача» sheet gets it added and is asked to fill it.
    """
    deal_dir_path = Path(deal_dir).resolve()
    identity = _open_deal(deal_dir_path, None)
    if _deal_kind(deal_dir_path, identity) == "incoming":
        typer.echo("❌ The transfer act is for outgoing deals (this one is incoming).", err=True)
        raise typer.Exit(code=1)
    input_xlsx = _find_input_file(deal_dir_path)
    if input_xlsx is None:
        typer.echo(f"❌ No input_*.xlsx found in {deal_dir_path}", err=True)
        raise typer.Exit(code=1)

    if TRANSFER_SHEET not in sheet_names(input_xlsx):
        if _excel_has_open(input_xlsx):
            typer.echo(f"❌ {input_xlsx.name} is open in Excel — close it and run again "
                       f"(the sheet «{TRANSFER_SHEET}» has to be added to it).", err=True)
            raise typer.Exit(code=1)
        add_transfer_sheet(input_xlsx)
        typer.echo(f"📄 Added sheet «{TRANSFER_SHEET}» to {input_xlsx.name}.")
        typer.echo("   Fill it (number, date, places, representatives) and run again.")
        raise typer.Exit(code=1)

    typer.echo(f"📖 Loading {input_xlsx.name}...")
    try:
        deal = load_deal(input_xlsx, deal_kind="outgoing")
    except Exception as e:
        raise _exit_with_error(e, "Failed to load deal") from e
    try:
        rendered = render_transfer(deal)
    except Exception as e:
        raise _exit_with_error(e, "Cannot render the transfer act") from e

    num_slug = slug(deal.agreement_num)
    out_dir = deal_dir_path / OUTPUT_DIR.format(num=num_slug) / TRANSFER_DIR.format(num=num_slug)
    _write_transfer_docs(rendered, out_dir)
    # Terminal only: ❗ПРОВЕРЬТЕ.txt belongs to generate and is not overwritten here.
    from dca.checks import transfer_date_warnings
    for w in transfer_date_warnings(deal):
        typer.secho(f"⚠  {w}", fg=typer.colors.YELLOW, bold=True)
    typer.echo(f"\n✅ Done. Open: {out_dir}")


@app.command()
def backup(
    setup: bool = typer.Option(False, "--setup", help="One time: password into the Keychain, create the repository."),
    repo: str | None = typer.Option(None, "--repo", help="With --setup: repository folder (default: Google Drive)."),
    schedule: bool = typer.Option(False, "--schedule", help="Run the backup every evening (macOS LaunchAgent)."),
    unschedule: bool = typer.Option(False, "--unschedule", help="Remove the evening schedule."),
    test_restore: bool = typer.Option(False, "--test-restore", help="Restore the latest version into a temp folder and compare."),
) -> None:
    """Encrypted, versioned backup of the data (restic): companies, «Новые сделки», Data/.

    Without options: back up now, keep 7 daily + 4 weekly + 12 monthly versions, check
    the repository. The code folder is not included — it is on GitHub.
    """
    from dca import backup as bk

    root = _data_root()
    try:
        if setup:
            _backup_setup(root, Path(repo).expanduser() if repo else None)
        elif schedule:
            plist = bk.install_schedule(code_root() / "dca", root)
            typer.echo(f"✅ Бэкап по расписанию: каждый день в {bk.SCHEDULE_HOUR}:00 ({plist}).")
            typer.echo("   Если Mac выключен или спит, бэкап выполнится после пробуждения.")
            typer.echo(f"   Журнал: {root.data_dir / bk.LOG_FILE}")
        elif unschedule:
            typer.echo("✅ Расписание снято." if bk.remove_schedule() else "· Расписания не было.")
        elif test_restore:
            typer.echo("🔁 Пробное восстановление последней версии во временную папку...")
            report = bk.restore_check(root)
            typer.echo(f"  проверено файлов: {report['checked']}")
            for kind, label in (("missing", "нет в копии"), ("different", "отличается"), ("unreadable", "не открывается")):
                for item in report[kind]:
                    typer.secho(f"  ❌ {label}: {item}", fg=typer.colors.RED)
            if not report["ok"]:
                raise typer.Exit(code=1)
            typer.echo("✅ Копия восстанавливается и совпадает с данными.")
        else:
            excel = bk.open_in_excel(root)
            typer.echo(f"💾 Бэкап {root.path} → {bk.configured_repo(root) or '(не настроен)'}...")
            status = bk.run_backup(root)
            for name in excel:
                typer.secho(f"  ⚠ открыт в Excel, в копию попало последнее СОХРАНЁННОЕ состояние: {name}",
                            fg=typer.colors.YELLOW)
            size = (status.get("bytes") or 0) / 1024 / 1024
            typer.echo(f"✅ Бэкап готов: версия {status.get('snapshot')}, файлов {status.get('files')}, "
                       f"{size:.0f} МБ. Хранятся 7 дневных, 4 недельных и 12 месячных версий.")
    except bk.BackupError as e:
        typer.secho(f"❌ {e}", fg=typer.colors.RED, err=True)
        raise typer.Exit(code=1) from e


def _backup_setup(root: DataRoot, repo: Path | None) -> None:
    from dca import backup as bk

    if bk.find_restic() is None:
        raise bk.BackupError("restic не установлен — выполните в Терминале: brew install restic, затем повторите")
    if repo is None:
        drives = bk.google_drive_dirs()
        if not drives:
            raise bk.BackupError("папка Google Drive не найдена — запустите Google Drive для компьютера "
                                 "или укажите папку: dca backup --setup --repo <папка>")
        repo = drives[0] / bk.REPO_DIR_NAME
    repo = repo.resolve()
    if repo == root.path or root.path in repo.parents:
        raise bk.BackupError("репозиторий не может лежать внутри папки с данными — выберите другую папку")
    typer.echo("🔑 Придумайте пароль для бэкапов. macOS попросит ввести его дважды и сохранит в Связке ключей.")
    typer.echo("   ЗАПИШИТЕ его вне этого Mac (бумага, телефон): без пароля бэкап не открыть.")
    if not repo.joinpath("config").exists():
        bk.store_password()
    bk.init_repo(repo)
    marker = bk.save_repo(root, repo)
    typer.echo(f"✅ Репозиторий: {repo}")
    typer.echo(f"   Записано в {marker.name}. Дальше: dca backup, затем dca backup --schedule.")


@app.command()
def restore(
    snapshot: str | None = typer.Argument(None, help="Version id from --list (or «latest»)."),
    to: str | None = typer.Option(None, "--to", help="NEW empty folder outside the data folder."),
    list_versions: bool = typer.Option(False, "--list", help="List the backup versions."),
) -> None:
    """List backup versions, or restore one into a new folder (never over the working data)."""
    from dca import backup as bk

    root = _data_root()
    try:
        if list_versions or snapshot is None:
            for s in bk.snapshots(root):
                typer.echo(f"  {s.get('short_id') or s.get('id', '')[:8]}  {str(s.get('time', ''))[:16]}")
            return
        if not to:
            raise bk.BackupError("укажите новую папку: dca restore <версия> --to <папка>")
        restored = bk.restore(root, snapshot, Path(to).expanduser())
    except bk.BackupError as e:
        typer.secho(f"❌ {e}", fg=typer.colors.RED, err=True)
        raise typer.Exit(code=1) from e
    typer.echo(f"✅ Восстановлено в {restored}")
    typer.echo("   Сравните с рабочей папкой и перенесите нужные файлы вручную.")


@app.command(name="check-ledger")
def check_ledger_cmd(
    dry_run: bool = typer.Option(False, "--dry-run", help="Only check; do not rebuild «Баланс»."),
) -> None:
    """Check the batch ledger (rows typed by hand too) and rebuild its «Баланс».

    Checks what «Баланс» counts: deal_agreement_num, use_kg and the batch quantity
    as numbers, batch keys against the master database, double write-offs and
    over-use. Errors → nothing is written (exit 1). Otherwise «Баланс» is rebuilt
    for every batch — without waiting for the next `dca generate`.
    """
    from dca.excel_io import BATCHES_SHEET
    from dca.ledger import check_ledger, rebuild_ledger
    from dca.master_db import read_archives

    root = _data_root()
    path = root.ledger_path
    if not path.exists():
        typer.echo(f"❌ Ledger not found: {path}", err=True)
        raise typer.Exit(code=1)

    keys = None
    if root.master_db_path.exists():
        try:
            keys = set(read_archives(root.master_db_path, strict=False).keys(BATCHES_SHEET))
        except Exception as e:  # the check still runs, just without the key comparison
            typer.echo(f"  ⚠ master database not read ({e}) — batch keys are not compared with it")

    typer.echo(f"📒 Checking {path}...")
    try:
        report = check_ledger(path, keys)
    except Exception as e:
        raise _exit_with_error(e, "Cannot read the ledger") from e

    if report.balance:
        typer.echo("\nПартия             всего, кг   списано, кг   остаток, кг   сделки")
        for b in report.balance:
            typer.echo(f"  {b['key']:<18} {math.ceil(b['capacity']):>9} {math.ceil(b['used']):>13} "
                       f"{math.ceil(b['remaining']):>13}   {', '.join(b['deals'])}")
    for w in report.warnings:
        typer.secho(f"  ⚠ {w}", fg=typer.colors.YELLOW)
    if report.errors:
        typer.secho(f"\n❌ {len(report.errors)} ошибк(и) — реестр не изменён:", fg=typer.colors.RED, bold=True)
        for err in report.errors:
            typer.secho(f"  • {err}", fg=typer.colors.RED)
        typer.echo("Исправьте строки в Excel, сохраните, закройте файл и запустите команду снова.")
        raise typer.Exit(code=1)
    if dry_run:
        typer.echo("\n✅ Ошибок нет. --dry-run: «Баланс» не пересобран.")
        return
    try:
        n, rebuild_warnings = rebuild_ledger(path, root.master_db_path)
    except Exception as e:
        raise _exit_with_error(e, "Cannot rebuild the ledger") from e
    for w in rebuild_warnings:
        typer.secho(f"  ⚠ {w}", fg=typer.colors.YELLOW)
    typer.echo(f"\n✅ «Баланс» пересобран: {n} парт(ий). Прежняя версия — в {path.parent.name}/_ledger_versions/.")


@app.command(name="init-db")
def init_db() -> None:
    """Create the empty master database <data folder>/Data/master_db.xlsx.

    Four tabs — Стороны А / Стороны Б / Изделия / Партии — in the same layout as the
    deal Excel (rows = fields, columns C… = records, no 10-record limit). Fill it in
    Excel; every `dca new` / `dca new-incoming` then copies all its records into
    the new deal, and `dca refresh-db` re-copies them into an existing one.
    """
    from dca.master_db import init_master

    path = _data_root().master_db_path
    try:
        init_master(path)
    except FileExistsError as e:
        typer.echo(f"❌ {e}", err=True)
        raise typer.Exit(code=1) from e
    typer.echo(f"✅ Created {path}")
    typer.echo("   Fill the four tabs (one column per record, key in row 2) and save.")


@app.command(name="upgrade-db")
def upgrade_db() -> None:
    """Add fields that the input template has but the master database lacks (rows at the bottom)."""
    from dca.master_db import upgrade_master

    path = _data_root().master_db_path
    if not path.exists():
        typer.echo(f"❌ Master database not found: {path}\n   Create it with: dca init-db", err=True)
        raise typer.Exit(code=1)
    if _excel_has_open(path):
        typer.echo(f"❌ {path.name} is open in Excel — close it and run again.", err=True)
        raise typer.Exit(code=1)
    added = upgrade_master(path)
    if not added:
        typer.echo("✅ The database already has every field.")
        return
    for sheet, fields in added.items():
        typer.echo(f"  + «{sheet}»: {', '.join(fields)}")
    typer.echo(f"✅ Updated {path}")


@app.command(name="refresh-db")
def refresh_db(deal_dir: str) -> None:
    """Re-copy the archive sheets (Стороны А / Стороны Б / Изделия / Партии) of a deal
    from the master database.

    Only those four sheets change; «Сделка», «Передача», «Приложения» and the keys
    you selected stay as they are. Refuses if a selected key is no longer in the
    database. The previous Excel is kept in _archive/input_before_refresh_<time>.xlsx,
    and a summary shows which records were added, removed or changed.
    """
    from datetime import datetime

    from dca.excel_io import read_excel
    from dca.master_db import ARCHIVES, BATCH_KEY_FIELD, diff, read_archives
    from dca.xlsx_patch import fill_archive_sheets

    deal_dir_path = Path(deal_dir).resolve()
    _open_deal(deal_dir_path, None)
    input_xlsx = _find_input_file(deal_dir_path)
    if input_xlsx is None:
        typer.echo(f"❌ No input_*.xlsx found in {deal_dir_path}", err=True)
        raise typer.Exit(code=1)
    if _excel_has_open(input_xlsx):
        typer.echo(f"❌ {input_xlsx.name} is open in Excel — close it and run again.", err=True)
        raise typer.Exit(code=1)
    db_path = _data_root(deal_dir_path).master_db_path
    if not db_path.exists():
        typer.echo(f"❌ Master database not found: {db_path}\n   Create it with: dca init-db", err=True)
        raise typer.Exit(code=1)

    try:
        db = read_archives(db_path, strict=True)
        current = read_archives(input_xlsx, strict=False)
        selected = read_excel(input_xlsx)
    except Exception as e:
        raise _exit_with_error(e, "Cannot read") from e

    key_sheets = {
        "party_a_key": "Стороны А", "party_b_key": "Стороны Б", "item_key": "Изделия",
        **{f"{BATCH_KEY_FIELD}_{n}": "Партии" for n in range(1, 6)},
    }
    lost = [
        f"{field} = '{selected[field]}' (no such record in «{sheet}» of the database)"
        for field, sheet in key_sheets.items()
        if selected.get(field) not in (None, "") and str(selected[field]) not in db.keys(sheet)
    ]
    if lost:
        typer.echo("❌ Not refreshed — the deal uses records that are missing in the database:", err=True)
        for line in lost:
            typer.echo(f"   • {line}", err=True)
        raise typer.Exit(code=1)

    changes = diff(current, db)
    if not any(v for c in changes.values() for v in c.values()):
        typer.echo("✅ Already up to date — nothing changed.")
        return

    backup = deal_dir_path / "_archive" / f"input_before_refresh_{datetime.now():%Y%m%d-%H%M%S}.xlsx"
    backup.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(input_xlsx, backup)
    fill_archive_sheets(input_xlsx, db.records)

    for sheet in ARCHIVES:
        c = changes[sheet]
        if not any(c.values()):
            continue
        typer.echo(f"📇 {sheet}:")
        for label, sign in (("added", "+"), ("removed", "−"), ("changed", "~")):
            for item in c[label]:
                typer.echo(f"   {sign} {item}")
    for w in db.warnings:
        typer.echo(f"  ⚠ {w}")
    typer.echo(f"✅ Refreshed {input_xlsx.name} (previous version: _archive/{backup.name})")
    if any((deal_dir_path / d).is_dir() for d in (p.name for p in deal_dir_path.glob("output_*"))):
        typer.echo("   Documents in output_* still have the OLD data — run generate again to update them.")


@app.command(name="from-log")
def from_log(log_path: str) -> None:
    """Re-render documents from a previously saved log.json.

    Makes a NEW version of the package from the saved data with the CURRENT
    templates (the previous output goes to _archive/). Templates changed since the
    snapshot are listed in ❗ПРОВЕРЬТЕ; a snapshot from a newer DCA is refused.
    """
    log_path_obj = Path(log_path).resolve()
    if not log_path_obj.exists():
        typer.echo(f"❌ log.json not found: {log_path_obj}", err=True)
        raise typer.Exit(code=1)

    from dca.log import changed_templates, load_log_with_metadata

    typer.echo(f"📖 Loading {log_path_obj}...")
    try:
        deal, metadata = load_log_with_metadata(log_path_obj)
    except Exception as e:
        raise _exit_with_error(e, "Failed to load log") from e
    notices = []
    changed = changed_templates(metadata)
    if changed:
        # from-log = a NEW version from the old data with TODAY's templates.
        notices.append(f"Шаблоны изменились после сохранения log.json: {', '.join(changed)} — документы "
                       f"собраны ТЕКУЩИМИ шаблонами; прежние документы — в _archive/.")

    # Deal dir is parent-of-parent of log (output_<num>/log.json → deal_dir)
    deal_dir = (
        log_path_obj.parent.parent
        if log_path_obj.parent.name.startswith("output")
        else log_path_obj.parent
    )
    # Same entry as generate: lock, kind, duplicate number.
    identity = _open_deal(deal_dir, deal.deal_kind)
    identity = _confirm_identity(deal, deal_dir, identity)
    notices = _run_pipeline(deal, deal_dir, _data_root(deal_dir), notices, deal_id=identity.deal_id)
    typer.echo(f"\n✅ Done. Open: {deal_dir / OUTPUT_DIR.format(num=slug(deal.agreement_num))}")
    _show_notices(notices)


@app.command(name="rewrite-runners")
def rewrite_runners(
    deal_dir: str = typer.Argument(
        None,
        help="Deal folder to update. Omit when using --all.",
    ),
    all_deals: bool = typer.Option(
        False,
        "--all",
        help="Update runners in every deal folder (Новые сделки, <company>/Outgoing, <company>/Incoming).",
    ),
) -> None:
    """Rewrite the deal's *.command runners to the current format.

    Use this once after an upgrade that changes how runners are generated —
    e.g. to migrate older runners that contain a hardcoded `cd <project folder>`
    project path, or that rely only on `dca` being on PATH, into the current
    form (finds the code by walking up from the deal folder). Incoming deals keep
    their `generate-incoming` command.

    Examples:
        dca rewrite-runners deals/1-2026-1705_1200kg
        dca rewrite-runners --all
    """
    if all_deals:
        if deal_dir is not None:
            typer.echo("❌ Pass either <deal_dir> OR --all, not both.", err=True)
            raise typer.Exit(code=1)
        root = _data_root()
        deals = root.deal_dirs()
        if not deals:
            typer.echo(
                f"❌ No deal folders under {root.path} (data folder from {root.source}).", err=True
            )
            raise typer.Exit(code=1)
        n_changed = 0
        n_skipped = 0
        for sub in deals:
            ok, msg = _rewrite_runners_in_dir(sub)
            label = sub.relative_to(root.path).as_posix()
            if ok:
                typer.echo(f"  ✓ {label}: {msg}")
                n_changed += 1
            else:
                typer.echo(f"  ⊘ {label}: {msg}")
                n_skipped += 1
        typer.echo(f"\n✅ Done. Updated {n_changed} folder(s), skipped {n_skipped}.")
        return

    if not deal_dir:
        typer.echo("❌ Pass <deal_dir> or use --all.", err=True)
        raise typer.Exit(code=1)
    path = Path(deal_dir).resolve()
    ok, msg = _rewrite_runners_in_dir(path)
    if ok:
        typer.echo(f"✅ {path.name}: {msg}")
    else:
        typer.echo(f"❌ {path.name}: {msg}", err=True)
        raise typer.Exit(code=1)


@app.command()
def doctor() -> None:
    """Check where DCA finds its code, templates and data; print exact fixes.

    Read-only. Exit code 1 if something blocks normal use.
    """
    from dca.config import AMENDMENT_TEMPLATE, TEMPLATE_FILES, TEMPLATES_DIR, TRANSFER_TEMPLATE

    problems = 0
    code = code_root()
    reinstall = f'"{sys.executable}" -m pip install -e "{code}"'

    typer.echo("Code")
    typer.echo(f"  ✓ Python {sys.version.split()[0]}  ({sys.executable})")
    typer.echo(f"  ✓ code folder: {code}")
    templates = [*TEMPLATE_FILES.values(), AMENDMENT_TEMPLATE, TRANSFER_TEMPLATE]
    missing = sorted({t for t in templates if not (TEMPLATES_DIR / t).exists()})
    if not EMPTY_TEMPLATE.exists():
        missing.append(str(EMPTY_TEMPLATE.relative_to(code)))
    if missing:
        problems += 1
        typer.echo(f"  ❌ missing in {code}: {', '.join(missing)}")
    else:
        typer.echo(f"  ✓ templates + {EMPTY_TEMPLATE.name} present")

    typer.echo("`dca` command")
    # `./dca` runs the code beside it with no install: a missing or
    # stale pip pointer only breaks `dca` on PATH — a warning, not a blocking problem.
    launcher = code / "dca"
    typer.echo(f"  {'✓' if launcher.exists() else '⚠'} without install: {launcher}"
               f"{'' if launcher.exists() else ' — not found'}")
    pointers = _editable_pointers()
    if not pointers:
        typer.echo("  · dca is not pip-installed in this Python (runners and the launcher above still work).")
        typer.echo(f"    Optional, for `dca` on PATH:  {reinstall}")
    stale = False
    for pth, target in pointers:
        if not target.exists():
            stale = True
            typer.echo(f"  ⚠ {pth.name} points to a folder that no longer exists: {target}")
            typer.echo(f"    Use {launcher} instead, or fix `dca` on PATH:  {reinstall}")
        elif target.resolve() != (code / "src").resolve():
            typer.echo(f"  ⚠ {pth.name} points to another copy of DCA: {target}")
            typer.echo(f"    To use this one:  {reinstall}")
        else:
            typer.echo(f"  ✓ {pth.name} → {target}")
    on_path = shutil.which("dca")
    if on_path and stale:
        typer.echo(f"  ⚠ on PATH: {on_path} — fails with ModuleNotFoundError until the fix above")
    else:
        typer.echo(f"  {'✓' if on_path else '⚠'} on PATH: {on_path or 'not found'}")

    typer.echo("Data")
    try:
        root = _data_root()
    except ValueError as e:
        typer.echo(f"  ❌ {e}")
        raise typer.Exit(code=1) from e
    typer.echo(f"  ✓ data folder: {root.path}")
    typer.echo(f"    chosen by: {root.source}")
    for label, d in (
        ("Новые сделки/", root.new_deals_dir),
        ("Data/", root.data_dir),
        ("Data/scans/", root.scans_dir),
        ("ledger", root.ledger_path),
    ):
        typer.echo(f"  {'✓' if d.exists() else '·'} {label}: {d}{'' if d.exists() else '  (not yet created)'}")

    if root.master_db_path.exists():
        from dca.master_db import read_archives
        try:
            db = read_archives(root.master_db_path, strict=True)
            typer.echo(f"  ✓ master database: {root.master_db_path.name} ({db.count_line()})")
            for w in db.warnings:
                typer.echo(f"  ⚠ {w}  →  dca upgrade-db")
            # Data readiness (audit 2026-10-28 B02): an empty sheet is not an error, but
            # every new deal would start with empty dropdowns.
            empty = [sheet for sheet, recs in db.records.items() if not recs]
            if empty:
                typer.echo(f"  ⚠ data readiness: no records on {', '.join(empty)} — new deals get "
                           f"empty dropdowns there. Fill {root.master_db_path.name}, then dca refresh-db.")
            else:
                typer.echo("  ✓ data readiness: parties, items and batches present")
            from dca.attachments import _company_profile
            for rec in db.records["Стороны А"]:
                folder = rec.get("party_a_folder_name")
                if folder:
                    cert = _company_profile(root.company_dir(str(folder)))
                    mark = "✓" if cert else "⚠"
                    typer.echo(f"  {mark} company «{folder}»: company profile "
                               f"{cert.name if cert else 'NOT FOUND in ' + str(root.company_dir(str(folder)))}")
        except ValueError as e:
            problems += 1
            typer.echo(f"  ❌ master database has errors:\n{e}")
    else:
        typer.echo(f"  · master database: {root.master_db_path}  (not yet created — dca init-db)")

    typer.echo("Backup")
    from dca import backup as bk
    bk_repo = bk.configured_repo(root)
    if bk_repo is None:
        typer.echo("  ⚠ backups are not set up — dca backup --setup")
    else:
        status = bk.read_status(root)
        typer.echo(f"  {'✓' if bk.find_restic() else '⚠'} restic: {bk.find_restic() or 'not found — brew install restic'}")
        typer.echo(f"  · repository: {bk_repo}")
        notice = bk.stale_notice(root)
        last = status.get("last_success")
        typer.echo(f"  {'⚠ ' + notice if notice else '✓ last backup: ' + str(last)[:16]}")
        scheduled = bk.scheduled_launcher()
        if scheduled is None:
            typer.echo("  ⚠ no schedule — dca backup --schedule")
        elif not scheduled.exists():
            typer.echo(f"  ⚠ the schedule runs {scheduled}, which no longer exists — dca backup --schedule")
        else:
            typer.echo(f"  ✓ every day at {bk.SCHEDULE_HOUR}:00")
        test = status.get("last_restore_test")
        typer.echo(f"  {'✓' if test and test.get('ok') else '·'} test restore: "
                   f"{test['at'][:10] + (' OK' if test['ok'] else ' FAILED') if test else 'never — dca backup --test-restore'}")

    typer.echo("Runners")
    outdated = []
    for sub in root.deal_dirs():
        for runner_file in [*sub.glob("run_*.command"), *sub.glob("transfer_*.command"),
                            *sub.glob("amend_*.command")]:
            if RUNNER_TAG not in runner_file.read_text(encoding="utf-8", errors="replace"):
                outdated.append(sub.relative_to(root.path).as_posix())
                break
    if outdated:
        typer.echo(f"  ⚠ {len(outdated)} deal(s) with outdated runners: {', '.join(outdated)}")
        typer.echo("    Fix:  dca rewrite-runners --all")
    else:
        typer.echo("  ✓ all runners current")

    if problems:
        typer.echo(f"\n❌ {problems} problem(s) found.")
        raise typer.Exit(code=1)
    typer.echo("\n✅ OK")


# ─── Helpers ───────────────────────────────────────────────────────────────


def _editable_pointers() -> list[tuple[Path, Path]]:
    """(pth file, target) for every `pip install -e` pointer to dca in this Python."""
    import site

    dirs = [*site.getsitepackages(), site.getusersitepackages()]
    found = []
    for d in dirs:
        for pth in sorted(Path(d).glob("__editable__.dca-*.pth")):
            for line in pth.read_text(encoding="utf-8", errors="replace").splitlines():
                line = line.strip()
                if line and not line.startswith(("#", "import ")):
                    found.append((pth, Path(line)))
    return found

# Error types we expect during normal use — operator typo, bad xlsx, missing
# file, locked ledger. These get a one-line message, no traceback (the
# traceback would just be noise for a non-developer operator).
#   - ValidationError: pydantic field/cross-field validation
#   - ValueError: our own load_deal / load_log raises
#   - FileNotFoundError / PermissionError: filesystem
#   - zipfile.BadZipFile: corrupt xlsx (openpyxl raises this — xlsx is a zip)
_EXPECTED_ERRORS: tuple[type[BaseException], ...] = (
    ValidationError,
    ValueError,
    FileNotFoundError,
    PermissionError,
    zipfile.BadZipFile,
)


def _exit_with_error(e: Exception, context: str) -> typer.Exit:
    """Show an error to the operator and return a typer.Exit to raise.

    Expected errors → one short line + exit(1).
    Unexpected errors → full traceback + exit(2) so the operator can copy-paste
    it when reporting a bug (the bare message `e` often loses crucial info).
    """
    if isinstance(e, _EXPECTED_ERRORS):
        typer.echo(f"❌ {context}: {e}", err=True)
        return typer.Exit(code=1)
    typer.echo(f"❌ {context} — UNEXPECTED ERROR (please report):", err=True)
    typer.echo(traceback.format_exc(), err=True)
    return typer.Exit(code=2)


def _find_input_file(deal_dir: Path) -> Path | None:
    """Find the one `input_*.xlsx` in a deal folder. Returns None if not found."""
    if not deal_dir.is_dir():
        return None
    # Several candidates → refuse: picking the first by name would choose
    # input_BACKUP.xlsx over input_CURRENT.xlsx.
    matches = sorted(p for p in deal_dir.glob("input_*.xlsx") if not p.name.startswith("~$"))
    if len(matches) > 1:
        typer.echo(f"❌ Several input files in {deal_dir.name}: "
                   f"{', '.join(p.name for p in matches)} — keep exactly one input_*.xlsx "
                   f"(move the others into _archive/).", err=True)
        raise typer.Exit(code=1)
    return matches[0] if matches else None


def _run_pipeline(
    deal: Deal, deal_dir: Path, root: DataRoot, extra_notices: list[str] | None = None,
    input_xlsx: Path | None = None, deal_id: str | None = None, old_agreement_num: str | None = None,
) -> list[str]:
    """Core pipeline: build in a temp folder → publish → ledger → check file.

    Everything is built in `_сборка_output_<num>_<id>/` (a visible name: Finder
    remembers a dot folder as hidden); only a complete set replaces `output_<num>/`
    (the old one goes to `_archive/vN/`). A failure while building leaves the
    previous output untouched. The ledger is updated after the documents are published.

    Returns the soft notices; they are written to output_<num>/❗ПРОВЕРЬТЕ.txt and
    the caller shows them last (_show_notices).
    """
    import uuid
    from datetime import date as _date

    from dca.checks import (
        attachment_warnings,
        batch_company_warnings,
        date_warnings,
        write_check_file,
    )
    from dca.config import SIGNED_DIR

    notices: list[str] = list(extra_notices or [])
    num_slug = slug(deal.agreement_num)
    output_dir_name = OUTPUT_DIR.format(num=num_slug)
    output_dir = deal_dir / output_dir_name
    if deal.letter_date is None:
        # Resolve the date once and keep it: a re-generation (any later day) prints
        # the date of the previous generation, not «today».
        deal.letter_date = _previous_letter_date(output_dir / OUTPUT_FILES["log"]) or _date.today()

    # Operator-managed folder for signed PDFs: created empty if missing, never touched.
    signed_dir = deal_dir / SIGNED_DIR.format(num=num_slug)
    signed_version = None
    if not signed_dir.exists():
        signed_dir.mkdir(parents=True)
        typer.echo(f"📁 Created empty {signed_dir.name}/ for signed PDFs")
    else:
        n = sum(1 for p in signed_dir.iterdir() if not p.name.startswith("."))
        typer.echo(f"📁 {signed_dir.name}/ exists ({n} item(s)) — left untouched")
        signed_version = signatures.bind(signed_dir, output_dir)

    # Leftovers of failed builds go first — the deal lock rules out a parallel build.
    _remove_build_leftovers(deal_dir)
    building = deal_dir / BUILD_DIR.format(output=output_dir_name, id=uuid.uuid4().hex[:8])
    building.mkdir()
    try:
        mode_note = f"deal_kind={deal.deal_kind}" + (f", variant={deal.variant}" if deal.variant else "")
        n_annexes = len(deal.resolved_annexes())
        annex_note = f", {n_annexes} annexes" if n_annexes > 1 else ""
        typer.echo(f"🖨  Rendering documents ({mode_note}{annex_note})...")
        rendered = render_all(deal)
        typer.echo(f"   {len(rendered)} file(s):")
        for filename, content in rendered.items():
            (building / filename).write_bytes(content)
            typer.echo(f"  ✓ {filename}  ({len(content) // 1024} KB)")

        # Attachment pack + batch ledger are OUTGOING-only.
        if deal.deal_kind == "outgoing":
            company = root.company_dir(deal.party_a_folder_name) if deal.party_a_folder_name else None
            notices += _copy_attachments(deal, building, root.scans_dir, company)
            notices += attachment_warnings(deal)
            notices += batch_company_warnings(deal)
        notices += date_warnings(deal)   # the agreement's term: both deal kinds

        written = [p for p in building.iterdir() if p.suffix == ".docx"]
        if len(written) != len(rendered):
            # Two names that the file system treats as one.
            raise ValueError(f"{len(rendered)} documents rendered but {len(written)} files written — "
                             f"two document names collide on this disk; rename the annexes.")

        typer.echo("💾 Writing log.json...")
        dump_log(deal, building / OUTPUT_FILES["log"], input_xlsx=input_xlsx, deal_id=deal_id)
        typer.echo(f"  ✓ {OUTPUT_FILES['log']}")

        typer.echo("🗜  Building ZIP package...")
        pkg_path = building / OUTPUT_FILES["package"].format(num=num_slug)
        _zip_package(building, pkg_path)
        typer.echo(f"  ✓ {pkg_path.name}  ({pkg_path.stat().st_size // 1024} KB)")

        # The transfer act is re-rendered from the same data (not part of the ZIP).
        # Its data comes late in a deal: an unfilled «Передача» is normal and is NOT a
        # notice — notices are kept for real problems (no banner blindness).
        if deal.deal_kind == "outgoing" and deal.transfer is not None and deal.transfer.missing():
            typer.echo("📝 Transfer act: «Передача» not filled yet — skipped "
                       f"(later: transfer_{num_slug}.command)")
        elif deal.deal_kind == "outgoing" and deal.transfer is not None:
            typer.echo("📝 Rendering the transfer act («Передача»)...")
            try:
                _write_transfer_docs(render_transfer(deal), building / TRANSFER_DIR.format(num=num_slug))
            except ValueError as e:
                typer.echo(f"  ⚠ not rendered — {e}\n    Fill «Передача», then: dca transfer <deal folder>")
                notices.append(f"Акт передачи не создан: {e}")

        if signed_version is not None:
            stale = signatures.stale_documents(signed_version, building, signed_dir)
            if stale:
                notices.append(
                    f"В {signed_dir.name}/ лежат подписи к ПРЕДЫДУЩЕЙ версии документов — изменились: "
                    f"{', '.join(Path(s).stem for s in stale)}. Переподпишите и замените файлы в "
                    f"{signed_dir.name}/ (до замены предупреждение будет повторяться).")
    except BaseException:
        shutil.rmtree(building, ignore_errors=True)
        typer.echo(f"  ✗ build failed — previous {output_dir_name}/ left as it was", err=True)
        raise

    typer.echo("📦 Publishing (previous output → _archive/)...")
    _publish(building, deal_dir, output_dir_name)
    if old_agreement_num:
        notices += _retire_old_number(deal_dir, old_agreement_num, deal.agreement_num)

    try:
        if deal.deal_kind == "outgoing":
            notices += _update_ledger(deal, root, deal_id)
    finally:
        write_check_file(output_dir, notices)   # not part of the package ZIP
    return notices


def _previous_letter_date(prev_log: Path) -> _dt.date | None:
    """letter_date of the previous generation (its log.json), if any."""
    import json

    try:
        value = json.loads(prev_log.read_text(encoding="utf-8")).get("deal", {}).get("letter_date")
        return _dt.date.fromisoformat(value) if value else None
    except (OSError, ValueError, AttributeError, TypeError):
        return None


BUILD_DIR = "_сборка_{output}_{id}"            # visible on purpose (no leading dot)
BUILD_GLOBS = ("_сборка_*",)


def _remove_build_leftovers(deal_dir: Path) -> None:
    """Delete build folders that a crashed run left in the deal folder."""
    for pattern in BUILD_GLOBS:
        for p in deal_dir.glob(pattern):
            if p.is_dir():
                shutil.rmtree(p, ignore_errors=True)
                typer.echo(f"🧹 удалена служебная папка прерванной сборки: {p.name}")


def _make_visible(output_dir: Path, deal_dir: Path) -> None:
    """Clear the «hidden» flag on output_<num> and touch the deal folder so Finder re-reads it.

    Finder may keep a folder hidden that was ever built under a dot name.
    Cosmetic: never stops the command; file flags exist on macOS / BSD only.
    """
    import os
    import stat

    try:
        hidden = getattr(stat, "UF_HIDDEN", 0)
        flags = getattr(os.stat(output_dir), "st_flags", 0)
        if hidden and flags & hidden and hasattr(os, "chflags"):
            os.chflags(output_dir, flags & ~hidden)
        os.utime(deal_dir)
    except OSError as e:
        typer.echo(f"  ⚠ не удалось обновить вид папки {output_dir.name} в Finder: {e}", err=True)


def _publish(building: Path, deal_dir: Path, output_dir_name: str) -> None:
    """Swap the new output in; if that fails, put the previous output back.

    Then make output_<num> visible in Finder."""
    from dca.archive import archive_previous_output

    output_dir = deal_dir / output_dir_name
    previous = archive_previous_output(deal_dir, output_dir_name)
    try:
        building.rename(output_dir)
    except BaseException:
        if previous is not None and not output_dir.exists():
            shutil.move(str(previous), str(output_dir))
            if not any(previous.parent.iterdir()):
                previous.parent.rmdir()
        shutil.rmtree(building, ignore_errors=True)
        typer.echo(f"  ✗ publishing failed — previous {output_dir_name}/ restored", err=True)
        raise
    _make_visible(output_dir, deal_dir)


def _zip_package(output_dir: Path, pkg_path: Path) -> None:
    """The package for handing over: the documents and the attachment scans ONLY.

    Never log.json (internal data), ❗ПРОВЕРЬТЕ.txt or transfer_docs.
    """
    import zipfile

    members = sorted(p for p in output_dir.glob("*.docx") if not p.name.startswith("~$"))
    members += sorted(p for d in output_dir.glob("attachments_*") if d.is_dir()
                      for p in d.rglob("*") if p.is_file() and not p.name.startswith((".", "~$")))
    with zipfile.ZipFile(pkg_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for p in members:
            zf.write(p, p.relative_to(output_dir).as_posix())


def _file_and_finish(deal: Deal, deal_dir_path: Path, root: DataRoot, input_xlsx: Path,
                     excel_open: bool, notices: list[str]) -> None:
    """File the deal into its company folder, then say where the result is.

    The printed path is the final one (after the move), and «not moved» is a red
    block at the very end, never one line among the notices.
    """
    if excel_open:
        final, not_moved = deal_dir_path, (
            f"{input_xlsx.name} был открыт в Excel — в документы попало последнее СОХРАНЁННОЕ состояние",
            "сохраните файл, закройте Excel полностью (⌘Q), затем снова нажмите run (или dca generate)")
    else:
        final, not_moved = _place_deal(deal, deal_dir_path, root)
    out = final / OUTPUT_DIR.format(num=slug(deal.agreement_num))
    # Check the result where we say it is: never «Done» without documents at that path.
    n_docx = sum(1 for p in out.glob("*.docx") if not p.name.startswith("~$")) if out.is_dir() else 0
    _write_generate_log(root, deal, deal_dir_path, final, out, n_docx)
    if n_docx == 0:
        typer.secho(
            f"\n{'═' * 72}\n❌ ОШИБКА: папка с документами не найдена или пуста: {out}\n"
            f"   Ничего не удаляйте. Пришлите последнюю строку из "
            f"{root.data_dir / GENERATE_LOG_FILENAME} — по ней видно, куда ушла сделка.\n{'═' * 72}",
            fg=typer.colors.RED, bold=True)
        raise typer.Exit(code=1)
    if not_moved:
        typer.secho(f"\n⚠  Документы готовы: {out}", fg=typer.colors.YELLOW, bold=True)
    else:
        typer.echo(f"\n✅ Done. Open: {out}")
    _open_in_finder(out)
    rel = final.relative_to(root.path).as_posix() if _is_under(final, root.path) else str(final)
    # Not a deal problem, so not in ❗ПРОВЕРЬТЕ.txt; shown only when backups need attention.
    from dca.backup import stale_notice
    backup_notice = stale_notice(root)
    if backup_notice:
        typer.secho(f"\n⚠  {backup_notice}", fg=typer.colors.YELLOW, bold=True)
    _show_notices(notices, not_moved=(rel, *not_moved) if not_moved else None)


GENERATE_LOG_FILENAME = "_generate.log"


def _write_generate_log(root: DataRoot, deal: Deal, before: Path, final: Path, out: Path, n_docx: int) -> None:
    """One line per generate run in Data/_generate.log: where the deal and its output ended up.

    Never stops the command.
    """
    line = (f"{_dt.datetime.now():%Y-%m-%d %H:%M:%S} agreement={deal.agreement_num} kind={deal.deal_kind} "
            f"root={root.path} from={before} to={final} output={out} "
            f"output_exists={'yes' if out.is_dir() else 'no'} docx={n_docx}\n")
    try:
        root.data_dir.mkdir(parents=True, exist_ok=True)
        with open(root.data_dir / GENERATE_LOG_FILENAME, "a", encoding="utf-8") as f:
            f.write(line)
    except OSError as e:
        typer.echo(f"  ⚠ журнал {GENERATE_LOG_FILENAME} не записан: {e}", err=True)


def _open_in_finder(path: Path) -> None:
    """Show the result folder in Finder — only for a run from a Terminal on macOS.

    The operator no longer looks for the printed path by hand. Never under tests or pipes
    (stdout is not a terminal) and never with DCA_NO_OPEN set.
    """
    import os

    if sys.platform != "darwin" or os.environ.get("DCA_NO_OPEN") or not sys.stdout.isatty():
        return
    try:
        subprocess.run(["open", str(path)], check=False, timeout=10)
    except (OSError, subprocess.SubprocessError):
        pass


def _retire_old_number(deal_dir: Path, old_num: str, new_num: str) -> list[str]:
    """The agreement number changed: clear the old number out of the deal folder.

    - output_<old>/ goes to _archive/vN/ — old documents never lie next to the new ones;
    - the runners get the new number in their names;
    - an empty signed_<old>/ is removed; one with files is left as it is, with a notice.
    input_<old>.xlsx keeps its name: Excel may have it open. → the notices.
    """
    from dca.archive import archive_previous_output
    from dca.config import SIGNED_DIR

    old_slug, new_slug = slug(old_num), slug(new_num)
    if old_slug == new_slug:
        return []
    notices: list[str] = []
    archived = archive_previous_output(deal_dir, OUTPUT_DIR.format(num=old_slug))
    if archived is not None:
        typer.echo(f"  🗄 {OUTPUT_DIR.format(num=old_slug)}/ (прежний номер) → {archived.relative_to(deal_dir)}/")
    for name in (RUN_COMMAND_FILENAME, TRANSFER_RUN_COMMAND_FILENAME, AMENDMENT_RUN_COMMAND_FILENAME):
        src, dst = deal_dir / name.format(num=old_slug), deal_dir / name.format(num=new_slug)
        if src.exists() and not dst.exists():
            src.rename(dst)
            typer.echo(f"  ✎ {src.name} → {dst.name}")
    signed = deal_dir / SIGNED_DIR.format(num=old_slug)
    if signed.is_dir():
        entries = list(signed.iterdir())
        if all(p.name == ".DS_Store" for p in entries):     # empty (Finder may leave .DS_Store)
            for p in entries:
                p.unlink()
            signed.rmdir()
        else:
            notices.append(
                f"В {signed.name}/ лежат подписи к прежнему номеру соглашения {old_num}. Программа их не "
                f"трогает: перенесите нужные файлы в {SIGNED_DIR.format(num=new_slug)}/ или удалите папку сами.")
    return notices


def _show_notices(notices: list[str], not_moved: tuple[str, str, str] | None = None) -> None:
    """A yellow block at the very end of the output — the last thing the operator sees.

    `not_moved` = (where the deal stayed, why, what to do): a red block after it.
    """
    from dca.checks import CHECK_FILENAME

    bar = "═" * 72
    if notices:
        typer.secho(f"\n{bar}\n⚠  ПРОВЕРЬТЕ ПЕРЕД ОТПРАВКОЙ  (то же — в {CHECK_FILENAME} в папке output)",
                    fg=typer.colors.YELLOW, bold=True)
        for n in notices:
            typer.secho(f"  • {n}", fg=typer.colors.YELLOW, bold=True)
        typer.secho(bar, fg=typer.colors.YELLOW, bold=True)
    if not_moved:
        where, why, todo = not_moved
        typer.secho(f"\n{bar}\n⛔ СДЕЛКА НЕ ПЕРЕНЕСЕНА В ПАПКУ КОМПАНИИ — осталась здесь: {where}\n"
                    f"   Причина: {why}.\n   Что сделать: {todo}.\n{bar}", fg=typer.colors.RED, bold=True)
    if any(n.startswith("РЕЕСТР НЕ ОБНОВЛЁН") for n in notices):
        typer.secho("❌ Документы созданы, но реестр партий НЕ обновлён — сделка не завершена.",
                    fg=typer.colors.RED, bold=True)
        raise typer.Exit(code=1)


def _snapshot_master_db(input_xlsx: Path) -> None:
    """Copy every record of the master database into a new deal's archive sheets."""
    from dca.master_db import read_archives
    from dca.xlsx_patch import fill_archive_sheets

    db_path = _data_root(input_xlsx.parent).master_db_path
    if not db_path.exists():
        typer.echo(f"  ⚠ no master database ({db_path}) — archive sheets left empty. "
                   f"Create one with: dca init-db")
        return
    try:
        db = read_archives(db_path, strict=True)
    except ValueError as e:
        typer.echo(f"  ⚠ master database not copied — fix it first:\n{e}", err=True)
        return
    fill_archive_sheets(input_xlsx, db.records)
    typer.echo(f"  📇 archives from the database: {db.count_line()}")
    for w in db.warnings:
        typer.echo(f"  ⚠ {w}")


def _write_transfer_docs(rendered: dict[str, bytes], out_dir: Path) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)
    for filename, content in rendered.items():
        (out_dir / filename).write_bytes(content)
        typer.echo(f"  ✓ {out_dir.name}/{filename}  ({len(content) // 1024} KB)")


def _copy_attachments(deal: Deal, output_dir: Path, scans_dir: Path, company_dir: Path | None) -> list[str]:
    """Copy & rename the attachment scans from Data/scans into output/attachments_<num>/.

    Searches recursively in Data/scans by file name. The operator fills the
    batch_agreement_file / batch_annex_file rows of the «Партии» sheet; the files
    are renamed to "N. label.ext" matching the numbering of the cover list.
    Outgoing deals only.
    """
    from dca.attachments import copy_attachment_files

    if not scans_dir.exists() and company_dir is None:
        return []
    typer.echo("📁 Copying attachment scans (Data/scans → attachments_*)...")
    manifest = copy_attachment_files(deal, output_dir, scans_dir, company_dir)
    for src, tgt in manifest["copied"]:
        typer.echo(f"  ✓ {src}  →  {tgt}")
    for m in manifest["missing"]:
        typer.echo(f"  ⚠ MISSING in Data/scans: {m}")
    for s in manifest["skipped"]:
        typer.echo(f"  ⊘ #{s['n']} [{s['kind']}] batch #{s['pkg']}: no file name in Excel")
    for a in manifest.get("ambiguous", []):
        typer.echo(f"  ⚠ AMBIGUOUS '{a['source_name']}' — different files with this name, NOT copied:")
        for p in a["all_paths"]:
            typer.echo(f"      • {p}")
    for f in manifest.get("missing_folders", []):
        typer.echo(f"  ⚠ scans folder not found in Data/scans: {f}")
    notices = [f"Папка со сканами не найдена в Data/scans ({f}) — сканы искались по всей папке scans"
               for f in manifest.get("missing_folders", [])]
    notices += [f"Нет файла для пакета приложений: {m}" for m in manifest["missing"]]
    notices += [f"Пакет приложений, пункт {s['n']}: в «Партиях» не указано имя файла скана "
                f"(партия #{s['pkg']}, {s['kind']})" for s in manifest["skipped"]]
    notices += [f"Скан «{a['source_name']}» лежит в нескольких местах, и файлы РАЗНЫЕ — НЕ скопирован. "
                f"Укажите у партии «Папку со сканами» в базе (или оставьте один вариант): "
                + "; ".join(str(p) for p in a["all_paths"]) for a in manifest.get("ambiguous", [])]
    return notices


def _update_ledger(deal: Deal, root: DataRoot, deal_id: str | None = None) -> list[str]:
    """Update the batch ledger (audit trail + balance per batch).

    See `dca.ledger`. Outgoing deals only.
    """
    from dca.ledger import append_to_ledger

    typer.echo("📒 Updating batch ledger...")
    try:
        ledger_result = append_to_ledger(deal, root.ledger_path, deal_id=deal_id)
    except Exception as e:
        # Documents are made but the accounting is not: never report a clean
        # «Done» — the notice makes generate exit non-zero.
        # ANY failure, e.g. a damaged ledger (BadZipFile).
        known = isinstance(e, (PermissionError, ValueError, OSError))
        msg = f"РЕЕСТР НЕ ОБНОВЛЁН: {e}" if known else f"РЕЕСТР НЕ ОБНОВЛЁН: {type(e).__name__}: {e}"
        typer.echo(f"  ⚠ {msg}", err=True)
        return [msg]
    if ledger_result["skipped"]:
        typer.echo("  ⊘ no active batches — ledger not touched")
        return []
    for w in ledger_result["warnings"]:
        typer.echo(f"  ⚠ {w}")
    typer.echo(
        f"  ✓ {ledger_result['ledger_path'].name}: "
        f"-{len(ledger_result['removed'])} +{len(ledger_result['added'])} rows"
    )
    return [f"Реестр партий: {w}" for w in ledger_result["warnings"]]


def _company_folder(deal: Deal) -> str | None:
    """Our company's folder for this deal: party A's (outgoing) or party B's (incoming)."""
    return deal.party_b_folder_name if deal.deal_kind == "incoming" else deal.party_a_folder_name


def _require_company_folder(deal: Deal) -> None:
    """Stop BEFORE rendering if the chosen party isn't one of our companies,
    or its folder name can't be a company folder."""
    folder = _company_folder(deal)
    if folder:
        problem = _bad_company_folder_name(folder)
        if problem:
            typer.echo(f"❌ Company folder name «{folder}» {problem}. Fix it in the master "
                       f"database (then dca refresh-db).", err=True)
            raise typer.Exit(code=1)
        return
    if deal.deal_kind == "incoming":
        who, key, field = "«Стороны Б»", deal.party_b_key, "party_b_folder_name"
    else:
        who, key, field = "«Стороны А»", deal.party_a_key, "party_a_folder_name"
    typer.echo(
        f"❌ {who} record '{key}' has no {field} — it is not one of our companies, or the "
        f"field is empty.\n   Fill {field} in the master database (then dca refresh-db) or "
        f"pick our company in the deal's Excel.",
        err=True,
    )
    raise typer.Exit(code=1)


def _bad_company_folder_name(name: str) -> str | None:
    """Why `name` can't be a company folder (None = fine): it must be one plain folder
    next to «Новые сделки» and Data, not collide with them, not be hidden."""
    from dca.paths import DATA_DIR_NAME, INCOMING_DIR_NAME, NEW_DEALS_DIR_NAME, OUTGOING_DIR_NAME

    if name != name.strip():
        return "has leading/trailing spaces"
    if any(ch in name for ch in '/\\:*?"<>|'):
        return "contains a character not allowed in a folder name (/ \\ : * ? \" < > |)"
    if name.startswith((".", "_")):
        return "starts with «.» or «_» (such folders are hidden / reserved)"
    reserved = {DATA_DIR_NAME, NEW_DEALS_DIR_NAME, OUTGOING_DIR_NAME, INCOMING_DIR_NAME}
    if name.casefold() in {r.casefold() for r in reserved}:
        return "is reserved (Data / Новые сделки / Outgoing / Incoming)"
    return None


def _refuse_duplicate(agreement_num: str, except_dir: Path | None = None, deal_id: str | None = None) -> None:
    """An agreement number belongs to one deal across all companies: refuse a second deal
    with it anywhere in the data root — on `new`/`new-incoming`, `generate*` and `from-log`.

    Decided by what each folder holds (identity file → log.json → folder name), not by
    the folder name alone: a renamed folder kept its number. A folder
    with the same deal_id is a copy of this deal.
    """
    key = number_key(agreement_num)
    root = _data_root(except_dir)
    skip = except_dir.resolve() if except_dir else None
    for d in root.deal_dirs():
        if d.resolve() == skip:
            continue
        try:
            other = read_identity(d)
        except ValueError:
            other = None
        copy = bool(deal_id and other and other.deal_id == deal_id)
        num = known_number(d)
        if copy or (num is not None and number_key(num) == key):
            where = d.relative_to(root.path).as_posix()
            what = "A copy of this deal" if copy else f"A deal with agreement number {slug(agreement_num)}"
            typer.echo(f"❌ {what} already exists: {where}\n"
                       f"   Open that folder, or use another agreement number.", err=True)
            raise typer.Exit(code=1)


def _deal_kind(deal_dir: Path, identity: DealIdentity | None) -> str | None:
    """outgoing / incoming — fixed when the deal was created (identity file; without it:
    its runners or Outgoing/Incoming folder). None for a bare folder that says neither."""
    if identity is not None:
        return identity.kind
    runners = sorted(deal_dir.glob("run_*.command"))
    if not runners and deal_dir.parent.name not in ("Outgoing", "Incoming"):
        return None
    return "incoming" if _runner_command(deal_dir, runners) == "generate-incoming" else "outgoing"


def _open_deal(deal_dir: Path, kind: str | None) -> DealIdentity | None:
    """Common entry of every command that works on a deal folder:
    the deal lock first, then the deal's kind must match the command — the Excel
    never decides outgoing vs incoming."""
    _lock_deal(deal_dir)
    try:
        identity = read_identity(deal_dir)
    except ValueError as e:
        typer.echo(f"❌ {e}", err=True)
        raise typer.Exit(code=1) from e
    actual = _deal_kind(deal_dir, identity)
    if kind is not None and actual is not None and actual != kind:
        use = "generate-incoming" if actual == "incoming" else "generate"
        typer.echo(f"❌ This is an {actual} deal (its identity file / runner / folder says so). "
                   f"Use:  dca {use} {deal_dir}", err=True)
        raise typer.Exit(code=1)
    return identity


def _confirm_identity(deal: Deal, deal_dir: Path, identity: DealIdentity | None) -> DealIdentity:
    """After loading: refuse a duplicate number, then create / update the identity file."""
    _refuse_duplicate(deal.agreement_num, except_dir=deal_dir,
                      deal_id=identity.deal_id if identity else None)
    if identity is None:
        identity = new_identity(deal.agreement_num, deal.deal_kind)
    elif identity.agreement_num != deal.agreement_num:
        typer.echo(f"  ✎ agreement number changed: {identity.agreement_num} → {deal.agreement_num}")
        identity.agreement_num = deal.agreement_num
    else:
        return identity
    if deal_dir.is_dir():
        write_identity(deal_dir, identity)
    return identity


def _lock_deal(deal_dir: Path) -> None:
    """One dca command per deal at a time.

    A second generate / transfer / amendment / refresh-db on the same deal stops
    at once instead of interleaving writes. Released when the command ends.
    """
    import click
    from filelock import FileLock, Timeout

    if not deal_dir.is_dir():
        return
    lock = FileLock(str(deal_dir / ".dca.lock"), timeout=0)
    try:
        lock.acquire()
    except Timeout:
        typer.echo(f"❌ Another dca command is working on {deal_dir.name} right now — wait for it to "
                   f"finish and run again.", err=True)
        raise typer.Exit(code=1) from None
    click.get_current_context().call_on_close(lock.release)


def _excel_has_open(path: Path) -> bool:
    """Excel really has `path` open (owner file + Excel running); a leftover owner file
    from a crashed / force-quit Excel is deleted and reported (utils.atomic.open_in_excel)."""
    from dca.utils.atomic import open_in_excel

    removed: list[Path] = []
    is_open = open_in_excel(path, removed)
    for f in removed:
        typer.echo(f"🧹 Excel не запущен — удалён оставшийся служебный файл {f.name}")
    return is_open


def _excel_open_notice(input_xlsx: Path) -> list[str]:
    """Excel keeps a «~$<name>» lock file next to an open workbook."""
    if not _excel_has_open(input_xlsx):
        return []
    return [f"{input_xlsx.name} открыт в Excel: в документы попало последнее СОХРАНЁННОЕ состояние, а сделка "
            f"не перенесена в папку компании. Сохраните и закройте Excel, затем запустите generate ещё раз."]


def _place_deal(deal: Deal, deal_dir_path: Path, root: DataRoot) -> tuple[Path, tuple[str, str] | None]:
    """File the deal as <company>/Outgoing|Incoming/<num>_<quantity>kg (inside the data root);
    a deal outside the data root is only renamed in place.

    Returns (the deal's folder now, None) or (its unchanged folder, (why, what to do)).
    """
    from dca.paths import INCOMING_DIR_NAME, OUTGOING_DIR_NAME

    folder = _company_folder(deal)
    if folder and _is_under(deal_dir_path, root.path):
        parent = root.incoming_dir(folder) if deal.deal_kind == "incoming" else root.outgoing_dir(folder)
    else:
        parent = deal_dir_path.parent
    target = parent / _expected_folder_name(deal)
    if target == deal_dir_path:
        return deal_dir_path, None
    if target.exists():
        typer.echo(f"  ⚠ target folder {target} already exists, deal left in {deal_dir_path}", err=True)
        shown = target.relative_to(root.path).as_posix() if _is_under(target, root.path) else str(target)
        return deal_dir_path, (f"папка {shown} уже существует",
                               "проверьте, какая из двух папок нужная; лишнюю перенесите из DCA/ и снова нажмите run")
    parent.mkdir(parents=True, exist_ok=True)
    shutil.move(str(deal_dir_path), str(target))
    rel = target.relative_to(root.path).as_posix() if _is_under(target, root.path) else target.name
    old_company = deal_dir_path.parent.parent
    if deal_dir_path.parent.name in (OUTGOING_DIR_NAME, INCOMING_DIR_NAME) and old_company != target.parent.parent:
        typer.secho(f"  📦 MOVED to another company: {old_company.name} → {rel}", fg=typer.colors.YELLOW, bold=True)
    else:
        typer.echo(f"  ✓ filed → {rel}")
    return target, None


def _is_under(path: Path, parent: Path) -> bool:
    """True if `path` is `parent` or sits inside it."""
    try:
        path.resolve().relative_to(parent.resolve())
        return True
    except ValueError:
        return False


# «Сделка» rows that an INCOMING deal doesn't need: hidden by `dca new-incoming`.
# Matched by the machine name in column A (and by the section heading), never by a
# row number — the layout of the template may change.
OUTGOING_ONLY_FIELDS = frozenset(
    {"variant", "party_c_name", "party_c_address", "party_c_details", "letter_num", "letter_date"}
    | {f"batch_key_{n}" for n in range(1, 6)}
    | {f"batch_use_kg_{n}" for n in range(1, 6)}
)
OUTGOING_ONLY_SECTIONS = ("ВАРИАНТ", "СТОРОНА В", "ОПИСЬ ПРИЛОЖЕНИЙ", "ПАРТИИ")


def _deal_sheet_member(xlsx_path: Path) -> str:
    """Return the zip member path of the «Сделка» worksheet (e.g. xl/worksheets/sheet1.xml).

    Resolves the name→r:id→Target chain via xl/workbook.xml + its rels (any attribute
    order) rather than hardcoding sheet1.xml.
    """
    from dca.xlsx_patch import _read, _sheet_member

    data, _ = _read(xlsx_path)
    return _sheet_member(data, "Сделка")


def _outgoing_only_rows(input_path: Path) -> list[int]:
    """Row numbers of «Сделка» that hold an outgoing-only field or section heading."""
    from openpyxl import load_workbook

    wb = load_workbook(input_path, read_only=True)
    try:
        rows = []
        for r, (a,) in enumerate(wb["Сделка"].iter_rows(max_col=1, values_only=True), start=1):
            text = str(a).strip() if a is not None else ""
            if text in OUTGOING_ONLY_FIELDS or ("━" in text and any(s in text for s in OUTGOING_ONLY_SECTIONS)):
                rows.append(r)
        return rows
    finally:
        wb.close()


def _hide_outgoing_rows(input_path: Path) -> None:
    """Hide the outgoing-only «Сделка» rows in an incoming-deal Excel.

    Done by patching the worksheet XML inside the .xlsx zip directly (set
    hidden="1" on the existing <row> elements), NOT via openpyxl — an openpyxl
    save drops the extension data validations (the archive-key dropdowns). A
    direct zip/XML patch leaves every other part of the workbook (dropdowns,
    conditional formatting, sheet protection) byte-for-byte intact.
    """
    rows = _outgoing_only_rows(input_path)
    sheet_member = _deal_sheet_member(input_path)
    with zipfile.ZipFile(input_path, "r") as zin:
        infos = zin.infolist()
        data = {i.filename: zin.read(i.filename) for i in infos}

    xml = data[sheet_member].decode("utf-8")
    for r in rows:
        open_tag = f'<row r="{r}"'
        if open_tag in xml:
            xml = xml.replace(open_tag, f'<row r="{r}" hidden="1"', 1)
    data[sheet_member] = xml.encode("utf-8")

    with zipfile.ZipFile(input_path, "w") as zout:
        for i in infos:
            # Preserve each member's original compression type and metadata.
            zout.writestr(i, data[i.filename])


def _expected_folder_name(deal: Deal) -> str:
    """`<agreement_num>_<quantity_kg>kg` — same rounding rule as in documents (math.ceil)."""
    return f"{slug(deal.agreement_num)}_{math.ceil(deal.quantity_kg)}kg"


# Runner template. `@COMMAND@` is replaced by the dca subcommand. The tag line
# (`dca-runner v2 command=...`) lets `dca doctor` spot outdated runners and lets
# `dca rewrite-runners` keep each deal's command.
RUNNER_TAG = "dca-runner v2"

_RUNNER_SH = r"""#!/usr/bin/env bash
# Auto-generated by `dca`. Double-click to render the deal.
# dca-runner v2 command=@COMMAND@
# No hardcoded project path: finds the DCA code by walking up from this folder
#   - a folder with src/dca/            -> the code is right there;
#   - a .dca.toml / dca.toml with code_dir = "..."  -> the code is at that path.
# Falls back to `dca` on PATH when the code isn't found.
DEAL_DIR="$(cd "$(dirname "$0")" && pwd)"

find_code() {
    local d="$DEAL_DIR" code
    while :; do
        if [ -f "$d/src/dca/__init__.py" ]; then echo "$d"; return 0; fi
        for marker in "$d/.dca.toml" "$d/dca.toml"; do
            [ -f "$marker" ] || continue
            code="$(sed -n 's/^[[:space:]]*code_dir[[:space:]]*=[[:space:]]*"\(.*\)"[[:space:]]*$/\1/p' "$marker" | head -n 1)"
            case "$code" in
                "") ;;
                "~"*) code="$HOME${code:1}" ;;
                /*) ;;
                *) code="$d/$code" ;;
            esac
            if [ -n "$code" ] && [ -f "$code/src/dca/__init__.py" ]; then echo "$code"; return 0; fi
        done
        [ "$d" = "/" ] && return 1
        d="$(dirname "$d")"
    done
}

run_dca() {
    local code py
    if code="$(find_code)"; then
        for py in python3 /opt/miniconda3/bin/python3 "$HOME/miniconda3/bin/python3" \
                  /opt/homebrew/bin/python3 /usr/local/bin/python3; do
            if command -v "$py" >/dev/null 2>&1 && PYTHONPATH="$code/src" "$py" -c \
                'import sys, dca.cli; sys.exit(sys.version_info < (3, 11))' >/dev/null 2>&1; then
                PYTHONPATH="$code/src" "$py" -m dca @COMMAND@ "$DEAL_DIR"
                return $?
            fi
        done
        echo "⚠ DCA code found in $code, but no Python 3.11+ with its dependencies; trying dca on PATH." >&2
    fi
    if command -v dca >/dev/null 2>&1; then
        dca @COMMAND@ "$DEAL_DIR"
        return $?
    fi
    echo "❌ Cannot run DCA: no code found above $DEAL_DIR and dca is not on PATH." >&2
    echo "   See README -> Moving the project." >&2
    return 1
}

run_dca
echo ""
echo "Press any key to close..."
read -n 1 -s
"""

_RUNNER_COMMAND_RE = re.compile(r"dca-runner v\d+ command=([\w-]+)")


def _write_runners(
    deal_dir: Path, num_slug: str, command: str = "generate", name: str = RUN_COMMAND_FILENAME,
) -> None:
    """Write a double-click runner (macOS .command) into the deal folder.

    `command` is the dca subcommand the runner invokes — "generate" for outgoing
    deals (`dca new`), "generate-incoming" for incoming deals (`dca new-incoming`).

    The runners are portable: no hardcoded project path. They walk up from their
    own folder to the DCA code and run it with `python -m dca` (no `pip install`
    needed), falling back to `dca` on PATH. A deal folder can be moved anywhere
    inside the data folder, or the whole project moved, and its runner keeps working.
    """
    runner = deal_dir / name.format(num=num_slug)
    runner.write_text(_RUNNER_SH.replace("@COMMAND@", command), encoding="utf-8")
    runner.chmod(0o755)


def _write_transfer_runners(deal_dir: Path, num_slug: str) -> None:
    """transfer_<num>.command → `dca transfer` (the transfer act)."""
    _write_runners(deal_dir, num_slug, command="transfer", name=TRANSFER_RUN_COMMAND_FILENAME)


def _write_amendment_runners(deal_dir: Path, num_slug: str) -> None:
    """amend_<num>.command → `dca amend` (the amendment)."""
    _write_runners(deal_dir, num_slug, command="amend", name=AMENDMENT_RUN_COMMAND_FILENAME)


def _runner_command(deal_dir: Path, existing: list[Path]) -> str:
    """Which dca subcommand a deal's runners should call.

    Read from the existing runners' tag line; if none says, a folder under
    Incoming/ is an incoming deal.
    """
    for runner_file in existing:
        m = _RUNNER_COMMAND_RE.search(runner_file.read_text(encoding="utf-8", errors="replace"))
        if m:
            return m.group(1)
    return "generate-incoming" if deal_dir.parent.name == "Incoming" else "generate"


def _rewrite_runners_in_dir(deal_dir: Path) -> tuple[bool, str]:
    """Rewrite runners in one deal_dir using the current `_write_runners` format.

    Returns (changed, message) — `changed=False` when there's nothing to do
    (no runners found, or wrong directory shape). A forced rewrite: the deal's
    number is recovered from the existing runner's file name, the command from
    its content.
    """
    if not deal_dir.is_dir():
        return False, f"not a directory ({deal_dir})"
    candidates = sorted(deal_dir.glob("run_*.command"))
    if not candidates:
        return False, "no run_*.command found"
    num_slug = candidates[0].stem[len("run_"):]
    if not num_slug:
        return False, f"empty num slug in {candidates[0].name}"
    command = _runner_command(deal_dir, candidates)
    _write_runners(deal_dir, num_slug, command=command)
    _write_amendment_runners(deal_dir, num_slug)   # both deal kinds
    if command == "generate":  # outgoing deals also get / refresh the transfer runner
        _write_transfer_runners(deal_dir, num_slug)
        return True, f"rewrote runners for num={num_slug} (dca {command} + dca transfer + dca amend)"
    return True, f"rewrote runners for num={num_slug} (dca {command} + dca amend)"


if __name__ == "__main__":
    app()
