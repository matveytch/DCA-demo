"""Master database (Data/master_db.xlsx): init, read/validate, snapshot into new deals, refresh-db."""

import html
import re
import shutil
import zipfile
from datetime import datetime

import pytest
from openpyxl import load_workbook
from typer.testing import CliRunner

from dca.cli import app
from dca.config import EMPTY_TEMPLATE
from dca.excel_io import load_deal
from dca.master_db import ARCHIVES, init_master, read_archives, template_layout
from dca.xlsx_patch import fill_archive_sheets
from tests.helpers import FILLED

runner = CliRunner()
EMPTY_COUNT = "Стороны А 0, Стороны Б 0, Изделия 0, Партии 0"


@pytest.fixture
def db(home):
    """A database holding the example's records (copied from the FILLED example's archives)."""
    path = home / "Data" / "master_db.xlsx"
    init_master(path)
    _write_records(path, read_archives(FILLED, strict=False).records)
    return path


def _write_records(path, records):
    wb = load_workbook(path)
    for sheet, recs in records.items():
        ws = wb[sheet]
        row_of = {ws.cell(r, 1).value: r for r in range(2, ws.max_row + 1)}
        for c in range(3, ws.max_column + 1):
            for r in row_of.values():
                ws.cell(r, c).value = None
        for i, rec in enumerate(recs):
            for field, value in rec.items():
                ws.cell(row_of[field], 3 + i, value)
    wb.save(path)


def _set(path, sheet, field, col, value):
    wb = load_workbook(path)
    ws = wb[sheet]
    row = next(r for r in range(2, ws.max_row + 1) if ws.cell(r, 1).value == field)
    ws.cell(row, col, value)
    wb.save(path)


# ─── init / read ───────────────────────────────────────────────

def test_init_db_creates_template_fields_and_refuses_overwrite(home):
    result = runner.invoke(app, ["init-db"])
    assert result.exit_code == 0, result.output
    path = home / "Data" / "master_db.xlsx"
    wb = load_workbook(path)
    assert wb.sheetnames == list(ARCHIVES)
    for sheet, fields in template_layout().items():
        ws = wb[sheet]
        assert [ws.cell(r, 1).value for r in range(2, ws.max_row + 1)] == [f for f, _, _ in fields]
    assert read_archives(path, strict=True).count_line() == EMPTY_COUNT
    again = runner.invoke(app, ["init-db"])
    assert again.exit_code == 1 and "already exists" in again.output


def test_duplicate_key_and_unknown_field_are_errors(db):
    _set(db, "Стороны Б", "party_b_key", 5, "beta")                  # C has it already
    wb = load_workbook(db)
    wb["Изделия"].cell(wb["Изделия"].max_row + 1, 1, "item_nmae_ru")  # typo in a field name
    wb.save(db)
    with pytest.raises(ValueError) as e:
        read_archives(db, strict=True)
    assert "duplicate key 'beta' in columns C and E" in str(e.value)
    assert "unknown field 'item_nmae_ru'" in str(e.value)


def test_dates_become_text_and_numbers_stay_numbers(db):
    _set(db, "Партии", "batch_agreement_date", 3, datetime(2026, 1, 15))
    _set(db, "Партии", "batch_quantity_kg", 3, 20000.0)
    rec = read_archives(db, strict=True).records["Партии"][0]
    assert rec["batch_agreement_date"] == "15.01.2026"
    assert rec["batch_quantity_kg"] == 20000


# ─── dca new snapshot ──────────────────────────────────────────

def test_new_copies_all_records_and_deal_loads(db, home):
    result = runner.invoke(app, ["new", "1-2026-1705"])
    assert result.exit_code == 0, result.output
    assert "archives from the database: Стороны А 2, Стороны Б 2, Изделия 2, Партии 2" in result.output
    xlsx = home / "Новые сделки" / "1-2026-1705_pending" / "input_1-2026-1705.xlsx"
    assert read_archives(xlsx, strict=False).records == read_archives(db, strict=True).records

    # choose keys on «Сделка» like the example → the deal resolves them from the snapshot
    filled = load_workbook(FILLED)["Сделка"]
    wb = load_workbook(xlsx)
    ws = wb["Сделка"]
    for r in range(1, filled.max_row + 1):
        name = filled.cell(r, 1).value
        for r2 in range(1, ws.max_row + 1):
            if ws.cell(r2, 1).value == name and isinstance(name, str) and not name.startswith(" "):
                ws.cell(r2, 3).value = filled.cell(r, 3).value
    wb.save(xlsx)
    deal = load_deal(xlsx)
    assert deal.party_b_name_eng == "Beta Ltd"


def test_new_without_db_warns_and_keeps_archives_empty(home):
    result = runner.invoke(app, ["new", "1-2026-NODB"])
    assert result.exit_code == 0, result.output
    assert "no master database" in result.output
    xlsx = home / "Новые сделки" / "1-2026-NODB_pending" / "input_1-2026-NODB.xlsx"
    assert read_archives(xlsx, strict=False).count_line() == EMPTY_COUNT


def test_new_incoming_also_gets_the_snapshot(db, home):
    assert runner.invoke(app, ["new-incoming", "1-2026-INC"]).exit_code == 0
    xlsx = home / "Новые сделки" / "1-2026-INC_incoming_pending" / "input_1-2026-INC.xlsx"
    assert read_archives(xlsx, strict=False).keys("Стороны Б") == ["beta", "alpha-b"]


def test_more_than_ten_records_widen_dropdowns_and_preview(db, tmp_path):
    recs = read_archives(db, strict=True).records
    recs["Стороны Б"] = [{"party_b_key": f"B{i}", "party_b_name_eng": f"PARTY {i}"} for i in range(1, 13)]
    xlsx = tmp_path / "in.xlsx"
    shutil.copy(EMPTY_TEMPLATE, xlsx)
    deal_sheet_before = zipfile.ZipFile(xlsx).read("xl/worksheets/sheet1.xml").decode()
    fill_archive_sheets(xlsx, recs)

    z = zipfile.ZipFile(xlsx)
    deal_sheet = html.unescape(z.read("xl/worksheets/sheet1.xml").decode())
    deal_sheet_before = html.unescape(deal_sheet_before)
    assert "Стороны Б'!$C$2:$N$2" in deal_sheet                  # 12 records → C..N
    assert "Стороны А'!$C$2:$L$2" in deal_sheet                  # untouched archives keep C..L
    norm = lambda x: re.sub(r"\$[A-Z]+\$2<", "", x)  # noqa: E731
    assert norm(deal_sheet) == norm(deal_sheet_before)          # nothing else on «Сделка» changed
    preview = html.unescape(z.read("xl/worksheets/sheet7.xml").decode())
    assert set(re.findall(r"Стороны Б'!\$C\$2:\$([A-Z]+)\$", preview)) == {"N"}
    ws = load_workbook(xlsx)["Стороны Б"]
    assert ws.cell(1, 14).value == "Запись #12" and ws.cell(2, 14).value == "B12"


def test_multiline_values_survive(db, tmp_path):
    _set(db, "Стороны Б", "party_b_address_full_eng", 3, "Line one,\nLine two")
    xlsx = tmp_path / "in.xlsx"
    shutil.copy(EMPTY_TEMPLATE, xlsx)
    fill_archive_sheets(xlsx, read_archives(db, strict=True).records)
    rec = read_archives(xlsx, strict=False).records["Стороны Б"][0]
    assert rec["party_b_address_full_eng"] == "Line one,\nLine two"


# ─── refresh-db ────────────────────────────────────────────────

def _deal(tmp_path):
    d = tmp_path / "1-2026-1705_pending"
    d.mkdir()
    shutil.copy(FILLED, d / "input_1-2026-1705.xlsx")
    return d


def test_refresh_db_updates_archives_with_summary_and_backup(db, tmp_path):
    d = _deal(tmp_path)
    xlsx = d / "input_1-2026-1705.xlsx"
    before_values = [[c.value for c in row] for row in load_workbook(xlsx)["Сделка"].iter_rows()]

    _set(db, "Стороны Б", "party_b_address_full_eng", 3, "New address 1")   # change «beta»
    _set(db, "Стороны Б", "party_b_key", 5, "gamma")                        # add a record
    wb = load_workbook(db)                                                   # remove «Y» (column D)
    ws = wb["Изделия"]
    for r in range(2, ws.max_row + 1):
        ws.cell(r, 4).value = None
    wb.save(db)

    result = runner.invoke(app, ["refresh-db", str(d)])
    assert result.exit_code == 0, result.output
    assert "+ gamma" in result.output
    assert "~ beta (party_b_address_full_eng)" in result.output
    assert "− Y" in result.output
    assert len(list((d / "_archive").glob("input_before_refresh_*.xlsx"))) == 1
    assert read_archives(xlsx, strict=False).keys("Стороны Б") == ["beta", "alpha-b", "gamma"]
    after_values = [[c.value for c in row] for row in load_workbook(xlsx)["Сделка"].iter_rows()]
    assert after_values == before_values                                    # «Сделка» untouched
    assert load_deal(xlsx).party_b_address_full_eng == "New address 1"


def test_refresh_db_refuses_when_selected_key_vanished(db, tmp_path):
    d = _deal(tmp_path)
    xlsx = d / "input_1-2026-1705.xlsx"
    before = xlsx.read_bytes()
    _set(db, "Стороны Б", "party_b_key", 3, "renamed")              # the deal uses «beta»
    result = runner.invoke(app, ["refresh-db", str(d)])
    assert result.exit_code == 1
    assert "party_b_key = 'beta'" in result.output
    assert xlsx.read_bytes() == before
    assert not (d / "_archive").exists()


def test_refresh_db_nothing_to_do(db, tmp_path):
    d = _deal(tmp_path)
    result = runner.invoke(app, ["refresh-db", str(d)])
    assert result.exit_code == 0 and "Already up to date" in result.output
    assert not (d / "_archive").exists()


def test_refresh_db_refuses_open_excel_and_missing_db(tmp_path, home):
    d = _deal(tmp_path)
    no_db = runner.invoke(app, ["refresh-db", str(d)])
    assert no_db.exit_code == 1 and "dca init-db" in no_db.output
    (d / "~$input_1-2026-1705.xlsx").write_bytes(b"lock")
    opened = runner.invoke(app, ["refresh-db", str(d)])
    assert opened.exit_code == 1 and "open in Excel" in opened.output
