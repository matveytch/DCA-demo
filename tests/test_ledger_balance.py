"""«Баланс» grouped by our company + unused batches from the master database (2026-10-29)."""

from __future__ import annotations

import shutil

import pytest
from openpyxl import load_workbook
from typer.testing import CliRunner

from dca.cli import app
from dca.config import DATA_DIR
from dca.excel_io import BATCHES_SHEET, load_deal
from dca.ledger import (
    LEDGER_BALANCE_SHEET,
    NO_COMPANY_LABEL,
    UNUSED_FILL,
    _balance_layout,
    append_to_ledger,
    rebuild_ledger,
)
from dca.master_db import init_master

runner = CliRunner()
FILLED = DATA_DIR / "input_example_FILLED.xlsx"   # uses П-001 (700 kg) and П-002 (300 kg)

ALPHA = "Компания «Альфа»"
DB = [  # invented batches; the first two are the ones the example deal uses
    {"batch_key": "П-001", "batch_quantity_kg": 20000,
     "batch_company_key": "alpha", "batch_company_name_ru": ALPHA},
    {"batch_key": "П-002", "batch_quantity_kg": 18000,
     "batch_company_key": "alpha", "batch_company_name_ru": ALPHA},
    {"batch_key": "B-FRESH", "batch_quantity_kg": 5000,
     "batch_company_key": "alpha", "batch_company_name_ru": ALPHA},
    {"batch_key": "TestA", "batch_quantity_kg": 7000,
     "batch_company_key": "omega", "batch_company_name_ru": "Компания «Тест»"},
    {"batch_key": "TestB", "batch_quantity_kg": 3000,
     "batch_company_key": "omega", "batch_company_name_ru": "Компания «Тест»"},
    {"batch_key": "NoCo", "batch_quantity_kg": 1000},
]


def write_db(path, records=DB):
    """A master database whose «Партии» sheet holds `records` (one column per batch)."""
    init_master(path)
    wb = load_workbook(path)
    ws = wb[BATCHES_SHEET]
    row_of = {ws.cell(r, 1).value: r for r in range(2, ws.max_row + 1)}
    for i, rec in enumerate(records):
        for field, value in rec.items():
            ws.cell(row_of[field], 3 + i, value)
    wb.save(path)
    return path


def tx(num, key, use, car=None, company=None):
    return {"deal_agreement_num": num, "use_kg": use, "batch_key": key,
            "batch_quantity_kg": car, "batch_company_key": company}


def column_a(path):
    ws = load_workbook(path)[LEDGER_BALANCE_SHEET]
    return [ws.cell(r, 1).value for r in range(2, ws.max_row + 1)]


# ─── layout ────────────────────────────────────────────────────

def test_groups_by_company_used_first_then_unused():
    rows = [tx("E-1", "П-001", 500, 20000, "alpha"), tx("E-1", "П-002", 500, 18000, "alpha")]
    groups = _balance_layout(rows, DB)
    assert [(g["company"], [p["key"] for p in g["batches"]]) for g in groups] == [
        ("alpha", ["П-001", "П-002", "B-FRESH"]),
        ("omega", ["TestA", "TestB"]),
        (None, ["NoCo"]),                                        # no company — always last
    ]
    assert [p["unused"] for p in groups[0]["batches"]] == [False, False, True]
    assert groups[0]["open"] == 3 and groups[0]["name"] == ALPHA


def test_company_with_more_open_batches_goes_first():
    # alpha's two used batches are spent to 0 → it has 1 open (B-FRESH), omega has 2
    rows = [tx("E-1", "П-001", 20000, 20000, "alpha"), tx("E-2", "П-002", 18000, 18000, "alpha")]
    groups = _balance_layout(rows, DB)
    assert [g["company"] for g in groups] == ["omega", "alpha", None]
    assert [p["key"] for p in groups[1]["batches"]] == ["П-001", "П-002", "B-FRESH"]


def test_ties_keep_first_appearance():
    rows = [tx("E-1", "TestA", 100, 7000, "omega"), tx("E-2", "П-001", 100, 20000, "alpha")]
    db = [r for r in DB if r["batch_key"] in ("TestA", "П-001")]
    assert [g["company"] for g in _balance_layout(rows, db)] == ["omega", "alpha"]


def test_weight_and_company_come_from_the_database_when_rows_lack_them():
    rows = [tx("OLD-1", "B-FRESH", 1000)]                       # a hand-typed row: only A, D, E
    pkg = _balance_layout(rows, DB)[0]["batches"][0]
    assert (pkg["key"], pkg["capacity"], pkg["company"], pkg["remaining"], pkg["unused"]) == \
           ("B-FRESH", 5000, "alpha", 4000, False)


def test_without_database_only_used_batches():
    rows = [tx("E-1", "X", 10, 100, None)]
    groups = _balance_layout(rows, [])
    assert [(g["company"], [p["key"] for p in g["batches"]]) for g in groups] == [(None, ["X"])]


# ─── written sheet ─────────────────────────────────────────────

@pytest.fixture
def ledger(tmp_path):
    path = tmp_path / "Data" / "batch_ledger.xlsx"
    path.parent.mkdir(parents=True)
    write_db(path.parent / "master_db.xlsx")
    append_to_ledger(load_deal(FILLED), path, deal_id="deal-1")   # master_db.xlsx next to it is used
    return path


def test_sheet_has_company_headings_formulas_and_light_green_unused(ledger):
    assert column_a(ledger) == [
        f"Компания: {ALPHA} (alpha)", "П-001", "П-002", "B-FRESH",
        "Компания: Компания «Тест» (omega)", "TestA", "TestB",
        NO_COMPANY_LABEL, "NoCo",
    ]
    ws = load_workbook(ledger)[LEDGER_BALANCE_SHEET]
    assert ws["D2"].value == "=SUM(D3:D5)"
    assert ws["E2"].value == '="открытых партий: "&COUNTIF(D3:D5,">0")'
    assert ws["F2"].value == "alpha"
    assert ws["B5"].value == 5000 and ws["C5"].value.startswith("=SUMIF(") and ws["D5"].value == "=B5-C5"
    assert ws["A5"].fill.fgColor.rgb.endswith(UNUSED_FILL)          # B-FRESH: not used yet
    assert ws["F5"].fill.fgColor.rgb.endswith(UNUSED_FILL)
    assert not ws["A3"].fill.fgColor.rgb.endswith(UNUSED_FILL)      # П-001: used


def test_unreadable_database_still_writes_the_ledger(tmp_path):
    path = tmp_path / "batch_ledger.xlsx"
    (tmp_path / "master_db.xlsx").write_bytes(b"not an xlsx")
    result = append_to_ledger(load_deal(FILLED), path, deal_id="deal-1")
    assert any("не удалось прочитать" in w for w in result["warnings"])
    assert column_a(path) == ["Компания: alpha", "П-001", "П-002"]


def test_rebuild_picks_up_a_batch_added_to_the_database(ledger):
    records = DB + [{"batch_key": "B-NEW", "batch_quantity_kg": 900,
                     "batch_company_key": "alpha", "batch_company_name_ru": ALPHA}]
    (ledger.parent / "master_db.xlsx").unlink()
    write_db(ledger.parent / "master_db.xlsx", records)
    assert rebuild_ledger(ledger) == (7, [])
    assert column_a(ledger)[:5] == [f"Компания: {ALPHA} (alpha)", "П-001", "П-002", "B-FRESH", "B-NEW"]


# ─── end to end ────────────────────────────────────────────────

def test_generate_and_check_ledger_show_unused_batches(tmp_path, monkeypatch):
    home = tmp_path / "home"
    monkeypatch.setenv("DCA_HOME", str(home))
    write_db(home / "Data" / "master_db.xlsx")
    assert runner.invoke(app, ["new", "1-2026-1705"]).exit_code == 0
    deal = home / "Новые сделки" / "1-2026-1705_pending"
    shutil.copy(FILLED, deal / "input_1-2026-1705.xlsx")
    result = runner.invoke(app, ["generate", str(deal)])
    assert result.exit_code == 0, result.output
    ledger = home / "Data" / "batch_ledger.xlsx"
    assert "B-FRESH" in column_a(ledger) and "TestA" in column_a(ledger)

    result = runner.invoke(app, ["check-ledger"])
    assert result.exit_code == 0, result.output
    assert "«Баланс» пересобран: 6 парт(ий)" in result.output
