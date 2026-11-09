"""`dca check-ledger`: rows typed by hand into the batch ledger (deals made before DCA),
checked the way «Баланс» counts them, and «Баланс» rebuilt without a generate."""

from __future__ import annotations

import pytest
from openpyxl import load_workbook
from typer.testing import CliRunner

from dca.cli import app
from dca.config import DATA_DIR
from dca.excel_io import load_deal
from dca.ledger import (
    LEDGER_BALANCE_SHEET,
    LEDGER_TX_SHEET,
    LEDGER_VERSIONS_DIR,
    _read_tx_rows,
    append_to_ledger,
    check_ledger,
    rebuild_ledger,
)

runner = CliRunner()
FILLED = DATA_DIR / "input_example_FILLED.xlsx"   # batches П-001 (5 000 kg, uses 700) and П-002 (2 000 kg, uses 300)

# What the operator types for an old deal: the four columns «Баланс» counts, plus a date.
OLD = {"deal_agreement_num": "OLD-2025-01", "deal_agreement_date": "2025-11-03",
       "use_kg": 3000, "batch_key": "П-001"}


@pytest.fixture
def ledger(tmp_path):
    """A ledger written by DCA for the example deal (2 batches, with their weights)."""
    path = tmp_path / "Data" / "batch_ledger.xlsx"
    append_to_ledger(load_deal(FILLED), path, deal_id="deal-1")
    return path


def add_rows(path, *rows, extra_sheet=None):
    """Type rows into «Транзакции» the way the operator does in Excel (by column name)."""
    wb = load_workbook(path)
    ws = wb[LEDGER_TX_SHEET]
    headers = [c.value for c in ws[1]]
    for row in rows:
        ws.append([row.get(h) for h in headers])
    if extra_sheet:
        wb.create_sheet(extra_sheet)
    wb.save(path)


def balance(report):
    return {b["key"]: b for b in report.balance}


# ─── check ─────────────────────────────────────────────────────

def test_good_hand_rows_count_in_the_balance(ledger):
    add_rows(ledger, OLD, {**OLD, "deal_agreement_num": "OLD-2025-02", "use_kg": 1500,
                           "batch_key": "B-OLD", "batch_quantity_kg": 10000})
    report = check_ledger(ledger, batch_keys={"П-001", "П-002", "B-OLD"})
    assert report.errors == [] and report.warnings == []
    b = balance(report)
    assert b["П-001"]["used"] == 3700 and b["П-001"]["remaining"] == 1300   # 700 (DCA) + 3000 (hand)
    assert b["П-001"]["deals"] == ["1-2026-1705", "OLD-2025-01"]
    assert b["B-OLD"]["capacity"] == 10000 and b["B-OLD"]["remaining"] == 8500


def test_row_without_deal_number_would_be_dropped(ledger):
    add_rows(ledger, {**OLD, "deal_agreement_num": None})
    report = check_ledger(ledger)
    assert report.errors == ["строка 4: нет deal_agreement_num (A) — при следующей записи реестра строка пропадёт"]


@pytest.mark.parametrize("value, problem", [
    ("3000", "записано текстом — Excel не учтёт его"),
    ("=1000+2000", "формула «=1000+2000»"),
    (0, "должно быть больше 0"),
    (None, "пусто"),
])
def test_use_kg_must_be_a_positive_number(ledger, value, problem):
    add_rows(ledger, {**OLD, "use_kg": value})
    report = check_ledger(ledger)
    assert len(report.errors) == 1
    assert report.errors[0].startswith("строка 4 (OLD-2025-01, партия П-001): use_kg (D)")
    assert problem in report.errors[0]


def test_batch_key_problems(ledger):
    add_rows(ledger, {**OLD, "batch_key": None},
             {**OLD, "deal_agreement_num": "OLD-2", "batch_key": "П-001 "},
             {**OLD, "deal_agreement_num": "OLD-3", "batch_key": "Pak-typo",
              "batch_quantity_kg": 5000})
    report = check_ledger(ledger, batch_keys={"П-001", "П-002"})
    assert "строка 4 (OLD-2025-01): нет batch_key (E) — строка не попадёт в «Баланс»" in report.errors
    assert any("строка 5 (OLD-2): в batch_key «П-001 » лишние пробелы" in e for e in report.errors)
    assert any("строка 6 (OLD-3): партии «Pak-typo» нет в базе" in w for w in report.warnings)


def test_batch_without_its_weight_is_an_error(ledger):
    add_rows(ledger, {**OLD, "batch_key": "B-OLD"})
    report = check_ledger(ledger)
    assert report.errors == ["партия B-OLD: нет batch_quantity_kg (L) ни в одной строке — "
                             "остаток не посчитается"]


def test_weight_text_is_an_error(ledger):
    add_rows(ledger, {**OLD, "batch_quantity_kg": "20 000"})
    report = check_ledger(ledger)
    assert any("batch_quantity_kg (L) «20 000» записано текстом" in e for e in report.errors)


def test_warnings_double_write_off_different_weight_overuse(ledger):
    add_rows(ledger, OLD, OLD,
             {**OLD, "deal_agreement_num": "OLD-2", "use_kg": 17000, "batch_quantity_kg": 19000})
    report = check_ledger(ledger)
    assert report.errors == []
    assert any("сделка OLD-2025-01, партия П-001: строки 4, 5 — списание учтено 2 раза" in w
               for w in report.warnings)
    assert any("партия П-001: количество партии различается" in w for w in report.warnings)
    # 700 + 3000 + 3000 + 17000 = 23 700 kg against the last quantity 19 000 kg
    assert any("ПЕРЕРАСХОД 4700 кг" in w for w in report.warnings)


def test_broken_structure_is_reported_not_raised(ledger):
    add_rows(ledger, extra_sheet="Заметки")
    report = check_ledger(ledger)
    assert len(report.errors) == 1 and "посторонние листы Заметки" in report.errors[0]


# ─── rebuild ───────────────────────────────────────────────────

def test_rebuild_adds_hand_typed_batches_to_the_balance(ledger):
    add_rows(ledger, {**OLD, "batch_key": "B-OLD", "batch_quantity_kg": 10000, "заметка": "x"})
    wb = load_workbook(ledger)
    wb[LEDGER_TX_SHEET].cell(1, wb[LEDGER_TX_SHEET].max_column + 1, "заметка")   # operator's own column
    wb[LEDGER_TX_SHEET].cell(4, wb[LEDGER_TX_SHEET].max_column, "старая сделка")
    wb.save(ledger)
    assert "B-OLD" not in [c.value for c in load_workbook(ledger)[LEDGER_BALANCE_SHEET]["A"]]

    assert rebuild_ledger(ledger) == (3, [])
    wb = load_workbook(ledger)
    bal = wb[LEDGER_BALANCE_SHEET]
    row = next(r for r in range(2, bal.max_row + 1) if bal.cell(r, 1).value == "B-OLD")
    assert bal.cell(row, 2).value == 10000                       # batch weight
    assert bal.cell(row, 3).value.startswith("=SUMIF(")          # used — a live formula
    rows, extras = _read_tx_rows(ledger)
    old = next(r for r in rows if r["deal_agreement_num"] == "OLD-2025-01")
    assert old["use_kg"] == 3000 and old["заметка"] == "старая сделка" and extras == ["заметка"]
    assert list((ledger.parent / LEDGER_VERSIONS_DIR).iterdir())   # previous file kept


def test_next_generate_keeps_hand_rows(ledger):
    add_rows(ledger, OLD)
    rebuild_ledger(ledger)
    other = load_deal(FILLED).model_copy(update={"agreement_num": "1-2026-2000"})
    append_to_ledger(other, ledger, deal_id="deal-2")
    nums = [r["deal_agreement_num"] for r in _read_tx_rows(ledger)[0]]
    assert nums.count("OLD-2025-01") == 1 and "1-2026-1705" in nums and "1-2026-2000" in nums


# ─── `dca check-ledger` ────────────────────────────────────────

@pytest.fixture
def home(tmp_path, monkeypatch):
    h = tmp_path / "home"
    monkeypatch.setenv("DCA_HOME", str(h))
    append_to_ledger(load_deal(FILLED), h / "Data" / "batch_ledger.xlsx", deal_id="deal-1")
    return h


def test_cli_ok_rebuilds_the_balance(home):
    path = home / "Data" / "batch_ledger.xlsx"
    add_rows(path, {**OLD, "batch_key": "B-OLD", "batch_quantity_kg": 10000})
    result = runner.invoke(app, ["check-ledger"])
    assert result.exit_code == 0, result.output
    assert "B-OLD" in result.output and "OLD-2025-01" in result.output
    assert "«Баланс» пересобран: 3 парт(ий)" in result.output
    assert "B-OLD" in [c.value for c in load_workbook(path)[LEDGER_BALANCE_SHEET]["A"]]


def test_cli_errors_leave_the_file_untouched(home):
    path = home / "Data" / "batch_ledger.xlsx"
    add_rows(path, {**OLD, "use_kg": "3000"})
    before = path.read_bytes()
    result = runner.invoke(app, ["check-ledger"])
    assert result.exit_code == 1
    assert "use_kg (D) «3000» записано текстом" in result.output and "реестр не изменён" in result.output
    assert path.read_bytes() == before


def test_cli_dry_run_writes_nothing(home):
    path = home / "Data" / "batch_ledger.xlsx"
    add_rows(path, OLD)
    before = path.read_bytes()
    result = runner.invoke(app, ["check-ledger", "--dry-run"])
    assert result.exit_code == 0 and "не пересобран" in result.output
    assert path.read_bytes() == before


def test_cli_refuses_while_excel_has_the_ledger_open(home):
    path = home / "Data" / "batch_ledger.xlsx"
    add_rows(path, OLD)
    (path.parent / f"~${path.name}").write_bytes(b"lock")
    before = path.read_bytes()
    result = runner.invoke(app, ["check-ledger"])
    assert result.exit_code == 1 and "открыт в Excel" in result.output
    assert path.read_bytes() == before


def test_cli_without_ledger(tmp_path, monkeypatch):
    monkeypatch.setenv("DCA_HOME", str(tmp_path / "empty"))
    result = runner.invoke(app, ["check-ledger"])
    assert result.exit_code == 1 and "Ledger not found" in result.output
