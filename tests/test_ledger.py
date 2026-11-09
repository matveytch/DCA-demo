"""Unit tests for dca.ledger — batch transaction log."""

from __future__ import annotations

from datetime import datetime

import pytest
from openpyxl import load_workbook

from dca.config import DATA_DIR
from dca.excel_io import load_deal
from dca.ledger import (
    LEDGER_BALANCE_SHEET,
    LEDGER_TX_SHEET,
    TX_COLUMNS,
    _build_rows_for_deal,
    _detect_drift,
    _read_tx_rows,
    append_to_ledger,
)

FILLED_PATH = DATA_DIR / "input_example_FILLED.xlsx"


@pytest.fixture
def deal_filled():
    return load_deal(FILLED_PATH)


@pytest.fixture
def deal_no_batches(deal_filled):
    """A FILLED deal but with all batch_key cleared → no active batches."""
    from dca.models import Deal
    overrides = {f"batch_key_{n}": None for n in (1, 2, 3, 4, 5)}
    # Also clear use weights so weight-sum validator doesn't blow up
    overrides.update({f"batch_use_kg_{n}": None for n in (1, 2, 3, 4, 5)})
    return Deal(**deal_filled.model_copy(update=overrides).model_dump())


# ─── _build_rows_for_deal ────────────────────────────────────────


def test_build_rows_uses_active_batches_only(deal_filled):
    rows = _build_rows_for_deal(deal_filled)
    # FILLED has 2 active batches (П-001 + П-002)
    assert len(rows) == 2
    assert all(r["deal_agreement_num"] == deal_filled.agreement_num for r in rows)
    keys = {r["batch_key"] for r in rows}
    assert keys == {"П-001", "П-002"}


def test_build_rows_uses_ceil_for_weights(deal_filled):
    """Weights in ledger rows should be ceil-rounded integers (matches docs rendering)."""
    rows = _build_rows_for_deal(deal_filled)
    # 500.0 → ceil = 500 (integer); FILLED has whole-number weights
    assert all(isinstance(r["use_kg"], int) for r in rows)
    assert all(isinstance(r["deal_quantity_kg"], int) for r in rows)
    # capacity_kg is also ceil-int when present
    for r in rows:
        if r["batch_quantity_kg"] is not None:
            assert isinstance(r["batch_quantity_kg"], int)


def test_build_rows_full_snapshot_columns(deal_filled):
    """Every TX_COLUMNS key should be present in each row dict."""
    rows = _build_rows_for_deal(deal_filled)
    for r in rows:
        for col in TX_COLUMNS:
            assert col in r, f"missing column {col!r} in row {r['batch_key']!r}"


def test_build_rows_dates_as_iso_strings(deal_filled):
    rows = _build_rows_for_deal(deal_filled)
    for r in rows:
        # ISO format yyyy-mm-dd, 10 chars, when present
        for col in ("deal_agreement_date", "batch_agreement_date", "batch_annex_date"):
            v = r.get(col)
            if v is not None:
                assert isinstance(v, str) and len(v) == 10 and v[4] == "-" and v[7] == "-", \
                    f"{col}={v!r} should be yyyy-mm-dd"


# ─── _detect_drift ────────────────────────────────────────────────


def test_drift_identical_rows_no_warnings():
    rows = [{"deal_agreement_num": "X", "batch_key": "K", "use_kg": 100,
             **{c: None for c in TX_COLUMNS if c not in ("deal_agreement_num", "batch_key", "use_kg")}}]
    assert _detect_drift(rows, rows) == []


def test_drift_changed_use_kg_warns():
    base_old = {"deal_agreement_num": "X", "batch_key": "K", "use_kg": 100,
                **{c: None for c in TX_COLUMNS if c not in ("deal_agreement_num", "batch_key", "use_kg")}}
    base_new = {**base_old, "use_kg": 200}
    warnings = _detect_drift([base_old], [base_new])
    assert any("use_kg" in w and "100" in w and "200" in w for w in warnings)


def test_drift_added_key_warns():
    old = []
    new = [{c: None for c in TX_COLUMNS} | {"deal_agreement_num": "X", "batch_key": "K-new"}]
    warnings = _detect_drift(old, new)
    assert any("K-new" in w and "добавлен" in w for w in warnings)


def test_drift_removed_key_warns():
    old = [{c: None for c in TX_COLUMNS} | {"deal_agreement_num": "X", "batch_key": "K-old"}]
    new = []
    warnings = _detect_drift(old, new)
    assert any("K-old" in w and "больше не" in w for w in warnings)


# ─── append_to_ledger (end-to-end) ───────────────────────────────


def test_append_creates_new_ledger(deal_filled, tmp_path):
    ledger = tmp_path / "ledger.xlsx"
    assert not ledger.exists()
    result = append_to_ledger(deal_filled, ledger_path=ledger)
    assert ledger.exists()
    assert result["skipped"] is False
    assert len(result["added"]) == 2  # two active batches
    assert len(result["removed"]) == 0
    assert result["warnings"] == []

    # Verify sheets
    wb = load_workbook(ledger)
    assert LEDGER_TX_SHEET in wb.sheetnames
    assert LEDGER_BALANCE_SHEET in wb.sheetnames

    # Check header + 2 data rows in Транзакции
    ws = wb[LEDGER_TX_SHEET]
    headers = [c.value for c in ws[1]]
    assert headers == TX_COLUMNS
    data = list(ws.iter_rows(min_row=2, values_only=True))
    # Filter out fully-empty rows
    data = [r for r in data if any(v is not None for v in r)]
    assert len(data) == 2


def test_append_overwrites_same_deal(deal_filled, tmp_path):
    ledger = tmp_path / "ledger.xlsx"
    # First write
    r1 = append_to_ledger(deal_filled, ledger_path=ledger)
    assert len(r1["added"]) == 2
    # Re-run with identical deal → old removed, new added, no warnings
    r2 = append_to_ledger(deal_filled, ledger_path=ledger)
    assert len(r2["removed"]) == 2
    assert len(r2["added"]) == 2
    assert r2["warnings"] == []
    # Total rows still 2 (not 4 — overwrote, not appended)
    rows, _extras = _read_tx_rows(ledger)
    assert len(rows) == 2


def test_append_drift_detection_on_modified_use_kg(deal_filled, tmp_path):
    ledger = tmp_path / "ledger.xlsx"
    append_to_ledger(deal_filled, ledger_path=ledger)
    # Modify use_kg for batch 1 (700 → 900) AND the quantity (1000 → 1200)
    # to keep the batch-sum validator happy; one annex instead of two
    from dca.models import Deal
    modified = Deal(**deal_filled.model_copy(update={
        "quantity_kg": 1200.0,
        "annexes": [],
        "batch_use_kg_1": 900.0,
        "batch_use_kg_2": 300.0,
    }).model_dump())
    result = append_to_ledger(modified, ledger_path=ledger)
    drift_warnings = [w for w in result["warnings"] if "use_kg" in w]
    assert len(drift_warnings) >= 1, f"expected a use_kg drift warning, got: {result['warnings']}"


def test_append_skipped_when_no_batches(deal_no_batches, tmp_path):
    ledger = tmp_path / "ledger.xlsx"
    result = append_to_ledger(deal_no_batches, ledger_path=ledger)
    assert result["skipped"] is True
    assert result["added"] == []
    assert not ledger.exists()


def test_append_multiple_deals_coexist(deal_filled, tmp_path):
    """Two different outgoing deals — both should be in the ledger after writes."""
    ledger = tmp_path / "ledger.xlsx"
    # Write deal A
    append_to_ledger(deal_filled, ledger_path=ledger)
    # Write a second deal with a different agreement_num
    from dca.models import Deal
    deal_b = Deal(**deal_filled.model_copy(update={
        "agreement_num": "2-2026-9999",
    }).model_dump())
    append_to_ledger(deal_b, ledger_path=ledger)
    rows, _extras = _read_tx_rows(ledger)
    # 2 batches × 2 deals = 4 rows
    assert len(rows) == 4
    deal_nums = {r["deal_agreement_num"] for r in rows}
    assert deal_nums == {"1-2026-1705", "2-2026-9999"}


def test_balance_sheet_has_formulas(deal_filled, tmp_path):
    ledger = tmp_path / "ledger.xlsx"
    append_to_ledger(deal_filled, ledger_path=ledger)
    wb = load_workbook(ledger)
    ws = wb[LEDGER_BALANCE_SHEET]
    # Row 2 is the company heading (2026-10-29); row 3 — the first batch: SUMIF in C (used_total_kg)
    cell = ws.cell(row=3, column=3)
    assert isinstance(cell.value, str) and cell.value.startswith("=SUMIF"), \
        f"Expected SUMIF formula, got {cell.value!r}"
    # Column D = remaining_kg = B - C
    rem = ws.cell(row=3, column=4).value
    assert isinstance(rem, str) and rem.startswith("=B")


def test_append_blocks_when_lock_held(deal_filled, tmp_path, monkeypatch):
    """If another process holds the ledger lock, append should time out fast
    in tests and surface as PermissionError (not hang)."""
    from filelock import FileLock

    import dca.ledger as ledger_mod

    # Shrink the timeout for the test so we don't wait 30s
    monkeypatch.setattr(ledger_mod, "_LEDGER_LOCK_TIMEOUT_SEC", 0.5)

    ledger = tmp_path / "ledger.xlsx"
    lock_path = tmp_path / "ledger.xlsx.lock"
    other_lock = FileLock(str(lock_path))
    other_lock.acquire()
    try:
        with pytest.raises(PermissionError, match="Cannot acquire ledger lock"):
            append_to_ledger(deal_filled, ledger_path=ledger)
    finally:
        other_lock.release()

    # After release, the same call should succeed
    result = append_to_ledger(deal_filled, ledger_path=ledger)
    assert not result["skipped"]


def test_drift_no_warning_when_date_string_vs_datetime(deal_filled):
    """If Excel reformatted an ISO date string back into a datetime cell,
    `_detect_drift` must treat them as equal (no false-positive warning).

    Simulates: 1) we wrote `"2026-01-15"` (string).
               2) Operator opened ledger in Excel — cell got reformatted as date.
               3) Next `_read_tx_rows` returns `datetime(2026,1,15)`.
               4) Next pipeline run builds fresh row with `"2026-01-15"` again.
    Expected: no drift, nothing changed semantically.
    """
    new_rows = _build_rows_for_deal(deal_filled)
    # Simulate read-from-Excel: datetime objects for date columns
    old_rows = []
    for r in new_rows:
        r2 = dict(r)
        # Convert any ISO-string dates into datetime as Excel would
        for col in ("deal_agreement_date", "batch_agreement_date", "batch_annex_date"):
            v = r2.get(col)
            if isinstance(v, str) and len(v) == 10 and v[4] == "-":
                r2[col] = datetime.fromisoformat(v)
        old_rows.append(r2)

    warnings = _detect_drift(old_rows, new_rows)
    assert warnings == [], f"Expected no false-positive drift, got: {warnings}"


def test_drift_no_warning_when_int_vs_float(deal_filled):
    """Similarly, if Excel/openpyxl returns 500.0 where we wrote 500,
    treat them as equal."""
    new_rows = _build_rows_for_deal(deal_filled)
    old_rows = []
    for r in new_rows:
        r2 = dict(r)
        for col in ("deal_quantity_kg", "use_kg", "batch_quantity_kg"):
            if isinstance(r2.get(col), int):
                r2[col] = float(r2[col])
        old_rows.append(r2)

    warnings = _detect_drift(old_rows, new_rows)
    assert warnings == [], f"Expected no false-positive drift on int↔float, got: {warnings}"


def test_drift_still_catches_real_changes(deal_filled):
    """Sanity: normalization must not mask real changes — use_kg actually different."""
    new_rows = _build_rows_for_deal(deal_filled)
    old_rows = [dict(r) for r in new_rows]
    old_rows[0]["use_kg"] = old_rows[0]["use_kg"] + 1   # genuine change

    warnings = _detect_drift(old_rows, new_rows)
    assert any("use_kg" in w for w in warnings)


def test_extra_columns_preserved_across_regenerate(deal_filled, tmp_path):
    """Operator adds a `notes` column to «Транзакции» in Excel and types in
    a comment. Next `dca generate` for the same deal MUST keep the column
    and the value — overwriting the engine-managed columns only.
    """
    from openpyxl import load_workbook

    ledger = tmp_path / "ledger.xlsx"
    # Run 1: pipeline writes baseline
    append_to_ledger(deal_filled, ledger_path=ledger)

    # Simulate operator: open ledger, add 'notes' column, type a value in row 2
    wb = load_workbook(ledger)
    ws = wb[LEDGER_TX_SHEET]
    notes_col = len(TX_COLUMNS) + 1  # next column after managed ones
    ws.cell(row=1, column=notes_col, value="notes")
    ws.cell(row=2, column=notes_col, value="operator's note for first row")
    wb.save(ledger)

    # Run 2: pipeline runs again for same deal
    append_to_ledger(deal_filled, ledger_path=ledger)

    # Verify: notes column still exists, value still there
    wb2 = load_workbook(ledger)
    ws2 = wb2[LEDGER_TX_SHEET]
    headers = [ws2.cell(row=1, column=i).value for i in range(1, ws2.max_column + 1)]
    assert "notes" in headers, f"'notes' column was wiped; headers: {headers}"

    rows, extras = _read_tx_rows(ledger)
    assert "notes" in extras
    # The note we typed should round-trip via row dict
    notes_values = [r.get("notes") for r in rows]
    assert "operator's note for first row" in notes_values


def test_extra_column_position_after_managed_columns(deal_filled, tmp_path):
    """Extras must be written to the right of TX_COLUMNS, never mixed in."""
    from openpyxl import load_workbook

    ledger = tmp_path / "ledger.xlsx"
    append_to_ledger(deal_filled, ledger_path=ledger)

    # Add extras
    wb = load_workbook(ledger)
    ws = wb[LEDGER_TX_SHEET]
    ws.cell(row=1, column=len(TX_COLUMNS) + 1, value="комментарий")
    wb.save(ledger)

    # Re-run
    append_to_ledger(deal_filled, ledger_path=ledger)

    wb2 = load_workbook(ledger)
    ws2 = wb2[LEDGER_TX_SHEET]
    headers = [ws2.cell(row=1, column=i).value for i in range(1, ws2.max_column + 1)]
    # First N must be exactly TX_COLUMNS, in order
    assert headers[: len(TX_COLUMNS)] == list(TX_COLUMNS)
    # Then extras
    assert "комментарий" in headers[len(TX_COLUMNS):]


def test_balance_formulas_have_5000_row_headroom(deal_filled, tmp_path):
    """SUMIF/TEXTJOIN ranges must extend far enough that future ledger growth
    (new deals appended by the pipeline, manual rows added by the operator)
    is picked up without rebuilding the workbook. 5000 = thousands of deals
    × batches — comfortable headroom for the ledger's lifetime.
    """
    ledger = tmp_path / "ledger.xlsx"
    append_to_ledger(deal_filled, ledger_path=ledger)
    wb = load_workbook(ledger)
    ws = wb[LEDGER_BALANCE_SHEET]
    sumif = ws.cell(row=3, column=3).value            # row 2 — company heading
    textjoin = ws.cell(row=3, column=5).value.text   # an ArrayFormula
    assert "$5000" in sumif, f"SUMIF should cover up to row 5000, got: {sumif}"
    assert "$5000" in textjoin, f"TEXTJOIN should cover up to row 5000, got: {textjoin}"


def test_balance_sheet_lists_each_unique_archive_key_once(deal_filled, tmp_path):
    ledger = tmp_path / "ledger.xlsx"
    append_to_ledger(deal_filled, ledger_path=ledger)
    wb = load_workbook(ledger)
    ws = wb[LEDGER_BALANCE_SHEET]
    keys = [ws.cell(row=r, column=1).value for r in range(2, ws.max_row + 1)
            if ws.cell(row=r, column=1).value and not str(ws.cell(row=r, column=1).value).startswith("Компания")]
    assert keys == ["П-001", "П-002"]


# ─── Company columns (audit 2026-10-28) ─────────────────────────

def test_rows_carry_party_a_and_batch_owner(deal_filled):
    from dca.ledger import _build_rows_for_deal
    rows = _build_rows_for_deal(deal_filled)
    assert {r["deal_party_a_key"] for r in rows} == {"alpha"}
    assert [r["batch_company_key"] for r in rows] == ["alpha", "alpha"]


def test_balance_formula_columns_did_not_move():
    """«Баланс» formulas use Транзакции!A/D/E by letter — new columns go at the end."""
    from openpyxl.utils import get_column_letter
    pos = {c: get_column_letter(i) for i, c in enumerate(TX_COLUMNS, start=1)}
    assert (pos["deal_agreement_num"], pos["use_kg"], pos["batch_key"],
            pos["batch_quantity_kg"]) == ("A", "D", "E", "L")
    assert TX_COLUMNS[12:] == ["deal_party_a_key", "batch_company_key", "deal_id"]


def test_old_ledger_without_company_columns_is_backfilled_quietly(tmp_path, deal_filled):
    """A ledger written before the company columns existed (with an operator «notes» column):
    regenerating the same deal fills the companies in, keeps the notes, raises no drift warning."""
    from openpyxl import Workbook

    from dca.ledger import _build_rows_for_deal
    old_cols = [c for c in TX_COLUMNS if c not in ("deal_party_a_key", "batch_company_key")]
    ledger = tmp_path / "batch_ledger.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Транзакции"
    ws.append([*old_cols, "notes"])
    for row in _build_rows_for_deal(deal_filled):
        ws.append([row[c] for c in old_cols] + ["my note"])
    wb.save(ledger)

    result = append_to_ledger(deal_filled, ledger_path=ledger)
    assert result["warnings"] == []
    ws = load_workbook(ledger)["Транзакции"]
    headers = [c.value for c in ws[1]]
    assert headers == [*TX_COLUMNS, "notes"]
    row2 = {h: ws.cell(2, i + 1).value for i, h in enumerate(headers)}
    assert row2["deal_party_a_key"] == "alpha"
    assert row2["batch_company_key"] == "alpha"
    assert row2["notes"] == "my note"


def test_balance_shows_batch_owner_company(tmp_path, deal_filled):
    ledger = tmp_path / "batch_ledger.xlsx"
    append_to_ledger(deal_filled, ledger_path=ledger)
    ws = load_workbook(ledger)["Баланс"]
    assert ws.cell(1, 6).value == "batch_company_key"
    assert {ws.cell(r, 6).value for r in range(2, ws.max_row + 1)} == {"alpha"}


def test_changing_the_batch_owner_is_reported_as_drift(tmp_path, deal_filled):
    ledger = tmp_path / "batch_ledger.xlsx"
    append_to_ledger(deal_filled, ledger_path=ledger)
    changed = deal_filled.model_copy(update={"batch_company_key_1": "omega"})
    warnings = append_to_ledger(changed, ledger_path=ledger)["warnings"]
    assert any("batch_company_key" in w and "omega" in w for w in warnings)


def test_deals_using_it_is_an_xlfn_array_formula(tmp_path, deal_filled):
    """Plain `=TEXTJOIN(...)` showed #NAME? in Excel / LibreOffice (found by the 2026-10-28 audit)."""
    from openpyxl.worksheet.formula import ArrayFormula
    ledger = tmp_path / "batch_ledger.xlsx"
    append_to_ledger(deal_filled, ledger_path=ledger)
    cell = load_workbook(ledger)["Баланс"]["E3"].value   # E2 — the company heading
    assert isinstance(cell, ArrayFormula)
    assert cell.text.startswith("=_xlfn.TEXTJOIN(")
