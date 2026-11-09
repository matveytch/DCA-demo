"""Batch ledger.

Persistent record of which outgoing deal used which source batch for how many
kg. Stored in `Data/batch_ledger.xlsx` with two sheets:

  «Транзакции» — event log, one row per active batch per deal
                 (overwritten when the same deal is regenerated)
  «Баланс»     — computed via Excel SUMIF formulas — for each batch:
                 its whole quantity, total used, remaining

The «Баланс» sheet uses live formulas so Excel always shows the current
state on open; no need for a separate `dca balance` command.

Public API:
  append_to_ledger(deal: Deal) -> dict
  check_ledger(path) / rebuild_ledger(path)   — `dca check-ledger`

Pipeline integration: called from `cli.py:_run_pipeline` after the documents
are published.
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any

from filelock import FileLock, Timeout
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.formula import ArrayFormula

from dca.models import Deal
from dca.paths import resolve_data_root
from dca.utils.atomic import refuse_if_open_in_excel, save_workbook_atomic

# How long to wait for the ledger lock before giving up. 30 s is comfortable
# for back-to-back `dca generate` runs but short enough that a stuck/orphaned
# lock surfaces quickly rather than hanging the CLI.
_LEDGER_LOCK_TIMEOUT_SEC = 30.0


LEDGER_TX_SHEET = "Транзакции"
LEDGER_BALANCE_SHEET = "Баланс"

# Canonical column order for «Транзакции» sheet.
# Row 1 of the sheet holds these as headers.
TX_COLUMNS = [
    "deal_agreement_num",        # A
    "deal_agreement_date",
    "deal_quantity_kg",
    "use_kg",                    # D
    "batch_key",                 # E
    "batch_agreement_num",
    "batch_agreement_date",
    "batch_annex_num",
    "batch_annex_date",
    "batch_party_name_ru",
    "batch_party_location_ru",
    "batch_quantity_kg",         # L
    # The «Баланс» formulas reference columns A, D, E by letter, so new columns
    # must only ever be appended at the END.
    "deal_party_a_key",          # our company of the outgoing deal (party A key)
    "batch_company_key",         # our company that owns the batch
    "deal_id",                   # the deal's identity (.dca-deal.json)
]

# Columns a row typed by hand may leave empty: filling them in on the next generate
# is a backfill, not a change worth a warning.
BACKFILL_COLUMNS = {"deal_party_a_key", "batch_company_key", "deal_id"}

# «Баланс» sheet columns
BALANCE_COLUMNS = [
    "batch_key",
    "batch_quantity_kg",
    "used_total_kg",
    "remaining_kg",
    "deals_using_it",
    "batch_company_key",
]


# ─── Helpers ─────────────────────────────────────────────────────


def _iso_or_none(d: date | None) -> str | None:
    """ISO-format a date (or None → None). xlsx-safe."""
    if d is None:
        return None
    return d.isoformat()


def _build_rows_for_deal(deal: Deal, deal_id: str | None = None) -> list[dict[str, Any]]:
    """Build the new ledger rows for this deal — one per active batch.

    Quantities are ceil-ed to match what gets printed in documents.
    """
    rows: list[dict[str, Any]] = []
    for n in deal.active_batches():
        use_kg_raw = getattr(deal, f"batch_use_kg_{n}")
        use_kg = math.ceil(use_kg_raw) if use_kg_raw is not None else None
        capacity_raw = getattr(deal, f"batch_quantity_kg_{n}")
        capacity = math.ceil(capacity_raw) if capacity_raw is not None else None

        row = {
            "deal_agreement_num":     deal.agreement_num,
            "deal_agreement_date":    _iso_or_none(deal.agreement_date),
            "deal_quantity_kg":  math.ceil(deal.quantity_kg),
            "use_kg":                  use_kg,
            "batch_key":      getattr(deal, f"batch_key_{n}"),
            "batch_agreement_num":     getattr(deal, f"batch_agreement_num_{n}"),
            "batch_agreement_date":    _iso_or_none(getattr(deal, f"batch_agreement_date_{n}")),
            "batch_annex_num":  getattr(deal, f"batch_annex_num_{n}"),
            "batch_annex_date": _iso_or_none(getattr(deal, f"batch_annex_date_{n}")),
            "batch_party_name_ru":     getattr(deal, f"batch_party_name_ru_{n}"),
            "batch_party_location_ru": getattr(deal, f"batch_party_location_ru_{n}"),
            "batch_quantity_kg":     capacity,
            "deal_party_a_key":      deal.party_a_key,
            "batch_company_key":     getattr(deal, f"batch_company_key_{n}"),
            "deal_id":                        deal_id,
        }
        rows.append(row)
    return rows


def _read_tx_rows(path: Path) -> tuple[list[dict[str, Any]], list[str]]:
    """Read «Транзакции» — return (rows, extra_column_names).

    `rows` contain ALL columns (including extras the operator added to the
    right of TX_COLUMNS), keyed by column name. `extra_column_names` lists
    the operator's extra columns in their original sheet order so the writer
    can re-emit them in the same position.

    Why preserve extras: operators may add a `notes` / `комментарий` column
    next to ours. Rewriting the workbook from scratch would silently destroy
    their annotations.

    Empty cells in extras become None (consistent with TX_COLUMNS handling).
    Columns with `None` header (truly empty column) are ignored.
    """
    if not path.exists():
        return [], []
    refuse_if_open_in_excel(path)
    wb = load_workbook(path)
    _check_structure(wb, path)
    ws = wb[LEDGER_TX_SHEET]

    # Header row — detect which columns are managed (TX_COLUMNS) vs extra
    headers_raw = [c.value for c in ws[1]]
    headers: list[str | None] = [h if isinstance(h, str) and h else None for h in headers_raw]
    extras: list[str] = [h for h in headers if h and h not in TX_COLUMNS]

    rows: list[dict[str, Any]] = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if all(v is None for v in r):
            continue
        row: dict[str, Any] = {}
        for col_name, val in zip(headers, r, strict=False):  # rows may be shorter than the header
            if col_name is None:
                continue   # unnamed column — drop
            row[col_name] = val
        if row.get("deal_agreement_num"):
            rows.append(row)
    return rows, extras


class LedgerStructureError(ValueError):
    """An existing ledger is not in the shape DCA writes — nothing is changed."""


def _check_structure(wb: Any, path: Path) -> None:
    """An existing ledger must be readable as a whole.

    A missing «Транзакции» sheet or a renamed header used to read as «no rows»,
    and the next save started the history from scratch; any other sheet would be
    dropped because the ledger is rebuilt on every save.
    """
    problems = []
    others = [s for s in wb.sheetnames if s not in (LEDGER_TX_SHEET, LEDGER_BALANCE_SHEET)]
    if LEDGER_TX_SHEET not in wb.sheetnames:
        problems.append(f"нет листа «{LEDGER_TX_SHEET}»")
    else:
        headers = [str(c.value).strip() for c in wb[LEDGER_TX_SHEET][1] if c.value not in (None, "")]
        missing = [c for c in TX_COLUMNS if c not in BACKFILL_COLUMNS and c not in headers]
        doubled = sorted({h for h in headers if headers.count(h) > 1})
        if missing:
            problems.append(f"на листе «{LEDGER_TX_SHEET}» нет колонок {', '.join(missing)}")
        if doubled:
            problems.append(f"повторяются колонки {', '.join(doubled)}")
    if others:
        problems.append(f"посторонние листы {', '.join(others)} (DCA пересобирает реестр и потерял бы их — "
                        f"перенесите их в отдельный файл)")
    if problems:
        raise LedgerStructureError(
            f"структура реестра {path.name} нарушена: {'; '.join(problems)}. Реестр не изменён; "
            f"прежние версии — в {path.parent.name}/{LEDGER_VERSIONS_DIR}/.")


def _normalize_for_compare(v: Any) -> Any:
    """Coerce a value to a canonical form for round-trip-safe equality checks.

    Why this exists — without it, drift detection cries wolf:
      - We write `"2026-01-15"` (ISO string) for dates. If the operator opens
        the ledger in Excel and saves, Excel auto-detects the string as a date
        and reformats the cell as date-type → openpyxl on next read returns
        `datetime(2026,1,15)`. Raw `!=` would scream "field changed", but
        nothing changed semantically.
      - openpyxl can return `500.0` (float) for a cell where we wrote `500`
        (int) under some conditions. Same false-positive.

    Canonical forms:
      - datetime/date  → ISO date string (`"YYYY-MM-DD"`), matching `_iso_or_none`
      - float-that-is-an-integer → `int(v)`, matching `math.ceil` outputs
      - everything else → unchanged
    """
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    if isinstance(v, float) and v == int(v):
        return int(v)
    return v


def _row_signature(row: dict[str, Any]) -> tuple:
    """Hashable signature of a row excluding deal-meta — used for drift detection.

    Two rows of the same deal_agreement_num are "the same" if the batch side
    of the row matches. If anything batch-related changes (use_kg, archive_key,
    any of the snapshot fields) — drift.
    """
    return tuple(row.get(c) for c in TX_COLUMNS)


def _detect_drift(old_rows: list[dict], new_rows: list[dict]) -> list[str]:
    """Compare old (existing) rows of same deal_agreement_num vs new ones.

    Returns list of human-readable warnings. Empty list = no drift.
    """
    warnings: list[str] = []
    # Index old rows by batch_key
    old_by_key = {r.get("batch_key"): r for r in old_rows}
    new_by_key = {r.get("batch_key"): r for r in new_rows}

    deal_num = (new_rows[0]["deal_agreement_num"]
                  if new_rows else (old_rows[0]["deal_agreement_num"] if old_rows else "?"))

    removed_keys = set(old_by_key) - set(new_by_key)
    added_keys = set(new_by_key) - set(old_by_key)
    common_keys = set(old_by_key) & set(new_by_key)

    for k in removed_keys:
        warnings.append(
            f"Сделка {deal_num}: партия {k!r} больше не используется (был в прошлой записи)"
        )
    for k in added_keys:
        warnings.append(
            f"Сделка {deal_num}: партия {k!r} добавлен (раньше не использовался)"
        )
    for k in common_keys:
        old = old_by_key[k]
        new = new_by_key[k]
        for col in TX_COLUMNS:
            ov = _normalize_for_compare(old.get(col))
            nv = _normalize_for_compare(new.get(col))
            if col in BACKFILL_COLUMNS and ov is None:
                continue
            if ov != nv:
                warnings.append(
                    f"Сделка {deal_num}, партия {k}: {col} был {ov!r}, стал {nv!r}"
                )
    return warnings


def _capacity_inconsistency_warnings(all_rows: list[dict]) -> list[str]:
    """Across the entire ledger, warn if same batch_key has different
    `batch_quantity_kg` recorded in different rows. Values are
    normalized before comparison so that `500` vs `500.0` round-trip-equal."""
    by_key: dict[str, set] = {}
    for r in all_rows:
        k = r.get("batch_key")
        v = _normalize_for_compare(r.get("batch_quantity_kg"))
        if k and v is not None:
            by_key.setdefault(k, set()).add(v)
    warnings = []
    for k, values in by_key.items():
        if len(values) > 1:
            warnings.append(
                f"Партия {k!r}: batch_quantity_kg различается между сделками: {sorted(values)}"
            )
    return warnings


# ─── Workbook writing ────────────────────────────────────────────


def _styles():
    return {
        "header_font": Font(bold=True, color="FFFFFF"),
        "header_fill": PatternFill("solid", fgColor="1F4E78"),
        "subheader_fill": PatternFill("solid", fgColor="D9E1F2"),
        "border": Border(*(Side("thin", color="CCCCCC"),) * 4),
    }


def _write_tx_sheet(
    ws,
    rows: list[dict[str, Any]],
    extras: list[str] | None = None,
) -> None:
    """Write «Транзакции» sheet — managed columns first, then operator extras.

    `extras` is a list of operator-added column names (e.g. `notes`,
    `комментарий`). They are appended to the right of TX_COLUMNS in the
    header, and each row writes its value for that column from `row[name]`
    (None / missing → blank cell).

    Styling is applied to managed columns only — extras keep the cell
    default so they don't visually masquerade as engine-managed.
    """
    extras = extras or []
    st = _styles()
    all_cols = list(TX_COLUMNS) + extras
    # Header
    for i, col in enumerate(all_cols, start=1):
        c = ws.cell(row=1, column=i, value=col)
        if col in TX_COLUMNS:
            c.font = st["header_font"]
            c.fill = st["header_fill"]
            c.alignment = Alignment(horizontal="center", vertical="center")
        else:
            # Mark extras visually so operator can tell engine vs theirs
            c.font = Font(bold=True, italic=True, color="555555")
            c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 22
    # Column widths (heuristic)
    widths = {
        "deal_agreement_num": 18, "deal_agreement_date": 13, "deal_quantity_kg": 12,
        "use_kg": 10,
        "batch_key": 16, "batch_agreement_num": 18, "batch_agreement_date": 13,
        "batch_annex_num": 16, "batch_annex_date": 13,
        "batch_party_name_ru": 28, "batch_party_location_ru": 24,
        "batch_quantity_kg": 15,
        "deal_party_a_key": 18, "batch_company_key": 18,
    }
    for i, col in enumerate(all_cols, start=1):
        ws.column_dimensions[get_column_letter(i)].width = widths.get(col, 20)
    # Data rows
    for r_idx, row in enumerate(rows, start=2):
        for c_idx, col in enumerate(all_cols, start=1):
            value = row.get(col)
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            if isinstance(value, str):
                cell.data_type = "s"     # «=1+1» typed as a number stays text
    ws.freeze_panes = "A2"


MASTER_DB_FILENAME = "master_db.xlsx"   # lives next to the ledger in Data/ (paths.DataRoot.master_db_path)
NO_COMPANY_LABEL = "Компания не указана"
UNUSED_FILL = "E2EFDA"                  # light green: a batch from the database with no write-off yet


def _db_packages(master_db_path: Path | None) -> tuple[list[dict[str, Any]], list[str]]:
    """Batches of the master database («Партии») → (records, warnings).

    No database → no records. An unreadable one never blocks the ledger: «Баланс»
    is written without the unused batches and a warning says why.
    """
    if master_db_path is None or not master_db_path.exists():
        return [], []
    from dca.excel_io import BATCHES_SHEET
    from dca.master_db import read_archives
    try:
        return list(read_archives(master_db_path, strict=False).records.get(BATCHES_SHEET, [])), []
    except Exception as e:
        return [], [f"«Баланс» без неиспользованных партий: базу {master_db_path.name} не удалось прочитать ({e})"]


def _number(v: Any) -> float | None:
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def _balance_layout(tx_rows: list[dict[str, Any]], db_records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Batches of «Баланс» grouped by our company that owns them.

    Inside a group: batches with write-offs in their first-seen order in «Транзакции»,
    then the database's batches nobody used yet (database order, `unused`). Groups:
    the company with more open batches (remaining > 0) first, ties by first
    appearance; batches without a company last. Quantity and company come from
    «Транзакции», the database fills what is empty there.
    """
    db = {str(r["batch_key"]): r for r in db_records if r.get("batch_key")}
    entries: dict[str, dict[str, Any]] = {}
    for r in tx_rows:
        key = r.get("batch_key")
        if not key:
            continue
        e = entries.setdefault(str(key), {"key": key, "capacity": None, "company": None,
                                          "used": 0.0, "unused": False})
        if r.get("batch_quantity_kg") is not None:
            e["capacity"] = r["batch_quantity_kg"]          # last write wins
        if r.get("batch_company_key"):
            e["company"] = r["batch_company_key"]
        e["used"] += _number(r.get("use_kg")) or 0.0
    for k, e in entries.items():
        rec = db.get(k, {})
        if e["capacity"] is None:
            e["capacity"] = rec.get("batch_quantity_kg")
        e["company"] = e["company"] or rec.get("batch_company_key")
    for k, rec in db.items():
        if k not in entries:
            entries[k] = {"key": rec["batch_key"], "capacity": rec.get("batch_quantity_kg"),
                          "company": rec.get("batch_company_key"), "used": 0.0, "unused": True}

    names = {r.get("batch_company_key"): r.get("batch_company_name_ru")
             for r in reversed(db_records) if r.get("batch_company_key")}
    groups: dict[Any, dict[str, Any]] = {}
    for e in entries.values():
        cap = _number(e["capacity"])
        e["remaining"] = None if cap is None else cap - e["used"]
        g = groups.setdefault(e["company"] or None, {"company": e["company"] or None,
                                                    "name": names.get(e["company"]), "batches": [], "open": 0})
        g["open"] += e["remaining"] is not None and e["remaining"] > 0
    for e in [e for e in entries.values() if not e["unused"]] + [e for e in entries.values() if e["unused"]]:
        groups[e["company"] or None]["batches"].append(e)
    ordered = sorted((g for g in groups.values() if g["company"]), key=lambda g: -g["open"])
    return ordered + [g for g in groups.values() if not g["company"]]


def _write_balance_sheet(ws: Any, all_tx_rows: list[dict[str, Any]],
                         db_records: list[dict[str, Any]] | None = None) -> int:
    """Write «Баланс»: a heading row per company, then one row per batch, with
    live Excel formulas referencing «Транзакции». → number of batch rows.

    Order and grouping — `_balance_layout`; batches from the database with no
    write-off yet get a light-green fill.
    """
    st = _styles()
    # Header
    for i, col in enumerate(BALANCE_COLUMNS, start=1):
        c = ws.cell(row=1, column=i, value=col)
        c.font = st["header_font"]
        c.fill = st["header_fill"]
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 22
    ws.column_dimensions["A"].width = 34  # batch_key (primary key); company headings
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 16
    ws.column_dimensions["E"].width = 60
    ws.column_dimensions["F"].width = 20

    # Column letters in «Транзакции» sheet (matches TX_COLUMNS order):
    #   A=deal_agreement_num, D=use_kg, E=batch_key
    tx = f"'{LEDGER_TX_SHEET}'"
    # Pad the formula range generously so future inserts in «Транзакции» (manual
    # edits by operator, or new deals up to the next pipeline run) still get
    # picked up by SUMIF/TEXTJOIN. 5000 rows = thousands of deals × batches —
    # comfortable headroom for the lifetime of the ledger.
    last_row = max(len(all_tx_rows) + 1, 5000)
    key_col_e    = f"{tx}!$E$2:$E${last_row}"   # batch_key
    use_col_d    = f"{tx}!$D$2:$D${last_row}"   # use_kg
    deal_col_a = f"{tx}!$A$2:$A${last_row}"   # deal_agreement_num
    unused_fill = PatternFill("solid", fgColor=UNUSED_FILL)

    r_idx = 2
    batches = 0
    for group in _balance_layout(all_tx_rows, db_records or []):
        head, first, last = r_idx, r_idx + 1, r_idx + len(group["batches"])
        company = group["company"]
        label = (f"Компания: {group['name']} ({company})" if group["name"] else f"Компания: {company}"
                 ) if company else NO_COMPANY_LABEL
        ws.cell(row=head, column=1, value=label)
        ws.cell(row=head, column=4, value=f"=SUM(D{first}:D{last})")
        ws.cell(row=head, column=5, value=f'="открытых партий: "&COUNTIF(D{first}:D{last},">0")')
        ws.cell(row=head, column=6, value=company)
        for c_idx in range(1, len(BALANCE_COLUMNS) + 1):
            ws.cell(row=head, column=c_idx).fill = st["subheader_fill"]
            ws.cell(row=head, column=c_idx).font = Font(bold=True)

        for r_idx, pkg in enumerate(group["batches"], start=first):
            key = pkg["key"]
            # A: batch_key (primary key in this sheet)
            ws.cell(row=r_idx, column=1, value=key)
            # B: capacity_kg — latest seen snapshot (static, captured at write time)
            ws.cell(row=r_idx, column=2, value=pkg["capacity"])
            # C: used_total_kg = SUMIF over use_kg where archive_key matches A{r}
            ws.cell(
                row=r_idx, column=3,
                value=f"=SUMIF({key_col_e},$A{r_idx},{use_col_d})",
            )
            # D: remaining_kg = capacity - used
            ws.cell(row=r_idx, column=4, value=f"=B{r_idx}-C{r_idx}")
            # E: deals_using_it = comma-joined outgoing agreement numbers.
            # TEXTJOIN is a post-2010 function: in the file it must be written as
            # `_xlfn.TEXTJOIN` (else Excel / LibreOffice show #NAME?), and IF over a
            # range needs an array formula — otherwise implicit intersection breaks it.
            ws[f"E{r_idx}"] = ArrayFormula(
                f"E{r_idx}",
                f'=_xlfn.TEXTJOIN(", ",TRUE,IF({key_col_e}=$A{r_idx},{deal_col_a},""))',
            )
            # F: which company owns the batch (latest snapshot, static)
            ws.cell(row=r_idx, column=6, value=pkg["company"])
            if pkg["unused"]:
                for c_idx in range(1, len(BALANCE_COLUMNS) + 1):
                    ws.cell(row=r_idx, column=c_idx).fill = unused_fill
            batches += 1
        r_idx = last + 1
    ws.freeze_panes = "A2"
    return batches


def _save_workbook(
    path: Path,
    tx_rows: list[dict[str, Any]],
    extras: list[str] | None = None,
    db_records: list[dict[str, Any]] | None = None,
) -> int:
    """Build a fresh workbook (overwriting any existing) with both sheets.

    Extra columns the operator added to «Транзакции» (passed as `extras`)
    are preserved — appended to the right of the managed columns, with each
    row's value taken from `row[extra_name]`. `db_records` — the master
    database's batches, so «Баланс» also lists the unused ones.
    → number of batches in «Баланс».
    """
    wb = Workbook()
    wb.remove(wb.active)
    ws_tx = wb.create_sheet(LEDGER_TX_SHEET)
    _write_tx_sheet(ws_tx, tx_rows, extras=extras)
    ws_bal = wb.create_sheet(LEDGER_BALANCE_SHEET)
    batches = _write_balance_sheet(ws_bal, tx_rows, db_records)
    path.parent.mkdir(parents=True, exist_ok=True)
    _atomic_save(wb, path)
    return batches


LEDGER_VERSIONS_DIR = "_ledger_versions"
LEDGER_VERSIONS_KEPT = 10


def _atomic_save(wb: Workbook, path: Path) -> None:
    """Save without ever leaving a half-written ledger: temp file →
    re-open → previous version into <dir>/_ledger_versions/ (last 10) → os.replace."""
    save_workbook_atomic(wb, path, versions_dir=path.parent / LEDGER_VERSIONS_DIR,
                         keep=LEDGER_VERSIONS_KEPT, required_sheets=[LEDGER_TX_SHEET])


def _overspend_warnings(all_rows: list[dict[str, Any]], new_rows: list[dict[str, Any]]) -> list[str]:
    """Packages used by THIS deal whose total write-off exceeds their netto.

    Computed from the transactions themselves (not Excel's cached formulas), after
    this deal's old rows were replaced by the new ones. A warning, not a refusal —
    the user's decision.
    """
    out = []
    for key in dict.fromkeys(r.get("batch_key") for r in new_rows):
        rows = [r for r in all_rows if r.get("batch_key") == key]
        capacity = next((r.get("batch_quantity_kg") for r in reversed(rows)
                         if r.get("batch_quantity_kg") is not None), None)
        used = sum(float(r.get("use_kg") or 0) for r in rows)
        if capacity is not None and used > float(capacity):
            deals = ", ".join(dict.fromkeys(str(r.get("deal_agreement_num")) for r in rows))
            out.append(f"Партия «{key}»: всего {math.ceil(float(capacity))} кг, списано всего "
                       f"{math.ceil(used)} кг — ПЕРЕРАСХОД {math.ceil(used - float(capacity))} кг "
                       f"(сделки: {deals})")
    return out


# ─── Public API ──────────────────────────────────────────────────


def append_to_ledger(
    deal: Deal, ledger_path: Path | None = None, deal_id: str | None = None,
    master_db_path: Path | None = None,
) -> dict[str, Any]:
    """Append/overwrite this deal's rows in «Транзакции».

    `master_db_path` (default: master_db.xlsx next to the ledger) adds the
    database's unused batches to «Баланс».

    Returns:
        {
            "ledger_path": Path,
            "removed":     list[dict],   # rows of this deal_agreement_num that were replaced
            "added":       list[dict],   # rows just written
            "warnings":    list[str],    # data-drift / inconsistency messages
            "skipped":     bool,         # True if no active batches → nothing to write
        }
    """
    if ledger_path is None:
        ledger_path = resolve_data_root().ledger_path
    if master_db_path is None:
        master_db_path = ledger_path.parent / MASTER_DB_FILENAME

    new_rows = _build_rows_for_deal(deal, deal_id)
    if not new_rows and not ledger_path.exists():
        # No active batches and no ledger yet — nothing to log
        return {
            "ledger_path": ledger_path,
            "removed": [],
            "added": [],
            "warnings": [],
            "skipped": True,
        }

    # Serialize concurrent `dca generate` runs. The lock guards the
    # read-modify-write sequence below — without it, two simultaneous calls
    # would each read the same starting state and the slower one would
    # silently overwrite the faster one's row.
    #
    # The lock file lives next to the ledger as `<name>.lock`. It is a
    # zero-byte sentinel managed by `filelock`; safe to delete manually if
    # ever orphaned (filelock auto-recovers).
    ledger_path.parent.mkdir(parents=True, exist_ok=True)
    lock_path = ledger_path.with_suffix(ledger_path.suffix + ".lock")
    try:
        with FileLock(str(lock_path), timeout=_LEDGER_LOCK_TIMEOUT_SEC):
            existing, extras = _read_tx_rows(ledger_path)

            # A row is this deal's by its deal_id; rows written before deal_id existed
            # by the agreement number. The same number under ANOTHER deal_id is a
            # different deal — never overwrite it.
            def mine(r: dict[str, Any]) -> bool:
                if deal_id and r.get("deal_id"):
                    return bool(r["deal_id"] == deal_id)
                return bool(r.get("deal_agreement_num") == deal.agreement_num)

            if deal_id:
                taken = sorted({str(r.get("deal_party_a_key")) for r in existing
                                if r.get("deal_agreement_num") == deal.agreement_num
                                and r.get("deal_id") and r["deal_id"] != deal_id})
                if taken:
                    raise ValueError(
                        f"номер {deal.agreement_num} в реестре принадлежит другой сделке "
                        f"(сторона А {', '.join(taken)}) — строки не тронуты. Дайте этой сделке другой номер.")
            old_rows = [r for r in existing if mine(r)]
            other_rows = [r for r in existing if not mine(r)]

            # Rows of this number written by ANOTHER party A are never replaced silently
            #. A different deal with the same number is stopped earlier
            # (dca generate refuses a number used by another deal folder), so here it
            # is the same deal moved to another of our companies — say so loudly.
            foreign = sorted({str(r.get("deal_party_a_key")) for r in old_rows
                              if r.get("deal_party_a_key")
                              and r.get("deal_party_a_key") != deal.party_a_key})

            if not new_rows and not old_rows:
                return {"ledger_path": ledger_path, "removed": [], "added": [], "warnings": [],
                        "skipped": True}

            warnings: list[str] = []
            if foreign:
                warnings.append(f"Сделка {deal.agreement_num}: в реестре была записана компанией "
                                f"{', '.join(foreign)}, теперь — {deal.party_a_key}; её строки "
                                f"переписаны. Если это ДРУГАЯ сделка с тем же номером — исправьте номер.")
            if old_rows and not new_rows:
                # The deal no longer uses any batch: its old write-offs go
                # away instead of staying in the balance.
                warnings.append(f"Сделка {deal.agreement_num}: партии больше не выбраны — "
                                f"прежнее списание ({len(old_rows)} строк(и)) снято из реестра")
            if old_rows:
                warnings.extend(_detect_drift(old_rows, new_rows))

            # Carry forward operator's extra-column values when (deal_agreement_num,
            # batch_key) is unchanged. Without this, a `notes` cell typed
            # by the operator for a row would survive `dca generate` only by luck.
            # If the row is gone (archive_key no longer used by this deal), the note
            # is dropped — by design (the row is no longer in the ledger).
            if extras:
                old_by_pkg = {
                    (r.get("deal_agreement_num"), r.get("batch_key")): r
                    for r in old_rows
                }
                for new_row in new_rows:
                    key = (new_row.get("deal_agreement_num"),
                           new_row.get("batch_key"))
                    prev = old_by_pkg.get(key)
                    if prev is not None:
                        for ex in extras:
                            if ex not in new_row and prev.get(ex) is not None:
                                new_row[ex] = prev[ex]

            all_rows = other_rows + new_rows

            # Cross-deal capacity consistency check
            warnings.extend(_capacity_inconsistency_warnings(all_rows))
            warnings.extend(_overspend_warnings(all_rows, new_rows))

            db_records, db_warnings = _db_packages(master_db_path)
            warnings.extend(db_warnings)
            try:
                _save_workbook(ledger_path, all_rows, extras=extras, db_records=db_records)
            except PermissionError as e:
                # Most likely: operator has the xlsx open in Excel — file is locked
                raise PermissionError(
                    f"Cannot write {ledger_path.name} — file is locked. "
                    f"Закрой {ledger_path.name} в Excel и повтори команду."
                ) from e
    except Timeout as e:
        raise PermissionError(
            f"Cannot acquire ledger lock within {_LEDGER_LOCK_TIMEOUT_SEC:.0f}s "
            f"({lock_path.name}) — another `dca generate` may be running, or a "
            f"stale lock file exists. If you're sure nothing else is running, "
            f"delete {lock_path} manually."
        ) from e

    return {
        "ledger_path": ledger_path,
        "removed": old_rows,
        "added": new_rows,
        "warnings": warnings,
        "skipped": False,
    }


# ─── Check + rebuild (rows typed by hand, `dca check-ledger`) ─────


@dataclass
class LedgerCheck:
    """Result of check_ledger: errors block the rebuild, warnings don't."""

    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    # One entry per batch, in first-seen order: key, capacity, used, remaining, deals.
    balance: list[dict[str, Any]] = field(default_factory=list)


def _kg(v: Any) -> tuple[float | None, str | None]:
    """A weight cell → (number, None) or (None, what is wrong with it)."""
    if v is None or (isinstance(v, str) and not v.strip()):
        return None, "пусто"
    if isinstance(v, bool):
        return None, f"«{v}» — не число"
    if isinstance(v, (int, float)):
        return (float(v), None) if v > 0 else (None, f"{v} — должно быть больше 0")
    text = str(v).strip()
    if text.startswith("="):
        return None, f"формула «{text}» — введите число"
    return None, f"«{text}» записано текстом — Excel не учтёт его в «Балансе», введите число"


def check_ledger(path: Path, batch_keys: set[str] | None = None) -> LedgerCheck:
    """Check «Транзакции» the way «Баланс» counts it.

    Rows typed by hand (deals made before DCA) must have what the balance uses:
    deal_agreement_num (A; a row without it is dropped on the next write),
    use_kg (D) and batch_quantity_kg (L) as numbers, and
    batch_key (E) exactly as in the master database. Read-only.
    """
    report = LedgerCheck()
    wb = load_workbook(path)
    try:
        _check_structure(wb, path)
    except LedgerStructureError as e:
        report.errors.append(str(e))
        return report
    ws = wb[LEDGER_TX_SHEET]
    headers = [c.value.strip() if isinstance(c.value, str) else None for c in ws[1]]

    rows: list[tuple[int, dict[str, Any]]] = []
    for r_idx, values in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        row = {h: v for h, v in zip(headers, values, strict=False) if h}
        if all(v is None or (isinstance(v, str) and not v.strip()) for v in row.values()):
            continue
        if not row.get("deal_agreement_num"):
            report.errors.append(f"строка {r_idx}: нет deal_agreement_num (A) — при следующей записи "
                                 f"реестра строка пропадёт")
            continue
        rows.append((r_idx, row))

    used: dict[str, float] = {}
    cap: dict[str, list[tuple[int, float]]] = {}
    deals: dict[str, list[str]] = {}
    pairs: dict[tuple[str, str], list[int]] = {}
    for r_idx, row in rows:
        num = str(row["deal_agreement_num"]).strip()
        key = row.get("batch_key")
        if key is None or not str(key).strip():
            report.errors.append(f"строка {r_idx} ({num}): нет batch_key (E) — строка не попадёт в «Баланс»")
            continue
        if not isinstance(key, str):
            key = str(key)
        if key != key.strip():
            report.errors.append(f"строка {r_idx} ({num}): в batch_key «{key}» лишние пробелы — "
                                 f"«Баланс» посчитает это отдельной партией")
        if batch_keys is not None and key.strip() not in batch_keys:
            report.warnings.append(f"строка {r_idx} ({num}): партии «{key}» нет в базе (лист «Партии») — "
                                   f"проверьте написание ключа")
        use, problem = _kg(row.get("use_kg"))
        if problem:
            report.errors.append(f"строка {r_idx} ({num}, партия {key}): use_kg (D) {problem}")
        cap_raw = row.get("batch_quantity_kg")
        if cap_raw is not None and not (isinstance(cap_raw, str) and not cap_raw.strip()):
            cap_kg, problem = _kg(cap_raw)
            if cap_kg is None:
                report.errors.append(f"строка {r_idx} ({num}, партия {key}): batch_quantity_kg (L) {problem}")
            else:
                cap.setdefault(key, []).append((r_idx, cap_kg))
        used[key] = used.get(key, 0.0) + (use or 0.0)
        deals.setdefault(key, [])
        if num not in deals[key]:
            deals[key].append(num)
        pairs.setdefault((num, key), []).append(r_idx)

    for (num, key), where in pairs.items():
        if len(where) > 1:
            report.warnings.append(f"сделка {num}, партия {key}: строки {', '.join(map(str, where))} — "
                                   f"списание учтено {len(where)} раза; оставьте одну строку")
    for key in used:
        weights = cap.get(key, [])
        if not weights:
            report.errors.append(f"партия {key}: нет batch_quantity_kg (L) ни в одной строке — "
                                 f"остаток не посчитается")
            continue
        if len({w for _, w in weights}) > 1:
            listed = ", ".join(f"строка {r}: {w:g}" for r, w in weights)
            report.warnings.append(f"партия {key}: количество партии различается ({listed}) — «Баланс» берёт последний")
        capacity = weights[-1][1]
        if used[key] > capacity:
            report.warnings.append(f"партия {key}: всего {math.ceil(capacity)} кг, списано {math.ceil(used[key])} кг — "
                                   f"ПЕРЕРАСХОД {math.ceil(used[key] - capacity)} кг")
        report.balance.append({"key": key, "capacity": capacity, "used": used[key],
                               "remaining": capacity - used[key], "deals": deals[key]})
    return report


def rebuild_ledger(path: Path, master_db_path: Path | None = None) -> tuple[int, list[str]]:
    """Rewrite the ledger as DCA writes it: same rows, fresh «Баланс» (every batch,
    hand-typed ones and the database's unused ones too). The previous file goes to
    _ledger_versions/. → (batch count in «Баланс», warnings)."""
    if master_db_path is None:
        master_db_path = path.parent / MASTER_DB_FILENAME
    lock_path = path.with_suffix(path.suffix + ".lock")
    try:
        with FileLock(str(lock_path), timeout=_LEDGER_LOCK_TIMEOUT_SEC):
            rows, extras = _read_tx_rows(path)
            db_records, warnings = _db_packages(master_db_path)
            batches = _save_workbook(path, rows, extras=extras, db_records=db_records)
    except Timeout as e:
        raise PermissionError(f"реестр занят другой командой dca ({lock_path.name}) — повторите позже") from e
    return batches, warnings
