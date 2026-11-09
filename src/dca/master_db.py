"""The master database of parties / items / source batches (`Data/master_db.xlsx`).

One Excel file in the DATA root (operator data, gitignored) with the same four
archive sheets as the deal input — «Стороны А», «Стороны Б», «Изделия», «Партии» —
in the same layout: rows = fields (column A), columns C… = records, key in row 2.

`dca new` / `dca new-incoming` copy the input template and write a SNAPSHOT of
every record of this file into the new deal's archive sheets; `dca refresh-db`
re-snapshots an existing deal. The deal keeps its own copy (traceability: an old
deal is not changed by later edits of the database unless refreshed on purpose).

The field list is taken from the input template's archive sheets — the template
is the source of truth, so a field added there is expected in the database too.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from dca.config import EMPTY_TEMPLATE
from dca.excel_io import (
    BATCHES_SHEET,
    ITEM_KEY_FIELD,
    ITEMS_SHEET,
    PARTY_A_KEY_FIELD,
    PARTY_A_SHEET,
    PARTY_B_KEY_FIELD,
    PARTY_B_SHEET,
)
from dca.utils.atomic import refuse_if_open_in_excel, save_workbook_atomic

BATCH_KEY_FIELD = "batch_key"
ARCHIVES: dict[str, str] = {           # sheet → key field
    PARTY_A_SHEET: PARTY_A_KEY_FIELD,
    PARTY_B_SHEET: PARTY_B_KEY_FIELD,
    ITEMS_SHEET: ITEM_KEY_FIELD,
    BATCHES_SHEET: BATCH_KEY_FIELD,
}
BATCH_COMPANY_FIELD = "batch_company_key"   # dropdown of our companies (party A keys)
FIRST_RECORD_COL = 3                   # column C
INIT_SLOTS = 30                        # prepared (styled) record columns in a new database

Record = dict[str, Any]


@dataclass
class ArchiveData:
    records: dict[str, list[Record]]                  # sheet → records in column order
    warnings: list[str] = field(default_factory=list)

    def keys(self, sheet: str) -> list[str]:
        return [str(r[ARCHIVES[sheet]]) for r in self.records[sheet]]

    def count_line(self) -> str:
        return ", ".join(f"{sheet} {len(self.records[sheet])}" for sheet in ARCHIVES)


def template_layout(template: Path = EMPTY_TEMPLATE) -> dict[str, list[tuple[str, str, str]]]:
    """sheet → [(field, description, record header prefix)] from the input template."""
    wb = load_workbook(template, read_only=True)
    layout = {}
    for sheet in ARCHIVES:
        rows = list(wb[sheet].iter_rows(min_row=1, max_col=3, values_only=True))
        header = str(rows[0][2] or "Запись #1")
        prefix = header.rsplit("#", 1)[0].strip() or "Запись"
        layout[sheet] = [
            (str(a).strip(), str(b or "").strip(), prefix)
            for a, b, _ in rows[1:] if isinstance(a, str) and a.strip()
        ]
    wb.close()
    return layout


MASTER_VERSIONS_DIR = "_master_db_versions"


def init_master(path: Path) -> None:
    """Create an empty database with the template's fields. Refuses to overwrite."""
    if path.exists():
        raise FileExistsError(f"{path} already exists — edit it in Excel instead")
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="4472C4")
    key_fill = PatternFill("solid", fgColor="FFF2CC")
    wrap = Alignment(wrap_text=True, vertical="top")

    wb = Workbook()
    wb.remove(wb.active)
    for sheet, fields in template_layout().items():
        ws = wb.create_sheet(sheet)
        prefix = fields[0][2]
        heads = ["Variable name", "Описание", *(f"{prefix} #{i}" for i in range(1, INIT_SLOTS + 1))]
        for c, text in enumerate(heads, start=1):
            cell = ws.cell(row=1, column=c, value=text)
            cell.font, cell.fill = header_font, header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
        for r, (name, desc, _) in enumerate(fields, start=2):
            ws.cell(row=r, column=1, value=name).font = Font(bold=True)
            ws.cell(row=r, column=2, value=desc).alignment = wrap
            for c in range(FIRST_RECORD_COL, FIRST_RECORD_COL + INIT_SLOTS):
                cell = ws.cell(row=r, column=c)
                cell.alignment = wrap
                if r == 2:
                    cell.fill = key_fill
        ws.column_dimensions["A"].width = 38
        ws.column_dimensions["B"].width = 44
        for c in range(FIRST_RECORD_COL, FIRST_RECORD_COL + INIT_SLOTS):
            ws.column_dimensions[get_column_letter(c)].width = 35
        ws.freeze_panes = "C2"
        _add_company_dropdown(ws)
    save_workbook_atomic(wb, path, required_sheets=list(ARCHIVES))


def upgrade_master(path: Path) -> dict[str, list[str]]:
    """Append template fields missing from an existing database (rows at the bottom).

    Returns sheet → added field names. Existing rows and records are untouched.
    """
    from filelock import FileLock

    refuse_if_open_in_excel(path)
    with FileLock(str(path.with_name(path.name + ".lock")), timeout=30):
        return _upgrade_locked(path)


def _upgrade_locked(path: Path) -> dict[str, list[str]]:
    wb = load_workbook(path)
    added: dict[str, list[str]] = {}
    for sheet, fields in template_layout().items():
        ws = wb[sheet]
        have = {str(ws.cell(r, 1).value).strip() for r in range(2, ws.max_row + 1) if ws.cell(r, 1).value}
        missing = [(f, d) for f, d, _ in fields if f not in have]
        for name, desc in missing:
            r = ws.max_row + 1
            ws.cell(row=r, column=1, value=name).font = Font(bold=True)
            ws.cell(row=r, column=2, value=desc).alignment = Alignment(wrap_text=True, vertical="top")
        if missing:
            added[sheet] = [f for f, _ in missing]
            if BATCH_COMPANY_FIELD in added[sheet]:
                _add_company_dropdown(ws)
    if added:
        # Never straight over the database: temp → check → previous
        # version into Data/_master_db_versions/ → replace.
        save_workbook_atomic(wb, path, versions_dir=path.parent / MASTER_VERSIONS_DIR,
                             required_sheets=list(ARCHIVES))
    return added


def _add_company_dropdown(ws: Any) -> None:
    """«Партии»: batch_company_key picks one of the «Стороны А» keys."""
    row = next((r for r in range(2, ws.max_row + 1) if ws.cell(r, 1).value == BATCH_COMPANY_FIELD), None)
    if row is None:
        return
    dv = DataValidation(type="list", formula1=f"='{PARTY_A_SHEET}'!$C$2:$ZZ$2", allow_blank=True)
    dv.add(f"C{row}:ZZ{row}")
    ws.add_data_validation(dv)


def read_archives(path: Path, *, strict: bool) -> ArchiveData:
    """Read the four archive sheets of `path` (the database, or a deal's input Excel).

    strict=True (the database): unknown field names and duplicate keys are errors —
    a typo must not silently drop data, and a key must point to exactly one record.
    """
    layout = template_layout()
    wb = load_workbook(path, data_only=True)
    errors: list[str] = []
    data = ArchiveData(records={})
    for sheet, key_field in ARCHIVES.items():
        if sheet not in wb.sheetnames:
            errors.append(f"sheet «{sheet}» is missing")
            data.records[sheet] = []
            continue
        ws = wb[sheet]
        expected = [f for f, _, _ in layout[sheet]]
        row_of: dict[str, int] = {}
        for r in range(2, ws.max_row + 1):
            name = ws.cell(row=r, column=1).value
            if isinstance(name, str) and name.strip():
                name = name.strip()
                if name not in expected:
                    if strict:
                        errors.append(f"«{sheet}»: unknown field '{name}' in A{r}")
                    continue
                row_of[name] = r
        if key_field not in row_of:
            errors.append(f"«{sheet}»: key row '{key_field}' not found in column A")
            data.records[sheet] = []
            continue
        missing = [f for f in expected if f not in row_of]
        if missing:
            data.warnings.append(f"«{sheet}»: fields not in the file (treated as empty): {', '.join(missing)}")

        records: list[Record] = []
        seen: dict[str, str] = {}
        for c in range(FIRST_RECORD_COL, ws.max_column + 1):
            key = _clean(ws.cell(row=row_of[key_field], column=c).value)
            if key in (None, ""):
                continue
            key = str(key)
            col = get_column_letter(c)
            if key in seen:
                errors.append(f"«{sheet}»: duplicate key '{key}' in columns {seen[key]} and {col}")
                continue
            seen[key] = col
            rec: Record = {key_field: key}
            for name, r in row_of.items():
                if name != key_field:
                    v = _clean(ws.cell(row=r, column=c).value)
                    if v not in (None, ""):
                        rec[name] = v
            records.append(rec)
        data.records[sheet] = records
    wb.close()
    if errors:
        raise ValueError(f"{path.name}:\n  - " + "\n  - ".join(errors))
    return data


def diff(old: ArchiveData, new: ArchiveData) -> dict[str, dict[str, list[str]]]:
    """Per sheet: added / removed keys and changed keys with the fields that differ."""
    out: dict[str, dict[str, list[str]]] = {}
    for sheet, key_field in ARCHIVES.items():
        before = {r[key_field]: r for r in old.records.get(sheet, [])}
        after = {r[key_field]: r for r in new.records.get(sheet, [])}
        changed = []
        for k in after.keys() & before.keys():
            fields = sorted(f for f in before[k].keys() | after[k].keys() if before[k].get(f) != after[k].get(f))
            if fields:
                changed.append(f"{k} ({', '.join(fields)})")
        out[sheet] = {
            "added": [k for k in after if k not in before],
            "removed": [k for k in before if k not in after],
            "changed": sorted(changed),
        }
    return out


def _clean(v: Any) -> Any:
    if isinstance(v, datetime):
        v = v.date()
    if isinstance(v, date):
        return v.strftime("%d.%m.%Y")
    if isinstance(v, float) and v == int(v):
        return int(v)
    if isinstance(v, str):
        return v.strip()
    return v
