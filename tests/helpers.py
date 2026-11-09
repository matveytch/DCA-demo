"""Shared helpers for the tests: the demo Excel files, deals built from them, Excel edits.

All data is the invented demo data of scripts/make_demo_templates.py.
"""

from __future__ import annotations

import io
import shutil
from pathlib import Path
from typing import Any

from docx import Document
from openpyxl import load_workbook

from dca.config import DATA_DIR
from dca.excel_io import load_deal
from dca.models import Deal

FILLED = DATA_DIR / "input_example_FILLED.xlsx"
EMPTY = DATA_DIR / "input_template_EMPTY.xlsx"
NUM = "1-2026-1705"            # agreement number of the example
COMPANY = "Компания А"         # party A's company folder in the example


def example_deal() -> Deal:
    """The filled example: outgoing, variant B, 2 annexes (600 + 400), 2 batches (700 + 300)."""
    return load_deal(FILLED)


def deal(**changes: Any) -> Deal:
    """A simple outgoing deal: one annex, one batch covering the whole quantity."""
    data = example_deal().model_dump()
    q = changes.get("quantity_kg", data["quantity_kg"])
    base = {"annexes": [], "batch_key_2": None, "batch_use_kg_2": None, "batch_use_kg_1": q}
    for key in [k for k in data if k.endswith("_2") and k.startswith("batch_")]:
        base[key] = None
    return Deal(**{**data, **base, **changes})


def incoming_deal(**changes: Any) -> Deal:
    """The example as an incoming deal (no variant, no batches)."""
    data = example_deal().model_dump()
    for key in [k for k in data if k.startswith("batch_")]:
        data[key] = None
    data.update(deal_kind="incoming", variant=None, party_c_name=None, party_c_address=None,
                party_c_details=None, party_b_key="alpha-b", party_b_folder_name=COMPANY)
    return Deal(**{**data, **changes})


def set_deal_values(xlsx: Path, values: dict[str, Any], sheet: str = "Сделка") -> None:
    """Write values into column C of a 4-column form sheet, by the name in column A."""
    wb = load_workbook(xlsx)
    ws = wb[sheet]
    seen = set()
    for row in ws.iter_rows():
        name = row[0].value
        if isinstance(name, str) and name.strip() in values:
            row[2].value = values[name.strip()]
            seen.add(name.strip())
    missing = set(values) - seen
    assert not missing, f"no rows {missing} on «{sheet}»"
    wb.save(xlsx)


def set_annexes(xlsx: Path, rows: list[tuple[Any, Any, Any]]) -> None:
    """Replace the data rows of «Приложения» (annex_num, annex_date, quantity_kg)."""
    wb = load_workbook(xlsx)
    ws = wb["Приложения"]
    for r in range(2, max(ws.max_row, 2) + 1):
        for c in (1, 2, 3):
            ws.cell(r, c).value = None
    for r, values in enumerate(rows, start=2):
        for c, v in enumerate(values, start=1):
            ws.cell(r, c).value = v
    wb.save(xlsx)


def set_archive_value(xlsx: Path, sheet: str, field: str, key: str, value: Any) -> None:
    """Set `field` of the archive record `key` (rows = fields, columns C… = records)."""
    wb = load_workbook(xlsx)
    ws = wb[sheet]
    rows = {ws.cell(r, 1).value: r for r in range(2, ws.max_row + 1)}
    col = next(c for c in range(3, ws.max_column + 1) if str(ws.cell(2, c).value) == key)
    ws.cell(rows[field], col).value = value
    wb.save(xlsx)


def copy_example(deal_dir: Path, num: str = NUM) -> Path:
    deal_dir.mkdir(parents=True, exist_ok=True)
    target = deal_dir / f"input_{num}.xlsx"
    shutil.copy(FILLED, target)
    return target


def docx_text(data: bytes | Path) -> str:
    """All text of a .docx: paragraphs, then table cells row by row."""
    doc = Document(io.BytesIO(data) if isinstance(data, bytes) else str(data))
    parts = [p.text for p in doc.paragraphs]
    for t in doc.tables:
        for row in t.rows:
            parts.append(" | ".join(c.text for c in row.cells))
    return "\n".join(parts)
