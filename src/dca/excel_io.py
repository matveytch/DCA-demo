"""Excel I/O — read the deal's input .xlsx and build a Deal.

Workbook structure:
  - «Сделка»      — deal-specific fields + archive keys (party A / party B / item / batches)
  - «Стороны А»   — party A archive (columns = records, rows = party_a_* fields)
  - «Стороны Б»   — party B archive (same structure)
  - «Изделия»     — items archive (rows = item_* fields)
  - «Партии»      — source batches archive (rows = batch_* fields)
  - «Приложения»  — optional table: several annexes of one agreement
  - «Передача»    — optional form: the transfer act
  - «Изменение»   — optional form: the amendment (`dca amend`)
  - «Превью»      — read-only formula-driven view (ignored by engine)

Flow:
  1. Read «Сделка» → all deal fields + archive keys
  2. Look up party_a_key / party_b_key / item_key in their archive sheets
  3. Look up batch_key_N in «Партии» → batch fields with suffix _N
  4. Read the optional sheets, merge into one flat dict and instantiate Deal

If an archive key is not found → raise with a clear error listing available keys.
"""

from __future__ import annotations

import re
from datetime import date, datetime
from pathlib import Path
from typing import Any

from dateutil.parser import parse as dateparse
from openpyxl import load_workbook
from pydantic import BaseModel, ValidationError

from dca.models import AmendmentInfo, Deal, TransferInfo

# Sheet names (Russian — matches the template)
DEAL_SHEET = "Сделка"
PARTY_A_SHEET = "Стороны А"
PARTY_B_SHEET = "Стороны Б"
ITEMS_SHEET = "Изделия"
BATCHES_SHEET = "Партии"        # source batches (looked up by batch_key_N)
ANNEXES_SHEET = "Приложения"    # several annexes per agreement (table, rows = annexes)
TRANSFER_SHEET = "Передача"     # transfer act (same 4-column layout as «Сделка»)
AMENDMENT_SHEET = "Изменение"   # amendment (`dca amend`), same layout
PREVIEW_SHEET = "Превью"        # ignored by engine

# Columns of the «Приложения» sheet (matched by header name)
_ANNEX_COLUMNS = ("annex_num", "annex_date", "quantity_kg")

# Archive key fields (special — used to look up the rest of the archive record)
PARTY_A_KEY_FIELD = "party_a_key"
PARTY_B_KEY_FIELD = "party_b_key"
ITEM_KEY_FIELD = "item_key"

# Date field names
_DATE_FIELDS = {"agreement_date", "annex_date", "due_date", "letter_date"}
for _n in (1, 2, 3, 4, 5):
    _DATE_FIELDS.update({f"batch_agreement_date_{_n}", f"batch_annex_date_{_n}"})
del _n

_SECTION_MARKER = "━"
_DATE_RE = re.compile(r"^\s*\d{1,2}[./\-]\d{1,2}[./\-]\d{2,4}\s*$")


# ─── Public API ─────────────────────────────────────────────────

def load_deal(path: Path | str, deal_kind: str | None = None) -> Deal:
    """Read multi-sheet input.xlsx, resolve archives, return a validated Deal.

    `deal_kind`, when given (e.g. "incoming" from `dca generate-incoming`), is injected
    into the deal data and overrides any value read from the sheet — the command
    that loads the deal knows its kind. When None, the kind comes from the sheet
    (if a `deal_kind` cell exists) or defaults to "outgoing" via the model.
    """
    path = Path(path)
    wb = load_workbook(path, data_only=True)

    if DEAL_SHEET not in wb.sheetnames:
        raise ValueError(
            f"Required sheet «{DEAL_SHEET}» not found in {path.name}. "
            f"Found sheets: {wb.sheetnames}"
        )

    # 1. Read Сделка
    _check_deal_sheet_names(wb[DEAL_SHEET])
    deal_raw = _read_main_sheet(wb[DEAL_SHEET])

    # 2. Extract archive keys (must be present)
    party_a_key = deal_raw.pop(PARTY_A_KEY_FIELD, None)
    party_b_key = deal_raw.pop(PARTY_B_KEY_FIELD, None)
    item_key = deal_raw.pop(ITEM_KEY_FIELD, None)

    # 3. Resolve archives
    if party_a_key:
        if PARTY_A_SHEET not in wb.sheetnames:
            raise ValueError(
                f"party_a_key='{party_a_key}' is set, but sheet «{PARTY_A_SHEET}» not found."
            )
        party_a_data = _lookup_in_archive(wb[PARTY_A_SHEET], party_a_key, PARTY_A_KEY_FIELD)
        # Archive data is authoritative — overrides anything in Сделка
        deal_raw.update(party_a_data)
    else:
        raise ValueError(
            f"party_a_key is required (pick party A from «{PARTY_A_SHEET}» dropdown)."
        )

    if party_b_key:
        if PARTY_B_SHEET not in wb.sheetnames:
            raise ValueError(
                f"party_b_key='{party_b_key}' is set, but sheet «{PARTY_B_SHEET}» not found."
            )
        party_b_data = _lookup_in_archive(wb[PARTY_B_SHEET], party_b_key, PARTY_B_KEY_FIELD)
        deal_raw.update(party_b_data)
    else:
        raise ValueError(
            f"party_b_key is required (pick party B from «{PARTY_B_SHEET}» dropdown)."
        )

    if item_key:
        if ITEMS_SHEET not in wb.sheetnames:
            raise ValueError(
                f"item_key='{item_key}' is set, but sheet «{ITEMS_SHEET}» not found."
            )
        item_data = _lookup_in_archive(wb[ITEMS_SHEET], item_key, ITEM_KEY_FIELD)
        deal_raw.update(item_data)
    else:
        raise ValueError(
            f"item_key is required (pick an item from «{ITEMS_SHEET}» dropdown)."
        )
    # Keep the chosen party keys: they file the deal into the company folder and
    # are compared with each batch's company (dca.checks).
    deal_raw[PARTY_A_KEY_FIELD] = party_a_key
    deal_raw[PARTY_B_KEY_FIELD] = party_b_key
    deal_raw[ITEM_KEY_FIELD] = item_key

    # 4. Batches — for each of 5 slots, lookup by batch_key in «Партии»
    # Each slot's atomic fields (agreement_num, date, etc.) get suffix `_N`.
    # weight_kg stays as deal_raw[batch_use_kg_N] (deal-specific).
    # Note: we KEEP `batch_key_N` in deal_raw (not pop) so that
    # downstream consumers (ledger, validators) know which archive entry
    # backs each slot.
    for pkg_n in (1, 2, 3, 4, 5):
        key_field = f"batch_key_{pkg_n}"
        pkg_key = deal_raw.get(key_field)
        if not pkg_key:
            continue  # slot empty — skip
        if BATCHES_SHEET not in wb.sheetnames:
            raise ValueError(
                f"{key_field}='{pkg_key}' is set, but sheet «{BATCHES_SHEET}» not found."
            )
        pkg_data = _lookup_in_archive(wb[BATCHES_SHEET], pkg_key, "batch_key")
        # Suffix each archive field with _N for this slot
        for field, value in pkg_data.items():
            deal_raw[f"{field}_{pkg_n}"] = value

    # 4b. Annexes — optional «Приложения» sheet (table, one row per annex).
    # When present and non-empty, populates deal.annexes (multi-annex deal).
    # When absent/empty, the deal has one annex from the deal-level fields.
    if ANNEXES_SHEET in wb.sheetnames:
        annexes = _read_annexes(wb[ANNEXES_SHEET])
        if annexes:
            deal_raw["annexes"] = annexes  # list[dict]; AnnexLine built by Deal

    # 4c. Transfer act — optional «Передача» sheet.
    if TRANSFER_SHEET in wb.sheetnames:
        transfer = read_transfer_sheet(wb[TRANSFER_SHEET])
        if transfer:
            deal_raw["transfer"] = transfer

    # 4d. Amendmentary agreement — optional «Изменение» sheet (`dca amendment`).
    if AMENDMENT_SHEET in wb.sheetnames:
        amendment = read_amendment_sheet(wb[AMENDMENT_SHEET])
        if amendment:
            deal_raw["amendment"] = amendment

    # 5. Type coerce and validate
    cleaned = _coerce_types(deal_raw)
    if deal_kind is not None:
        cleaned["deal_kind"] = deal_kind
    try:
        return Deal(**cleaned)
    except ValidationError as e:
        # Pydantic's «Field required [type=missing, input_value={…}]» says nothing to
        # the operator: name the field, its description and where to fill it.
        raise DealInputError(explain_errors(e, wb, party_a_key, party_b_key, item_key)) from None


class DealInputError(ValueError):
    """The input Excel can't make a valid deal; the message lists every problem in Russian."""


def explain_errors(e: ValidationError, wb: Any, party_a_key: Any = None, party_b_key: Any = None,
                   item_key: Any = None) -> str:
    """One Russian line per problem: what is wrong, which field, which sheet and row."""
    places = _field_places(wb)
    archives = {"party_a_": (PARTY_A_SHEET, party_a_key), "party_b_": (PARTY_B_SHEET, party_b_key),
                "item_": (ITEMS_SHEET, item_key)}
    lines = []
    for err in e.errors():
        loc = [str(x) for x in err["loc"]]
        kind = err["type"]
        msg = str(err.get("msg", "")).removeprefix("Value error, ")
        what = "не заполнено" if kind == "missing" else f"неверное значение — {msg}"
        if not loc:                                   # a check across several fields
            lines.append(f"• {msg}")
            continue
        if loc[0] in ("transfer", "amendment") and len(loc) > 1:
            name = loc[1]
            where = places.get((TRANSFER_SHEET if loc[0] == "transfer" else AMENDMENT_SHEET, name))
        elif loc[0] == "annexes" and len(loc) > 2:
            name = loc[2]
            where = (ANNEXES_SHEET, f"строка приложения №{int(loc[1]) + 1}", "")
        else:
            name = loc[0]
            where = places.get((DEAL_SHEET, name))
            if where is None:
                m = re.match(r"(batch_.+)_(\d)$", name)
                if m:
                    where = (BATCHES_SHEET, f"партия №{m.group(2)}", "")
                for prefix, (sheet, key) in archives.items():
                    if where is None and name.startswith(prefix):
                        where = (sheet, f"запись «{key}»" if key else "выбранная запись", "")
        if where is None:
            lines.append(f"• {name}: {what}")
        else:
            sheet, row, desc = where
            label = f"«{desc}» ({name})" if desc else name
            lines.append(f"• {label}: {what} — лист «{sheet}», {row}")
    head = "в Excel не хватает данных или они неверные:" if lines else str(e)
    return "\n".join([head, *lines])


def _field_places(wb: Any) -> dict[tuple[str, str], tuple[str, str, str]]:
    """(sheet, field) → (sheet, «строка N», description) for the «Сделка», «Передача» and
    «Изменение» rows."""
    out = {}
    for sheet in (DEAL_SHEET, TRANSFER_SHEET, AMENDMENT_SHEET):
        if sheet not in wb.sheetnames:
            continue
        for r, row in enumerate(wb[sheet].iter_rows(min_row=1, max_col=2, values_only=True), start=1):
            name, desc = (row + (None, None))[:2]
            if isinstance(name, str) and name.strip():
                out[(sheet, name.strip())] = (sheet, f"строка {r}", str(desc or "").split("\n")[0].strip())
    return out


def read_excel(path: Path | str) -> dict[str, Any]:
    """Read the main «Сделка» sheet → raw `{var_name: value}` dict (no archive lookup)."""
    path = Path(path)
    wb = load_workbook(path, data_only=True)
    if DEAL_SHEET not in wb.sheetnames:
        raise ValueError(
            f"Required sheet «{DEAL_SHEET}» not found in {path.name}. "
            f"Found sheets: {wb.sheetnames}"
        )
    return _read_main_sheet(wb[DEAL_SHEET])


# ─── Sheet readers ──────────────────────────────────────────────

# Fields the command decides, never the operator: a «deal_kind» row turned
# `dca generate` into an incoming deal.
SERVICE_FIELDS = frozenset({"deal_kind", "annexes", "transfer", "amendment"})
# Anything in column A that looks like a machine name — any letter case, spaces
# allowed — is checked, filled or not («Annex_date» must not slip through).
_NAME_LIKE_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_ ]*$")


def _check_deal_sheet_names(ws) -> None:
    """Column A of «Сделка» must hold the template's exact machine names.

    A renamed or mistyped name would lose its value silently and a default would
    be used instead.
    """
    allowed = set(Deal.model_fields) - SERVICE_FIELDS
    by_fold = {re.sub(r"[\s_]+", "_", f).casefold(): f for f in allowed}
    unknown, service = [], []
    for (cell,) in ws.iter_rows(min_row=1, max_col=1, values_only=True):
        if not isinstance(cell, str):
            continue
        name = cell.strip()
        if not name or name == "Variable name" or not _NAME_LIKE_RE.match(name) or name in allowed:
            continue
        if name in SERVICE_FIELDS:
            service.append(name)
            continue
        guess = by_fold.get(re.sub(r"[\s_]+", "_", name).casefold())
        unknown.append(f"{name} (should be {guess})" if guess else name)
    if service:
        raise ValueError(
            f"Sheet «{DEAL_SHEET}» has the service field(s) {', '.join(service)} — the command decides "
            f"them (dca generate = outgoing, dca generate-incoming = incoming). Delete the row."
        )
    if unknown:
        raise ValueError(
            f"Unknown variable(s) on sheet «{DEAL_SHEET}»: {', '.join(unknown)} — column A must keep "
            f"the template's machine names (a renamed or mistyped name loses its value)."
        )


def _read_main_sheet(ws) -> dict[str, Any]:
    """Read the «Сделка» sheet (4-column layout) → {var_name: value} dict."""
    result: dict[str, Any] = {}
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True):
        if len(row) < 3:
            continue
        var_name, _, value, *_rest = row

        if not isinstance(var_name, str):
            continue
        var_name = var_name.strip()
        if not var_name:
            continue
        if _SECTION_MARKER in var_name:
            continue
        if var_name in ("Variable name", "ЛЕГЕНДА:"):
            continue

        if value is None:
            continue
        if isinstance(value, str) and not value.strip():
            continue

        result[var_name] = value
    return result


def _lookup_in_archive(ws, key_value: str, key_field: str) -> dict[str, Any]:
    """Find a column in archive sheet by `key_value` matching `key_field` row.

    Archive layout:
      Row 1: header (Variable name | Описание | Slot 1 | Slot 2 | ... | Slot N)
      Row 2: key_field row (party_a_key | (desc) | "alpha" | "beta" | ...)
      Rows 3+: other variable rows

    Returns dict of all variables for the matching column.
    Raises ValueError if key_value not found.
    """
    # Find row with our key_field in column A
    key_row_idx = None
    for r in range(1, ws.max_row + 1):
        a = ws.cell(row=r, column=1).value
        if isinstance(a, str) and a.strip() == key_field:
            key_row_idx = r
            break
    if key_row_idx is None:
        raise ValueError(
            f"Archive sheet does not have a row with '{key_field}' in column A."
        )

    # Find column where row[key_row_idx] == key_value
    target_col = None
    available_keys: list[str] = []
    for c in range(3, ws.max_column + 1):
        cell_val = ws.cell(row=key_row_idx, column=c).value
        if cell_val is None:
            continue
        cell_str = str(cell_val).strip()
        if not cell_str:
            continue
        available_keys.append(cell_str)
        if cell_str == key_value:
            target_col = c
            break

    if target_col is None:
        raise ValueError(
            f"{key_field}='{key_value}' not found in archive. "
            f"Available keys: {available_keys}"
        )

    # Pull all variable rows for this column
    result: dict[str, Any] = {}
    for r in range(2, ws.max_row + 1):
        var_name = ws.cell(row=r, column=1).value
        if not isinstance(var_name, str):
            continue
        var_name = var_name.strip()
        if not var_name:
            continue
        if var_name == key_field:
            continue  # don't include the key itself in the deal data

        value = ws.cell(row=r, column=target_col).value
        if value is None:
            continue
        if isinstance(value, str) and not value.strip():
            continue
        result[var_name] = value
    return result


def _read_annexes(ws) -> list[dict[str, Any]]:
    """Read the «Приложения» table → list of annex dicts (one per data row).

    Row 1 holds column headers matched by name to `_ANNEX_COLUMNS`
    (annex_num | annex_date | quantity_kg); column order is
    flexible. Data starts at row 2. Fully-empty rows are skipped. Dates are
    parsed via `_parse_date`, weight → float, number → str.
    """
    # Find the header row: the first row (within the first 10) whose cells match
    # the machine field names. Allows an optional human title/hint row above it.
    header_row = None
    col_of: dict[str, int] = {}
    for r in range(1, min(ws.max_row, 10) + 1):
        mapping = {}
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and v.strip() in _ANNEX_COLUMNS:
                mapping[v.strip()] = c
        if "annex_num" in mapping and "quantity_kg" in mapping:
            header_row, col_of = r, mapping
            break
    if not col_of or header_row is None:
        # No recognisable header: fine for an untouched sheet, but a FILLED sheet
        # whose header was changed must not silently become «single annex».
        # From row 2: a single data row under a broken row-1 header counts too.
        filled = [ws.cell(row=r, column=c).value for r in range(2, ws.max_row + 1)
                  for c in range(1, ws.max_column + 1)]
        if any(v not in (None, "") for v in filled):
            raise ValueError(
                f"Sheet «{ANNEXES_SHEET}» has data but no header row with annex_num / "
                f"annex_date / quantity_kg — restore the header names."
            )
        return []

    annexes: list[dict[str, Any]] = []
    for r in range(header_row + 1, ws.max_row + 1):
        raw = {field: ws.cell(row=r, column=col).value for field, col in col_of.items()}
        # Skip blank rows (all cells empty / whitespace).
        if all(v is None or (isinstance(v, str) and not v.strip()) for v in raw.values()):
            continue
        annex: dict[str, Any] = {}
        num = raw.get("annex_num")
        if num is not None:
            annex["annex_num"] = str(int(num)) if isinstance(num, float) and num == int(num) else str(num).strip()
        d = raw.get("annex_date")
        if d is not None and not (isinstance(d, str) and not d.strip()):
            annex["annex_date"] = _parse_date(d, "annex_date")
        w = raw.get("quantity_kg")
        if w is not None and not (isinstance(w, str) and not str(w).strip()):
            annex["quantity_kg"] = float(w)
        annexes.append(annex)
    return annexes


_IDENTIFIER_RE = re.compile(r"^[a-z][a-z0-9_]*$")


def read_transfer_sheet(ws: Any) -> dict[str, Any]:
    """Read the «Передача» sheet → {field: value} (see read_form_sheet)."""
    return read_form_sheet(ws, TransferInfo, TRANSFER_SHEET)


def read_amendment_sheet(ws: Any) -> dict[str, Any]:
    """Read the «Изменение» sheet → {field: value} (see read_form_sheet)."""
    return read_form_sheet(ws, AmendmentInfo, AMENDMENT_SHEET)


def read_form_sheet(ws: Any, model: type[BaseModel], sheet_name: str) -> dict[str, Any]:
    """Read a 4-column form sheet (like «Сделка») → {field: value} for `model`.

    Only rows whose column A is a snake_case identifier are data rows; the
    read-only reference rows at the top (formulas pulled from «Сделка») have a
    human label there and are skipped. An identifier that is not a field of
    `model` is an error — a typo in a variable name must not silently drop data.
    """
    known = set(model.model_fields)
    date_fields = {n for n, f in model.model_fields.items() if f.annotation in (date, date | None)}
    float_fields = {n for n, f in model.model_fields.items() if f.annotation in (float, float | None)}
    result: dict[str, Any] = {}
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True):
        if len(row) < 3 or not isinstance(row[0], str):
            continue
        var_name = row[0].strip()
        if not _IDENTIFIER_RE.match(var_name):
            continue
        if var_name not in known:
            raise ValueError(
                f"Unknown variable '{var_name}' on sheet «{sheet_name}» "
                f"(column A must hold one of the sheet's field names)."
            )
        value = row[2]
        if value is None or (isinstance(value, str) and not value.strip()):
            continue
        if var_name in date_fields:
            value = _parse_date(value, var_name)
        elif var_name in float_fields:
            value = _parse_number(value, var_name, sheet_name)
        elif not isinstance(value, str):
            value = str(int(value)) if isinstance(value, float) and value == int(value) else str(value)
        result[var_name] = value.strip() if isinstance(value, str) else value
    return result


def _parse_number(value: Any, var_name: str, sheet_name: str) -> float:
    """A number cell; text like «1 200,5» (typed into a text cell) is accepted too."""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    text = str(value).replace(" ", "").replace(" ", "").replace(",", ".")
    try:
        return float(text)
    except ValueError:
        raise ValueError(f"{var_name} on sheet «{sheet_name}»: «{value}» is not a number") from None


# ─── Type coercion ──────────────────────────────────────────────


def _is_string_field(field) -> bool:
    """True if Pydantic field's annotation is `str` or `str | None`."""
    from typing import get_args
    ann = field.annotation
    if ann is str:
        return True
    args = get_args(ann)  # for Optional[str] = Union[str, None]
    return str in args


# Build once at import: set of field names whose target type is str.
# openpyxl returns int/float for numeric cells regardless of cell format,
# so we coerce them to str for these fields.
_STRING_FIELDS = {
    name for name, f in Deal.model_fields.items()
    if _is_string_field(f) and name not in _DATE_FIELDS
}


def _coerce_types(raw: dict[str, Any]) -> dict[str, Any]:
    """Convert Excel values to types expected by Deal model."""
    known_fields = set(Deal.model_fields.keys())
    out: dict[str, Any] = {}
    for key, val in raw.items():
        if key not in known_fields:
            continue

        if key in _DATE_FIELDS:
            out[key] = _parse_date(val, key)
            continue

        # If field expects str but openpyxl returned a number (e.g. "1811" entered
        # in a cell formatted as Text, but openpyxl still reads it as int 1811),
        # coerce: int → "1811", float-that-is-integer → "1811", float → "1811.5"
        if key in _STRING_FIELDS and not isinstance(val, str):
            if isinstance(val, float) and val == int(val):
                val = str(int(val))   # 1811.0 → "1811" (no trailing ".0")
            else:
                val = str(val)        # 1811 → "1811";  1811.5 → "1811.5"

        if isinstance(val, str):
            val = val.strip()
            if not val:
                continue

        out[key] = val
    return out


def _parse_date(val: Any, field: str) -> date:
    """Convert any reasonable date representation to a `date`."""
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    if isinstance(val, str):
        s = val.strip()
        if _DATE_RE.match(s):
            try:
                return dateparse(s, dayfirst=True).date()
            except (ValueError, OverflowError) as e:
                raise ValueError(
                    f"Cannot parse date for field '{field}': {val!r} — {e}"
                ) from e
        try:
            return date.fromisoformat(s)
        except ValueError:
            pass
    raise ValueError(
        f"Cannot parse date for field '{field}': {val!r} "
        f"(expected dd.mm.yyyy or ISO format, got {type(val).__name__})"
    )
