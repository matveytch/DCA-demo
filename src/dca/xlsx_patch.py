"""Add / hide sheets in the input .xlsx by patching its XML inside the zip.

Never through openpyxl load→save: that destroys the extension dropdowns (party /
item / batch keys), conditional formatting and data validations of the
«Сделка» sheet. Here every other zip member is copied byte-for-byte.

Used for the «Передача» sheet (transfer act) and the «Изменение» sheet
(`dca amend`): added to the Excel templates once by
scripts/make_demo_templates.py, and on the fly to a deal whose Excel lacks them.
"""

from __future__ import annotations

import html
import re
import zipfile
from pathlib import Path
from typing import Any
from xml.sax.saxutils import escape

TRANSFER_SHEET = "Передача"
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# Styles are borrowed from cells of the SAME workbook's «Сделка» sheet (style indices
# differ between files). role → cell on «Сделка» whose style it copies.
_STYLE_CELLS = {
    "header": "A1", "legend": "A2", "legend_orange": "C2", "legend_green": "C3", "legend_pale": "C4",
    "section": "A5", "section_fill": "B5",
    "var": "A10", "desc": "B10", "required": "C10", "note": "D10", "optional": "C12",
}
_LEGEND = (  # (cell, style role, text) — same legend as «Сделка», inline (shared strings may be reordered)
    ("A2", "legend", "ЛЕГЕНДА:"),
    ("C2", "legend_orange", "оранжевый = обязательное и ПУСТО"),
    ("C3", "legend_green", "зелёный = обязательное и ЗАПОЛНЕНО"),
    ("C4", "legend_pale", "бледно-зелёный = опциональное"),
)

# Row kinds: ("ref", label, description, var_in_deal, is_date)
#            ("section", title)
#            ("field", var, description, note, required, text_format)
_REF = "ref"
_SECTION = "section"
_FIELD = "field"

TRANSFER_ROWS: list[tuple[Any, ...]] = [
    (_SECTION, "  ━━━ ИЗ СДЕЛКИ (только просмотр — подтягивается с листа «Сделка») ━━━"),
    (_REF, "↳ agreement_num", "Номер соглашения", "agreement_num", False),
    (_REF, "↳ agreement_date", "Дата соглашения", "agreement_date", True),
    (_REF, "↳ party_b_key", "Сторона Б", "party_b_key", False),
    (_REF, "↳ item_key", "Изделие", "item_key", False),
    (_REF, "↳ quantity_kg", "Количество, кг", "quantity_kg", False),
    (_REF, "↳ place_ru", "Место исполнения", "place_ru", False),

    (_SECTION, "  ━━━ АКТ ПЕРЕДАЧИ ━━━"),
    (_FIELD, "transfer_num", "Номер акта", "[обязательное · текст]\nНапример 00017 — ведущие нули сохраняются.", True, True),
    (_FIELD, "transfer_date", "Дата передачи", "[обязательное · дата]\nDD.MM.YYYY. Не раньше даты соглашения.", True, False),
    (_FIELD, "place_from_ru", "Откуда (графа 5)", "[обязательное]\nОдна строка в форме.", True, False),
    (_FIELD, "place_to_ru", "Куда (графа 6)", "[обязательное]\nОдна строка в форме.", True, False),

    (_SECTION, "  ━━━ ПРЕДСТАВИТЕЛИ ━━━"),
    (_FIELD, "handed_by_ru", "Передал (ФИО)", "[обязательное]\nПредставитель стороны А.", True, False),
    (_FIELD, "received_by_ru", "Принял (ФИО)", "[обязательное]\nПредставитель стороны Б.", True, False),
    (_FIELD, "note_ru", "Примечание (графа 4)", "[multi-line (Alt+Enter)] Опционально. До 3 строк в форме.", False, False),
]

AMENDMENT_SHEET = "Изменение"

# `dca amend`: the amendment changes the agreement's quantity and/or price;
# the engine computes both totals, words and clause numbers.
AMENDMENT_ROWS: list[tuple[Any, ...]] = [
    (_SECTION, "  ━━━ ИЗ СДЕЛКИ (только просмотр — подтягивается с листа «Сделка») ━━━"),
    (_REF, "↳ agreement_num", "Номер соглашения", "agreement_num", False),
    (_REF, "↳ agreement_date", "Дата соглашения", "agreement_date", True),
    (_REF, "↳ quantity_kg", "Количество по соглашению, кг", "quantity_kg", False),
    (_REF, "↳ price_per_unit", "Цена по соглашению, USD за 1 кг", "price_per_unit", False),

    (_SECTION, "  ━━━ ИЗМЕНЕНИЕ ━━━"),
    (_FIELD, "amendment_num", "Номер изменения", "[обязательное · текст] Например: 1", True, True),
    (_FIELD, "amendment_date", "Дата изменения", "[обязательное · дата] DD.MM.YYYY. Не раньше даты соглашения.", True, False),

    (_SECTION, "  ━━━ КОЛИЧЕСТВО ━━━"),
    (_FIELD, "old_quantity_kg", "Прежнее количество, кг", "[число] Если пусто → количество с листа «Сделка». Заполните для второго и следующих изменений.", False, False),
    (_FIELD, "new_quantity_kg", "Новое количество, кг", "[обязательное · число] Новое общее количество по соглашению.", True, False),

    (_SECTION, "  ━━━ ЦЕНА (USD за 1 кг) ━━━"),
    (_FIELD, "old_price_per_unit", "Прежняя цена, USD за 1 кг", "[число] Если пусто → цена с листа «Сделка».", False, False),
    (_FIELD, "new_price_per_unit", "Новая цена, USD за 1 кг", "[число] Если пусто → цена не меняется.", False, False),
]


def sheet_names(xlsx: Path) -> list[str]:
    with zipfile.ZipFile(xlsx) as z:
        wb = z.read("xl/workbook.xml").decode("utf-8")
    return [_unescape(n) for n in re.findall(r'<sheet\b[^>]*\bname="([^"]*)"', wb)]


def add_transfer_sheet(xlsx: Path, values: dict[str, Any] | None = None) -> None:
    """Insert the «Передача» sheet right after «Сделка». `values` pre-fills fields (examples)."""
    _add_form_sheet(xlsx, TRANSFER_SHEET, TRANSFER_ROWS, values, after=("Сделка",))


def add_amendment_sheet(xlsx: Path, values: dict[str, Any] | None = None) -> None:
    """Insert the «Изменение» sheet after «Передача» (or «Сделка»). `values` pre-fills fields."""
    _add_form_sheet(xlsx, AMENDMENT_SHEET, AMENDMENT_ROWS, values, after=(TRANSFER_SHEET, "Сделка"))


def _add_form_sheet(xlsx: Path, name: str, form_rows: list[tuple[Any, ...]],
                    values: dict[str, Any] | None, after: tuple[str, ...]) -> None:
    """Insert a 4-column form sheet (like «Сделка») after the first existing sheet of `after`."""
    if name in sheet_names(xlsx):
        raise ValueError(f"{xlsx.name} already has a «{name}» sheet")
    data, infos = _read(xlsx)

    st, orange_dxf = _deal_styles(data)
    styles = data["xl/styles.xml"].decode("utf-8")
    styles, (st["required_text"], st["optional_text"], st["ref_date"]) = _append_xfs(styles, [
        (st["required"], 49), (st["optional"], 49), (st["desc"], 14),   # 49 = text «@», 14 = date
    ])
    data["xl/styles.xml"] = styles.encode("utf-8")

    n = 1 + max(int(m) for m in re.findall(r"xl/worksheets/sheet(\d+)\.xml", "\n".join(data)))
    member = f"xl/worksheets/sheet{n}.xml"
    data[member] = _form_sheet_xml(form_rows, values or {}, st, orange_dxf).encode("utf-8")

    rels = data["xl/_rels/workbook.xml.rels"].decode("utf-8")
    rid_nums = [int(m) for m in re.findall(r'Id="rId(\d+)"', rels)]
    rid = f"rId{1 + max(rid_nums)}"
    rels = rels.replace(
        "</Relationships>",
        f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
        f'relationships/worksheet" Target="worksheets/sheet{n}.xml"/></Relationships>',
    )
    data["xl/_rels/workbook.xml.rels"] = rels.encode("utf-8")

    ct = data["[Content_Types].xml"].decode("utf-8")
    ct = ct.replace(
        "</Types>",
        f'<Override PartName="/{member}" ContentType="application/vnd.openxmlformats-'
        f'officedocument.spreadsheetml.worksheet+xml"/></Types>',
    )
    data["[Content_Types].xml"] = ct.encode("utf-8")

    wb = data["xl/workbook.xml"].decode("utf-8")
    sheet_id = 1 + max(int(m) for m in re.findall(r'<sheet\b[^>]*\bsheetId="(\d+)"', wb))
    anchor = next((m for m in (re.search(rf'<sheet\b[^>]*\bname="{re.escape(a)}"[^>]*/>', wb) for a in after)
                   if m), None)
    # Declare r: on the tag itself: openpyxl-written workbooks bind it per <sheet>, not on the root.
    new_tag = (f'<sheet xmlns:r="{_REL_NS}" name="{escape(name)}" '
               f'sheetId="{sheet_id}" r:id="{rid}"/>')
    if anchor:
        wb = wb[: anchor.end()] + new_tag + wb[anchor.end():]
    else:
        wb = wb.replace("</sheets>", new_tag + "</sheets>")
    # Reference rows are formulas without cached values → let Excel compute on open.
    wb = re.sub(r"<calcPr\b(?![^>]*fullCalcOnLoad)", '<calcPr fullCalcOnLoad="1"', wb, count=1)
    data["xl/workbook.xml"] = wb.encode("utf-8")

    _write(xlsx, data, infos, extra=(member,))


def hide_sheet(xlsx: Path, name: str) -> None:
    """Mark a sheet hidden (state="hidden") — e.g. «Передача» in incoming deals."""
    data, infos = _read(xlsx)
    wb = data["xl/workbook.xml"].decode("utf-8")
    tag = re.search(rf'<sheet\b[^>]*\bname="{re.escape(escape(name))}"[^>]*/>', wb)
    if tag is None or 'state="' in tag.group(0):
        return
    hidden = tag.group(0).replace("<sheet ", '<sheet state="hidden" ', 1)
    data["xl/workbook.xml"] = (wb[: tag.start()] + hidden + wb[tag.end():]).encode("utf-8")
    _write(xlsx, data, infos)



# Structure protection with the same password ("1") as the archive sheets. Value cells
# (column C) are unlocked, so typing works; rows, labels and formulas can't be broken.
SHEET_PROTECTION = '<sheetProtection password="CE28" sheet="1"/>'


def protect_sheet(xlsx: Path, name: str) -> bool:
    """Protect sheet `name`. False if it was protected already."""
    data, infos = _read(xlsx)
    member = _sheet_member(data, name)
    xml = data[member].decode("utf-8")
    if "<sheetProtection" in xml:
        return False
    data[member] = _with_protection(xml).encode("utf-8")
    _write(xlsx, data, infos)
    return True


def _with_protection(xml: str) -> str:
    # Schema order: sheetProtection follows sheetData (and sheetCalcPr, which we never write).
    end = xml.index("</sheetData>") + len("</sheetData>")
    calc = re.match(r"<sheetCalcPr\b[^>]*/>", xml[end:])
    if calc:
        end += calc.end()
    return xml[:end] + SHEET_PROTECTION + xml[end:]

# ─── internals ──────────────────────────────────────────────────


def _deal_styles(data: dict[str, bytes]) -> tuple[dict[str, int], int]:
    """Style index per role + the "required and empty" CF dxf, read from «Сделка»."""
    sheet = data[_sheet_member(data, "Сделка")].decode("utf-8")

    styles = {}
    for role, ref in _STYLE_CELLS.items():
        m = re.search(rf'<c r="{ref}"[^>]*?\bs="(\d+)"', sheet)
        styles[role] = int(m.group(1)) if m else 0
    dxf = 0
    for sqref, dxf_id in re.findall(r'<conditionalFormatting sqref="([^"]*)"><cfRule[^>]*\bdxfId="(\d+)"', sheet):
        if any(_in_range("C10", rng) for rng in sqref.split()):
            dxf = int(dxf_id)
            break
    return styles, dxf


def _sheet_member(data: dict[str, bytes], name: str) -> str:
    """Zip member (xl/worksheets/sheetN.xml) of the sheet called `name`, via workbook.xml + rels."""
    wb = data["xl/workbook.xml"].decode("utf-8")
    rels = data["xl/_rels/workbook.xml.rels"].decode("utf-8")
    tag = re.search(rf'<sheet\b[^>]*\bname="{re.escape(escape(name))}"[^>]*/>', wb)
    rid_m = re.search(r'\b\w+:id="([^"]+)"', tag.group(0)) if tag else None
    rid = rid_m.group(1) if rid_m else None
    rel = next((r for r in re.findall(r"<Relationship\b[^>]*>", rels) if f'Id="{rid}"' in r), None)
    target = re.search(r'\bTarget="([^"]+)"', rel) if rel else None
    if target is None:
        raise ValueError(f"«{name}» sheet not found in workbook.xml")
    member = target.group(1).lstrip("/")
    return member if member.startswith("xl/") else "xl/" + member


def _in_range(ref: str, rng: str) -> bool:
    col, row = ref[0], int(ref[1:])
    a, _, b = rng.partition(":")
    b = b or a
    return a[0] <= col <= b[0] and int(a[1:]) <= row <= int(b[1:])


def _form_sheet_xml(form_rows: list[tuple[Any, ...]], values: dict[str, Any], st: dict[str, int],
                    orange_dxf: int) -> str:
    rows: list[str] = []
    merges: list[str] = []
    required_cells: list[str] = []

    def cell(ref: str, style: int, text: str | None = None) -> str:
        if text is None:
            return f'<c r="{ref}" s="{style}"/>'
        return f'<c r="{ref}" s="{style}" t="inlineStr"><is><t xml:space="preserve">{escape(text)}</t></is></c>'

    def value_cell(ref: str, style: int, v: Any) -> str:
        if v is None or v == "":
            return f'<c r="{ref}" s="{style}"/>'
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            return f'<c r="{ref}" s="{style}"><v>{v}</v></c>'
        return cell(ref, style, str(v))

    rows.append('<row r="1" ht="24" customHeight="1">' + "".join(
        cell(f"{c}1", st["header"], t) for c, t in zip("ABCD", ("Variable name", "Описание (что вводить)", "Value", "Notes / Help"), strict=True)
    ) + "</row>")
    for r in (2, 3, 4):
        cs = "".join(cell(ref, st[role], text) for ref, role, text in _LEGEND if ref[1:] == str(r))
        rows.append(f'<row r="{r}">{cs}</row>')

    r = 5
    for row in form_rows:
        kind = row[0]
        if kind == _SECTION:
            rows.append(f'<row r="{r}" ht="24" customHeight="1">{cell(f"A{r}", st["section"], row[1])}'
                        + "".join(cell(f"{c}{r}", st["section_fill"]) for c in "BCD") + "</row>")
            merges.append(f"A{r}:D{r}")
        elif kind == _REF:
            _, label, desc, var, is_date = row
            f = f'IFERROR(INDEX(Сделка!$C:$C,MATCH("{var}",Сделка!$A:$A,0))&amp;"","")'
            if is_date:  # keep it a real date for the date format; empty stays empty (not 30.12.1899)
                ref = f'INDEX(Сделка!$C:$C,MATCH("{var}",Сделка!$A:$A,0))'
                f = f'IFERROR(IF({ref}="","",{ref}),"")'
            rows.append(f'<row r="{r}">{cell(f"A{r}", st["desc"], label)}{cell(f"B{r}", st["desc"], desc)}'
                        f'<c r="C{r}" s="{st["ref_date"] if is_date else st["desc"]}"><f>{f}</f></c>'
                        f'{cell(f"D{r}", st["note"], "только просмотр")}</row>')
        else:
            _, var, desc, note, required, text_fmt = row
            kind_key = ("required" if required else "optional") + ("_text" if text_fmt else "")
            style = st[kind_key]
            if required:
                required_cells.append(f"C{r}")
            rows.append(f'<row r="{r}">{cell(f"A{r}", st["var"], var)}{cell(f"B{r}", st["desc"], desc)}'
                        f'{value_cell(f"C{r}", style, values.get(var))}{cell(f"D{r}", st["note"], note)}</row>')
        r += 1
    last = r - 1

    merge_xml = "".join(f'<mergeCell ref="{m}"/>' for m in merges)
    first = required_cells[0]
    cf = (f'<conditionalFormatting sqref="{" ".join(required_cells)}"><cfRule type="expression" '
          f'dxfId="{orange_dxf}" priority="1"><formula>ISBLANK({first})</formula></cfRule>'
          f"</conditionalFormatting>")
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<dimension ref="A1:D{last}"/>'
        '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" '
        'state="frozen"/></sheetView></sheetViews>'
        '<sheetFormatPr defaultRowHeight="15"/>'
        '<cols><col min="1" max="1" width="38" customWidth="1"/><col min="2" max="2" width="44" customWidth="1"/>'
        '<col min="3" max="3" width="50" customWidth="1"/><col min="4" max="4" width="60" customWidth="1"/></cols>'
        f'<sheetData>{"".join(rows)}</sheetData>'
        f"{SHEET_PROTECTION}"
        f'<mergeCells count="{len(merges)}">{merge_xml}</mergeCells>'
        f"{cf}"
        '<pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5"/>'
        "</worksheet>"
    )


def _append_xfs(styles: str, formats: list[tuple[int, int]]) -> tuple[str, list[int]]:
    """Append copies of existing cellXfs with another number format → their new indices."""
    m = re.search(r'<cellXfs count="(\d+)">(.*?)</cellXfs>', styles, re.S)
    if m is None:
        raise ValueError("styles.xml has no <cellXfs>")
    count, body = int(m.group(1)), m.group(2)
    xfs = re.findall(r"<xf\b[^>]*?(?:/>|>.*?</xf>)", body, re.S)
    added: list[str] = []
    idx: list[int] = []
    for base, num_fmt in formats:
        xf = re.sub(r'numFmtId="\d+"', f'numFmtId="{num_fmt}"', xfs[base], count=1)
        if "applyNumberFormat" not in xf:
            xf = xf.replace("<xf ", '<xf applyNumberFormat="1" ', 1)
        added.append(xf)
        idx.append(count + len(idx))
    new = f'<cellXfs count="{count + len(added)}">{body}{"".join(added)}</cellXfs>'
    return styles[: m.start()] + new + styles[m.end():], idx


def _read(xlsx: Path) -> tuple[dict[str, bytes], list[zipfile.ZipInfo]]:
    with zipfile.ZipFile(xlsx) as z:
        infos = z.infolist()
        return {i.filename: z.read(i.filename) for i in infos}, infos


def _write(
    xlsx: Path, data: dict[str, bytes], infos: list[zipfile.ZipInfo], extra: tuple[str, ...] = ()
) -> None:
    import os
    import tempfile

    from dca.utils.atomic import refuse_if_open_in_excel

    refuse_if_open_in_excel(xlsx)   # Excel would save its own copy over ours later
    # Unique temp name in the same folder (two runs never share it), then an atomic replace.
    fd, tmp_name = tempfile.mkstemp(prefix=f".{xlsx.stem}-", suffix=".tmp.xlsx", dir=xlsx.parent)
    os.close(fd)
    tmp = Path(tmp_name)
    try:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for i in infos:
                zout.writestr(i, data[i.filename])  # original compression + metadata preserved
            for name in extra:
                zout.writestr(name, data[name])
        tmp.replace(xlsx)
    finally:
        if tmp.exists():
            tmp.unlink()


def _unescape(s: str) -> str:
    # All XML entities, numeric ones too: openpyxl writes Cyrillic as &#1057;…
    return html.unescape(s)


# ─── Archive sheets (Стороны А, Стороны Б, Изделия, Партии) ────

MIN_ARCHIVE_SLOTS = 10   # the template's C..L; never shrink below it


def fill_archive_sheets(xlsx: Path, records: dict[str, list[dict[str, Any]]]) -> None:
    """Replace the records of the archive sheets with `records` (sheet → list of {field: value}).

    Rows = fields (matched by the text in column A), columns C… = records, in list order.
    More than 10 records → extra columns are added and every `Sheet!$C$a:$L$b` range in the
    workbook (the «Сделка» dropdowns, the «Превью» HLOOKUPs) is widened to the new last column.
    Columns A/B and all other sheets' content are left untouched; Excel recalculates on open.
    """
    data, infos = _read(xlsx)
    shared = _shared_strings(data)
    last_col: dict[str, str] = {}
    for sheet, recs in records.items():
        member = _sheet_member(data, sheet)
        last = _col_letter(2 + max(MIN_ARCHIVE_SLOTS, len(recs)))
        data[member] = _fill_archive_xml(data[member].decode("utf-8"), recs, shared, last).encode("utf-8")
        last_col[sheet] = last

    for member in [m for m in data if m.startswith("xl/worksheets/sheet")]:
        x = data[member].decode("utf-8")
        for sheet, last in last_col.items():
            # Excel writes the sheet name as is; openpyxl writes non-ASCII letters as &#NNNN;.
            spellings = {re.escape(sheet), re.escape("".join(c if ord(c) < 128 else f"&#{ord(c)};" for c in sheet))}
            name = "(?:" + "|".join(sorted(spellings)) + ")"
            x = re.sub(rf"((?:'{name}'|{name})!\$C\$\d+:\$)[A-Z]+(\$\d+)", rf"\g<1>{last}\g<2>", x)
        data[member] = x.encode("utf-8")

    wb = data["xl/workbook.xml"].decode("utf-8")
    data["xl/workbook.xml"] = re.sub(
        r"<calcPr\b(?![^>]*fullCalcOnLoad)", '<calcPr fullCalcOnLoad="1"', wb, count=1
    ).encode("utf-8")
    _write(xlsx, data, infos)


def _fill_archive_xml(xml: str, recs: list[dict[str, Any]], shared: list[str], last: str) -> str:
    last_idx = _col_index(last)
    cell_re = re.compile(r"<c\b[^>]*?(?:/>|>.*?</c>)", re.S)

    def rebuild(m: re.Match[str]) -> str:
        attrs, body = m.group(1), m.group(2)
        row_m = re.search(r'\br="(\d+)"', attrs)
        if row_m is None:
            return m.group(0)
        row = int(row_m.group(1))
        cells = cell_re.findall(body)
        keep = [c for c in cells if _col_index(_ref_col(c)) < 3]
        first = next((c for c in cells if _ref_col(c) == "C"), None)
        style_m = re.search(r'\bs="(\d+)"', first) if first else None
        style = style_m.group(1) if style_m else None
        s_attr = f' s="{style}"' if style else ""
        a_text = next((_cell_text(c, shared) for c in keep if _ref_col(c) == "A"), "")
        new: list[str] = []
        if row == 1:
            prefix = (_cell_text(first, shared) if first else "").rsplit("#", 1)[0].strip() or "Запись"
            new = [_text_cell(f"{_col_letter(i)}1", s_attr, f"{prefix} #{i - 2}") for i in range(3, last_idx + 1)]
        elif a_text:
            for i in range(3, last_idx + 1):
                ref = f"{_col_letter(i)}{row}"
                v = recs[i - 3].get(a_text) if i - 3 < len(recs) else None
                if v is None or v == "":
                    new.append(f'<c r="{ref}"{s_attr}/>')
                elif isinstance(v, (int, float)) and not isinstance(v, bool):
                    new.append(f'<c r="{ref}"{s_attr}><v>{v}</v></c>')
                else:
                    new.append(_text_cell(ref, s_attr, str(v)))
        else:
            return m.group(0)
        attrs = re.sub(r'\bspans="[^"]*"', f'spans="1:{last_idx}"', attrs)
        return f"<row{attrs}>{''.join(keep)}{''.join(new)}</row>"

    xml = re.sub(r"<row\b([^>]*)>(.*?)</row>", rebuild, xml, flags=re.S)
    xml = re.sub(r'(<dimension ref="A1:)[A-Z]+(\d+")', rf"\g<1>{last}\g<2>", xml)

    def widen(m: re.Match[str]) -> str:
        lo, hi = int(m.group(1)), int(m.group(2))
        return m.group(0).replace(f'max="{hi}"', f'max="{max(hi, last_idx)}"') if lo <= 3 <= hi else m.group(0)

    return re.sub(r'<col min="(\d+)" max="(\d+)"[^>]*/>', widen, xml)


def _text_cell(ref: str, s_attr: str, text: str) -> str:
    return f'<c r="{ref}"{s_attr} t="inlineStr"><is><t xml:space="preserve">{escape(text)}</t></is></c>'


def _ref_col(cell: str) -> str:
    m = re.match(r'<c\b[^>]*?\br="([A-Z]+)', cell)
    if m is None:
        raise ValueError(f"cell without a reference: {cell[:60]}")
    return m.group(1)


def _cell_text(cell: str, shared: list[str]) -> str:
    texts = re.findall(r"<t[^>]*>([^<]*)</t>", cell)
    if 't="s"' in cell:
        v = re.search(r"<v>(\d+)</v>", cell)
        return shared[int(v.group(1))].strip() if v else ""
    if texts:
        return _unescape("".join(texts)).strip()
    v = re.search(r"<v>([^<]*)</v>", cell)
    return _unescape(v.group(1)).strip() if v else ""


def _shared_strings(data: dict[str, bytes]) -> list[str]:
    raw = data.get("xl/sharedStrings.xml")
    if raw is None:
        return []
    items = re.findall(r"<si>(.*?)</si>", raw.decode("utf-8"), re.S)
    return [_unescape("".join(re.findall(r"<t[^>]*>([^<]*)</t>", si))) for si in items]


def _col_index(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + ord(ch) - 64
    return n


def _col_letter(idx: int) -> str:
    out = ""
    while idx:
        idx, rem = divmod(idx - 1, 26)
        out = chr(65 + rem) + out
    return out
