"""Generate the demo templates and Excel files of DCA.

    python3 scripts/make_demo_templates.py

Writes (overwrites):
  templates/agreement.docx      — conditional paragraphs (deal_kind, variant), a table row loop,
                                  dates in several formats, a declined company name, amount in words
  templates/annex.docx          — rendered once per annex
  templates/amendment.docx      — clause numbers computed by the engine
  templates/transfer_act.docx   — fixed-width form boxes filled line by line
  templates/cover_list.docx     — numbered list with a Jinja counter
  data/input_template_EMPTY.xlsx
  data/input_example_FILLED.xlsx

Every name and number in these files is invented. Document properties: author "demo".
Run it again after changing a template here; the tests use the generated files.
"""

from __future__ import annotations

import sys
from datetime import date, datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

from docx import Document  # noqa: E402
from docx.enum.text import WD_ALIGN_PARAGRAPH  # noqa: E402
from docx.shared import Pt  # noqa: E402
from openpyxl import Workbook  # noqa: E402
from openpyxl.formatting.rule import FormulaRule  # noqa: E402
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side  # noqa: E402
from openpyxl.worksheet.datavalidation import DataValidation  # noqa: E402

from dca.xlsx_patch import add_amendment_sheet, add_transfer_sheet  # noqa: E402

TEMPLATES = ROOT / "templates"
DATA = ROOT / "data"
AUTHOR = "demo"
FIXED_TIME = datetime(2026, 1, 1)   # stable document properties → stable sha256 in log.json


# ─── docx helpers ───────────────────────────────────────────────


def _doc() -> Document:
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style.font.size = Pt(10)
    cp = doc.core_properties
    cp.author = AUTHOR
    cp.last_modified_by = AUTHOR
    cp.title = ""
    cp.comments = ""
    cp.keywords = ""
    cp.subject = ""
    cp.category = ""
    cp.created = FIXED_TIME
    cp.modified = FIXED_TIME
    cp.revision = 1
    return doc


def _p(doc_or_cell, text: str = "", *, bold: bool = False, italic: bool = False,
       center: bool = False, size: float | None = None) -> None:
    """One paragraph with ONE run, so a Jinja tag is never split across runs."""
    para = doc_or_cell.add_paragraph()
    run = para.add_run(text)
    run.bold, run.italic = bold, italic
    if size:
        run.font.size = Pt(size)
    if center:
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER


def _cell(table, r: int, c: int, text: str, *, bold: bool = False, size: float | None = None) -> None:
    cell = table.cell(r, c)
    cell.paragraphs[0].text = ""
    run = cell.paragraphs[0].add_run(text)
    run.bold = bold
    if size:
        run.font.size = Pt(size)


def _save(doc: Document, name: str) -> None:
    TEMPLATES.mkdir(exist_ok=True)
    path = TEMPLATES / name
    doc.save(str(path))
    _clean_package(path)
    print("  ✓ templates/" + name)


def _clean_package(path: Path) -> None:
    """Strip what python-docx's default template carries but the demo doesn't need:
    the custom-style marks on styles and the company fields of docProps/app.xml."""
    import re
    import zipfile

    with zipfile.ZipFile(path) as z:
        infos = z.infolist()
        data = {i.filename: z.read(i.filename) for i in infos}
    for member in ("word/styles.xml", "word/stylesWithEffects.xml"):
        if member in data:
            data[member] = re.sub(rb' w:custom[A-Za-z]+="1"', b"", data[member])
    if "docProps/app.xml" in data:
        app = data["docProps/app.xml"].decode("utf-8")
        app = re.sub(r"<(Company|Manager|Template)>[^<]*</\1>", "", app)
        data["docProps/app.xml"] = app.encode("utf-8")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for i in infos:
            z.writestr(i, data[i.filename])


# ─── templates ──────────────────────────────────────────────────


def make_agreement() -> None:
    d = _doc()
    _p(d, "СОГЛАШЕНИЕ № {{agreement_num}}", bold=True, center=True, size=13)
    _p(d, "AGREEMENT No. {{agreement_num}}", bold=True, center=True, italic=True)
    _p(d, "{{party_a_city_ru}}, {{agreement_date_ru_long}}")
    _p(d, "{{party_a_city_eng}}, {{agreement_date_eng_long}}", italic=True)
    _p(d, "{{party_a_name_ru}}, в лице {{party_a_represented_by_ru}}, далее «Сторона А», и "
          "{{party_b_name_ru}}, в лице {{party_b_represented_by_ru}}, далее «Сторона Б», "
          "заключили настоящее соглашение.")
    _p(d, "{{party_a_name_eng}}, represented by {{party_a_represented_by_eng}}, hereinafter "
          "“Party A”, and {{party_b_name_eng}}, represented by {{party_b_represented_by_eng}}, "
          "hereinafter “Party B”, have concluded this agreement.", italic=True)

    _p(d, "1. Предмет / Subject", bold=True)
    _p(d, '{%p if deal_kind == "outgoing" %}')
    _p(d, "1.1. Сторона А передаёт, а Сторона Б принимает {{item_name_ru}} ({{item_description_ru}}) "
          "в количестве {{quantity_kg_phrase_ru}} в срок до {{due_date_ru_full}}.")
    _p(d, "1.1. Party A hands over and Party B accepts {{item_name_eng}} ({{item_description_eng}}), "
          "{{quantity_kg_phrase_eng}}, by {{due_date_eng_long}}.", italic=True)
    _p(d, "{%p else %}")
    _p(d, "1.1. Сторона Б получает от Стороны А {{item_name_ru}} ({{item_description_ru}}) "
          "в количестве {{quantity_kg_phrase_ru}} в месте: {{place_ru}}.")
    _p(d, "1.1. Party B receives from Party A {{item_name_eng}} ({{item_description_eng}}), "
          "{{quantity_kg_phrase_eng}}, at: {{place_eng}}.", italic=True)
    _p(d, "{%p endif %}")
    _p(d, "1.2. Количество и сроки по каждой части указаны в приложениях:")
    t = d.add_table(rows=3, cols=4)
    t.style = "Table Grid"
    for c, head in enumerate(("Приложение №", "Дата", "Количество, кг", "Сумма, USD")):
        _cell(t, 0, c, head, bold=True)
    _cell(t, 1, 0, "{%tr for a in annex_rows %}")
    _cell(t, 2, 0, "{{a.num}}")
    _cell(t, 2, 1, "{{a.date}}")
    _cell(t, 2, 2, "{{a.quantity_kg}}")
    _cell(t, 2, 3, "{{a.total}}")
    row = t.add_row()
    row.cells[0].paragraphs[0].add_run("{%tr endfor %}")

    _p(d, "2. Цена / Price", bold=True)
    _p(d, "2.1. Цена: {{price_per_kg_phrase_ru}}.")
    _p(d, "2.1. Price: {{price_per_kg_phrase_eng}}.", italic=True)
    _p(d, "2.2. Общая сумма: {{total_amount_str_ru}} ({{total_amount_words_ru}}).")
    _p(d, "2.2. Total amount: {{total_amount_str}} ({{total_amount_words_eng}}).", italic=True)

    _p(d, '{%p if variant == "B" %}')
    _p(d, "3. Сторона В / Party C", bold=True)
    _p(d, "3.1. Сторона А исполняет соглашение через Сторону В: {{party_c_name}}, {{party_c_address}}, "
          "{{party_c_details}}.")
    _p(d, "3.1. Party A performs the agreement through Party C: {{party_c_name}}, {{party_c_address}}, "
          "{{party_c_details}}.", italic=True)
    _p(d, "{%p endif %}")

    _p(d, "Срок действия / Term", bold=True)
    _p(d, "Соглашение действует до {{due_date_ru_full}}.")
    _p(d, "The agreement is valid until {{due_date_eng}}.", italic=True)

    _p(d, "Подписи / Signatures", bold=True)
    _p(d, "Сторона А: {{party_a_name_ru}}, {{party_a_address_full_ru}}, {{party_a_details_full_ru}}")
    _p(d, "Директор {{party_a_name_ru_genitive}} ______________ {{party_a_signatory_name_ru}}")
    _p(d, "Сторона Б: {{party_b_name_ru}}, {{party_b_address_full_ru}}, {{party_b_details_full_ru}}")
    _p(d, "______________ {{party_b_signatory_name_ru}}")
    _save(d, "agreement.docx")


def make_annex() -> None:
    d = _doc()
    _p(d, "ПРИЛОЖЕНИЕ № {{annex_num}}", bold=True, center=True, size=13)
    _p(d, "ANNEX No. {{annex_num}}", bold=True, center=True, italic=True)
    _p(d, "от {{annex_date_ru_long}} к соглашению № {{agreement_num}} от {{agreement_date_ru}}")
    _p(d, "dated {{annex_date_eng_long}} to agreement No. {{agreement_num}} of {{agreement_date_eng}}",
       italic=True)
    t = d.add_table(rows=2, cols=4)
    t.style = "Table Grid"
    for c, head in enumerate(("Изделие / Item", "Количество, кг", "Цена", "Сумма, USD")):
        _cell(t, 0, c, head, bold=True)
    _cell(t, 1, 0, "{{item_name_ru}} / {{item_name_eng}}")
    _cell(t, 1, 1, "{{quantity_kg}}")
    _cell(t, 1, 2, "{{price_per_unit_ru}}")
    _cell(t, 1, 3, "{{total_amount_str_ru}}")
    _p(d, "Итого: {{total_amount_number_words_ru}} {{currency_word_ru_for_total}} {{cents_phrase_ru}}.")
    _p(d, "Total: {{total_amount_number_words_eng}} {{currency_word_eng_for_total}} {{cents_phrase_eng}}.",
       italic=True)
    _p(d, "Место / Place: {{place_ru}} / {{place_eng}}")
    _p(d, "______________ {{party_a_signatory_name_ru}}          ______________ {{party_b_signatory_name_ru}}")
    _save(d, "annex.docx")


def make_amendment() -> None:
    d = _doc()
    _p(d, "ИЗМЕНЕНИЕ № {{amendment_num}}", bold=True, center=True, size=13)
    _p(d, "к соглашению № {{agreement_num}} от {{agreement_date_ru}}", center=True)
    _p(d, "{{party_a_city_ru}}, {{amendment_date_ru_long}}")
    _p(d, "1. Стороны договорились изменить соглашение № {{agreement_num}}.")
    _p(d, "{%p if price_changed %}")
    _p(d, "{{item_price}}. Цена: было {{old_price_per_kg_phrase_ru}}, стало {{new_price_per_kg_phrase_ru}}.")
    _p(d, "{%p endif %}")
    _p(d, "{%p if quantity_changed %}")
    _p(d, "{{item_quantity}}. Количество: было {{old_quantity_kg}} кг, стало {{new_quantity_kg}} кг.")
    _p(d, "{%p endif %}")
    _p(d, "{{item_total}}. Общая сумма: было {{old_total_amount_str_ru}}, стало {{new_total_amount_str_ru}} "
          "({{new_total_amount_words_ru}}).")
    _p(d, "{{item_rest}}. Остальные условия соглашения не меняются.")
    _p(d, "Директор {{party_a_name_ru_genitive}} ______________ {{party_a_signatory_name_ru}}")
    _p(d, "______________ {{party_b_signatory_name_ru}}")
    _save(d, "amendment.docx")


def make_transfer_act() -> None:
    d = _doc()
    _p(d, "АКТ ПЕРЕДАЧИ № {{transfer_num}} от {{transfer_date_ru}}", bold=True, center=True, size=12)
    _p(d, "TRANSFER ACT No. {{transfer_num}} of {{transfer_date_eng}}", center=True, italic=True)
    boxes = (("1. Передал / Handed over by", "sender", 4), ("2. Принял / Received by", "receiver", 4),
             ("3. Изделие / Item", "item", 3), ("4. Примечание / Note", "note", 3))
    n_rows = sum(rows for _, _, rows in boxes) + 4
    t = d.add_table(rows=n_rows, cols=2)
    t.style = "Table Grid"
    r = 0
    for label, box, rows in boxes:
        for i in range(1, rows + 1):
            _cell(t, r, 0, label if i == 1 else "", bold=True, size=8)
            _cell(t, r, 1, "{{transfer_" + box + "_line_" + str(i) + "}}", size=8)
            r += 1
    for label, key in (("5. Откуда / From", "place_from_ru"), ("6. Куда / To", "place_to_ru"),
                       ("7. Основание / Basis", "transfer_annexes_ref"),
                       ("8. Соглашение / Agreement", "agreement_num")):
        _cell(t, r, 0, label, bold=True, size=8)
        _cell(t, r, 1, "{{" + key + "}}", size=8)
        r += 1
    _p(d, "Количество: {{quantity_kg}} кг. Передал: {{handed_by_ru}}. Принял: {{received_by_ru}}.")
    _save(d, "transfer_act.docx")


def make_cover_list() -> None:
    d = _doc()
    _p(d, "Исх. № {{letter_num}} от {{letter_date_ru}}", size=9)
    _p(d, "ОПИСЬ ПРИЛОЖЕНИЙ", bold=True, center=True, size=13)
    _p(d, "к соглашению № {{agreement_num}} от {{agreement_date_ru_long}}. Стороны: {{party_a_name_ru}}, "
          "{{party_b_name_ru}}", center=True)
    # The counter must match dca.attachments.list_attachment_items (same order, same numbers).
    _p(d, "{% set ns = namespace(n=1) %}1. соглашение № {{agreement_num}} от {{agreement_date_ru}};")
    _p(d, "{%p for b in batches %}")
    _p(d, "{% set ns.n = ns.n + 1 %}{{ns.n}}. копия соглашения № {{b.agreement_num}} от {{b.agreement_date}} "
          "(партия {{b.key}}, {{b.party_name}}, {{b.party_location}}; использовано {{b.use_kg}} кг);")
    _p(d, "{%p if b.annex_num %}")
    _p(d, "{% set ns.n = ns.n + 1 %}{{ns.n}}. копия приложения № {{b.annex_num}} от {{b.annex_date}} "
          "к соглашению № {{b.agreement_num}} от {{b.agreement_date}};")
    _p(d, "{%p endif %}")
    _p(d, "{%p endfor %}")
    _p(d, "{{ns.n + 1}}. карточка компании.")
    _p(d, "Директор {{party_a_name_ru_genitive}} ______________ {{party_a_signatory_name_ru}}")
    _save(d, "cover_list.docx")


# ─── Excel ──────────────────────────────────────────────────────

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
SECTION_FILL = PatternFill("solid", fgColor="D9E1F2")
ORANGE = "FFC000"
GREEN = "C6EFCE"
PALE = "EBF1DE"
WRAP = Alignment(wrap_text=True, vertical="top")

# «Сделка»: ("section", title) or ("field", var, description, note, required, kind)
# kind: "text" (text format «@»), "date", "number", "list:<a,b>", "key:<sheet>"
REQ, OPT = True, False
DEAL_ROWS = [
    ("section", "  ━━━ СТОРОНЫ И ИЗДЕЛИЕ (выберите из списков) ━━━"),
    ("field", "party_a_key", "Сторона А", "[обязательное · список] Запись листа «Стороны А».", REQ, "key:Стороны А"),
    ("field", "party_b_key", "Сторона Б", "[обязательное · список] Запись листа «Стороны Б».", REQ, "key:Стороны Б"),
    ("field", "item_key", "Изделие", "[обязательное · список] Запись листа «Изделия».", REQ, "key:Изделия"),
    ("section", "  ━━━ СОГЛАШЕНИЕ ━━━"),
    ("field", "agreement_num", "Номер соглашения", "[обязательное · текст] Например 1-2026-1705. Уникален.", REQ, "text"),
    ("field", "agreement_date", "Дата соглашения", "[обязательное · дата] DD.MM.YYYY.", REQ, "date"),
    ("field", "annex_num", "Номер приложения", "[текст] Пусто → часть номера соглашения до первого «-».", OPT, "text"),
    ("field", "annex_date", "Дата приложения", "[дата] Пусто → дата соглашения.", OPT, "date"),
    ("field", "due_date", "Срок исполнения", "[дата] Пусто → дата соглашения + 2 месяца.", OPT, "date"),
    ("section", "  ━━━ ЦЕНА И КОЛИЧЕСТВО ━━━"),
    ("field", "price_per_unit", "Цена, USD за 1 кг", "[обязательное · число] Печатается как введена.", REQ, "number"),
    ("field", "quantity_kg", "Количество, кг", "[обязательное · число] В документах — целое, округление вверх.", REQ, "number"),
    ("field", "currency", "Валюта", "[список] Пусто → USD. Пока поддерживается только USD.", OPT, "list:USD,EUR"),
    ("field", "unit_of_measure", "Единица", "[список] Пусто → kg. Пока поддерживается только kg.", OPT, "list:kg,pcs,t"),
    ("section", "  ━━━ МЕСТО ━━━"),
    ("field", "place_ru", "Место исполнения (RU)", "[обязательное]", REQ, "text"),
    ("field", "place_eng", "Место исполнения (EN)", "[обязательное]", REQ, "text"),
    ("section", "  ━━━ ВАРИАНТ ━━━"),
    ("field", "variant", "Вариант исполнения", "[обязательное · список] A — без стороны В; B — через сторону В.", REQ, "list:A,B"),
    ("section", "  ━━━ СТОРОНА В (только для варианта B) ━━━"),
    ("field", "party_c_name", "Сторона В — название", "При варианте B обязательно.", OPT, "text"),
    ("field", "party_c_address", "Сторона В — адрес", "При варианте B обязательно.", OPT, "text"),
    ("field", "party_c_details", "Сторона В — реквизиты", "При варианте B обязательно.", OPT, "text"),
    ("section", "  ━━━ ОПИСЬ ПРИЛОЖЕНИЙ ━━━"),
    ("field", "letter_num", "Исходящий номер описи", "[текст] Пусто → 1.", OPT, "text"),
    ("field", "letter_date", "Дата описи", "[дата] Пусто → дата первой сборки.", OPT, "date"),
    ("section", "  ━━━ ПАРТИИ (до 5, без пропусков) ━━━"),
]
for _n in range(1, 6):
    DEAL_ROWS += [
        ("field", f"batch_key_{_n}", f"Партия №{_n}", "[список] Запись листа «Партии».", OPT, "key:Партии"),
        ("field", f"batch_use_kg_{_n}", f"Количество из партии №{_n}, кг",
         "[число] Сумма по партиям = количество сделки.", OPT, "number"),
    ]

ARCHIVES = {
    "Стороны А": [
        ("party_a_key", "Ключ записи (короткое имя)"),
        ("party_a_folder_name", "Папка компании (только для наших компаний)"),
        ("party_a_name_ru", "Название (RU)"), ("party_a_name_eng", "Название (EN)"),
        ("party_a_city_ru", "Город (RU)"), ("party_a_city_eng", "Город (EN)"),
        ("party_a_represented_by_ru", "В лице (RU, род. п.)"), ("party_a_represented_by_eng", "Represented by (EN)"),
        ("party_a_address_full_ru", "Адрес (RU)"), ("party_a_address_full_eng", "Адрес (EN)"),
        ("party_a_details_full_ru", "Реквизиты (RU)"), ("party_a_details_full_eng", "Реквизиты (EN)"),
        ("party_a_signatory_name_ru", "Подписант (RU)"), ("party_a_signatory_name_eng", "Подписант (EN)"),
    ],
    "Стороны Б": [
        ("party_b_key", "Ключ записи (короткое имя)"),
        ("party_b_folder_name", "Папка компании (только для наших компаний)"),
        ("party_b_name_ru", "Название (RU)"), ("party_b_name_eng", "Название (EN)"),
        ("party_b_represented_by_ru", "В лице (RU, род. п.)"), ("party_b_represented_by_eng", "Represented by (EN)"),
        ("party_b_address_full_ru", "Адрес (RU)"), ("party_b_address_full_eng", "Адрес (EN)"),
        ("party_b_details_full_ru", "Реквизиты (RU)"), ("party_b_details_full_eng", "Реквизиты (EN)"),
        ("party_b_signatory_name_ru", "Подписант (RU)"), ("party_b_signatory_name_eng", "Подписант (EN)"),
    ],
    "Изделия": [
        ("item_key", "Ключ записи"),
        ("item_name_ru", "Название (RU)"), ("item_name_eng", "Название (EN)"),
        ("item_description_ru", "Описание (RU)"), ("item_description_eng", "Описание (EN)"),
    ],
    "Партии": [
        ("batch_key", "Ключ партии"),
        ("batch_agreement_num", "Номер соглашения партии"), ("batch_agreement_date", "Дата соглашения партии"),
        ("batch_annex_num", "Номер приложения (опционально)"), ("batch_annex_date", "Дата приложения"),
        ("batch_party_name_ru", "Контрагент партии"), ("batch_party_location_ru", "Место контрагента"),
        ("batch_quantity_kg", "Количество партии, кг (всего)"),
        ("batch_agreement_file", "Файл скана соглашения (в Data/scans)"),
        ("batch_annex_file", "Файл скана приложения"),
        ("batch_company_key", "Чья партия (ключ «Стороны А»)"), ("batch_company_name_ru", "Чья партия (название)"),
        ("batch_scans_folder", "Папка со сканами в Data/scans (опционально)"),
    ],
}
ARCHIVE_SLOTS = 10

EXAMPLE_RECORDS = {
    "Стороны А": [
        {"party_a_key": "alpha", "party_a_folder_name": "Компания А",
         "party_a_name_ru": "Компания «Альфа»", "party_a_name_eng": "Alpha Company",
         "party_a_city_ru": "Город Б", "party_a_city_eng": "City B",
         "party_a_represented_by_ru": "директора Орлова Олега Олеговича",
         "party_a_represented_by_eng": "Director Oleg Orlov",
         "party_a_address_full_ru": "Город Б, ул. Первая, 1", "party_a_address_full_eng": "1 First Street, City B",
         "party_a_details_full_ru": "рег. № 0000000001", "party_a_details_full_eng": "reg. No. 0000000001",
         "party_a_signatory_name_ru": "Орлов О. О.", "party_a_signatory_name_eng": "O. Orlov"},
        {"party_a_key": "omega", "party_a_folder_name": "Компания О",
         "party_a_name_ru": "Общество «Омега»", "party_a_name_eng": "Omega Ltd",
         "party_a_city_ru": "Город Г", "party_a_city_eng": "City G",
         "party_a_represented_by_ru": "директора Петрова Петра Петровича",
         "party_a_represented_by_eng": "Director Petr Petrov",
         "party_a_address_full_ru": "Город Г, ул. Вторая, 2", "party_a_address_full_eng": "2 Second Street, City G",
         "party_a_details_full_ru": "рег. № 0000000002", "party_a_details_full_eng": "reg. No. 0000000002",
         "party_a_signatory_name_ru": "Петров П. П.", "party_a_signatory_name_eng": "P. Petrov"},
    ],
    "Стороны Б": [
        {"party_b_key": "beta", "party_b_name_ru": "Общество «Бета»", "party_b_name_eng": "Beta Ltd",
         "party_b_represented_by_ru": "директора Сидорова Сидора Сидоровича",
         "party_b_represented_by_eng": "Director Sidor Sidorov",
         "party_b_address_full_ru": "Город В, ул. Третья, 3", "party_b_address_full_eng": "3 Third Street, City V",
         "party_b_details_full_ru": "рег. № 0000000003", "party_b_details_full_eng": "reg. No. 0000000003",
         "party_b_signatory_name_ru": "Сидоров С. С.", "party_b_signatory_name_eng": "S. Sidorov"},
        {"party_b_key": "alpha-b", "party_b_folder_name": "Компания А",
         "party_b_name_ru": "Компания «Альфа»", "party_b_name_eng": "Alpha Company",
         "party_b_represented_by_ru": "директора Орлова Олега Олеговича",
         "party_b_represented_by_eng": "Director Oleg Orlov",
         "party_b_address_full_ru": "Город Б, ул. Первая, 1", "party_b_address_full_eng": "1 First Street, City B",
         "party_b_details_full_ru": "рег. № 0000000001", "party_b_details_full_eng": "reg. No. 0000000001",
         "party_b_signatory_name_ru": "Орлов О. О.", "party_b_signatory_name_eng": "O. Orlov"},
    ],
    "Изделия": [
        {"item_key": "X", "item_name_ru": "Изделие X", "item_name_eng": "Item X",
         "item_description_ru": "исполнение 1, партия по описи", "item_description_eng": "version 1, batch as listed"},
        {"item_key": "Y", "item_name_ru": "Изделие Y", "item_name_eng": "Item Y",
         "item_description_ru": "исполнение 2", "item_description_eng": "version 2"},
    ],
    "Партии": [
        {"batch_key": "П-001", "batch_agreement_num": "В-17/2025", "batch_agreement_date": "15.09.2025",
         "batch_annex_num": "1", "batch_annex_date": "20.09.2025",
         "batch_party_name_ru": "Компания «Дельта»", "batch_party_location_ru": "Город Д",
         "batch_quantity_kg": 5000, "batch_agreement_file": "соглашение В-17.pdf",
         "batch_annex_file": "приложение 1 к В-17.pdf",
         "batch_company_key": "alpha", "batch_company_name_ru": "Компания «Альфа»", "batch_scans_folder": "П-001"},
        {"batch_key": "П-002", "batch_agreement_num": "В-21/2025", "batch_agreement_date": "01.11.2025",
         "batch_party_name_ru": "Компания «Дельта»", "batch_party_location_ru": "Город Д",
         "batch_quantity_kg": 2000, "batch_agreement_file": "соглашение В-21.pdf",
         "batch_company_key": "alpha", "batch_company_name_ru": "Компания «Альфа»"},
    ],
}

EXAMPLE_DEAL = {
    "party_a_key": "alpha", "party_b_key": "beta", "item_key": "X",
    "agreement_num": "1-2026-1705", "agreement_date": date(2026, 5, 17),
    "price_per_unit": 0.45, "quantity_kg": 1000, "currency": "USD",
    "place_ru": "Город В", "place_eng": "City V",
    "variant": "B",
    "party_c_name": "Gamma Agency", "party_c_address": "4 Fourth Street, City G",
    "party_c_details": "reg. No. 0000000004",
    "letter_num": "7",
    "batch_key_1": "П-001", "batch_use_kg_1": 700,
    "batch_key_2": "П-002", "batch_use_kg_2": 300,
}
EXAMPLE_ANNEXES = [("1", date(2026, 5, 17), 600), ("2", date(2026, 6, 1), 400)]
EXAMPLE_TRANSFER = {
    "transfer_num": "00017", "transfer_date": "25.05.2026",
    "place_from_ru": "Город Б, склад 1", "place_to_ru": "Город В, склад 2",
    "handed_by_ru": "Орлов Олег Олегович", "received_by_ru": "Сидоров Сидор Сидорович",
}
EXAMPLE_AMENDMENT = {"amendment_num": "1", "amendment_date": "01.07.2026", "new_quantity_kg": 1200}


def _set_props(wb: Workbook) -> None:
    wb.properties.creator = AUTHOR
    wb.properties.lastModifiedBy = AUTHOR
    wb.properties.created = FIXED_TIME
    wb.properties.modified = FIXED_TIME
    wb.properties.title = None
    wb.properties.company = None if hasattr(wb.properties, "company") else None


def _deal_sheet(ws, values: dict) -> None:
    for c, (head, width) in enumerate((("Variable name", 30), ("Описание (что вводить)", 40),
                                       ("Value", 36), ("Notes / Help", 56)), start=1):
        cell = ws.cell(row=1, column=c, value=head)
        cell.font, cell.fill, cell.alignment = Font(bold=True, color="FFFFFF"), HEADER_FILL, Alignment(vertical="center")
        ws.column_dimensions["ABCD"[c - 1]].width = width
    ws.row_dimensions[1].height = 24
    ws["A2"] = "ЛЕГЕНДА:"
    ws["A2"].font = Font(bold=True)
    for ref, text, color in (("C2", "оранжевый = обязательное и ПУСТО", ORANGE),
                             ("C3", "зелёный = обязательное и ЗАПОЛНЕНО", GREEN),
                             ("C4", "бледно-зелёный = опциональное", PALE)):
        ws[ref] = text
        ws[ref].fill = PatternFill("solid", fgColor=color)

    required_cells: list[str] = []
    validations: dict[str, DataValidation] = {}
    r = 5
    for row in DEAL_ROWS:
        if row[0] == "section":
            ws.cell(row=r, column=1, value=row[1]).font = Font(bold=True)
            for c in range(1, 5):
                ws.cell(row=r, column=c).fill = SECTION_FILL
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
            ws.row_dimensions[r].height = 24
            r += 1
            continue
        _, var, desc, note, required, kind = row
        ws.cell(row=r, column=1, value=var).font = Font(bold=True)
        ws.cell(row=r, column=2, value=desc).alignment = WRAP
        ws.cell(row=r, column=4, value=note).alignment = WRAP
        value_cell = ws.cell(row=r, column=3, value=values.get(var))
        value_cell.border = BORDER
        value_cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        if kind == "text":
            value_cell.number_format = "@"
        elif kind == "date":
            value_cell.number_format = "DD.MM.YYYY"
        if required:
            required_cells.append(f"C{r}")
        else:
            value_cell.fill = PatternFill("solid", fgColor=PALE)
        if kind.startswith(("list:", "key:")):
            if kind.startswith("list:"):
                formula = f'"{kind[5:]}"'
            else:
                formula = f"'{kind[4:]}'!$C$2:$L$2"
            dv = validations.get(formula)
            if dv is None:
                dv = DataValidation(type="list", formula1=formula, allow_blank=True)
                validations[formula] = dv
                ws.add_data_validation(dv)
            dv.add(f"C{r}")
        r += 1
    rng = " ".join(required_cells)
    first = required_cells[0]
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f"ISBLANK({first})"],
                                                   fill=PatternFill("solid", fgColor=ORANGE, bgColor=ORANGE)))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f"NOT(ISBLANK({first}))"],
                                                   fill=PatternFill("solid", fgColor=GREEN, bgColor=GREEN)))
    ws.freeze_panes = "A2"


def _archive_sheet(ws, fields: list[tuple[str, str]], records: list[dict]) -> None:
    heads = ["Variable name", "Описание", *(f"Запись #{i}" for i in range(1, ARCHIVE_SLOTS + 1))]
    for c, text in enumerate(heads, start=1):
        cell = ws.cell(row=1, column=c, value=text)
        cell.font, cell.fill = Font(bold=True, color="FFFFFF"), HEADER_FILL
    for r, (name, desc) in enumerate(fields, start=2):
        ws.cell(row=r, column=1, value=name).font = Font(bold=True)
        ws.cell(row=r, column=2, value=desc).alignment = WRAP
        for c in range(3, 3 + ARCHIVE_SLOTS):
            cell = ws.cell(row=r, column=c)
            cell.alignment = WRAP
            if r == 2:
                cell.fill = PatternFill("solid", fgColor="FFF2CC")
        for i, rec in enumerate(records):
            if rec.get(name) not in (None, ""):
                ws.cell(row=r, column=3 + i, value=rec[name])
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 36
    for c in range(3, 3 + ARCHIVE_SLOTS):
        ws.column_dimensions[ws.cell(row=1, column=c).column_letter].width = 24
    ws.freeze_panes = "C2"


def _annex_sheet(ws, rows: list[tuple[str, date, float]]) -> None:
    for c, head in enumerate(("annex_num", "annex_date", "quantity_kg"), start=1):
        cell = ws.cell(row=1, column=c, value=head)
        cell.font, cell.fill = Font(bold=True, color="FFFFFF"), HEADER_FILL
        ws.column_dimensions["ABC"[c - 1]].width = 18
    ws["E1"] = "Несколько приложений к одному соглашению: сумма количеств = quantity_kg на «Сделке»."
    for r, (num, d, q) in enumerate(rows, start=2):
        ws.cell(row=r, column=1, value=num).number_format = "@"
        ws.cell(row=r, column=2, value=d).number_format = "DD.MM.YYYY"
        ws.cell(row=r, column=3, value=q)


def _preview_sheet(ws) -> None:
    ws["A1"] = "Превью (только просмотр, формулы)"
    ws["A1"].font = Font(bold=True)
    rows = (
        ("Сторона А", "=IFERROR(HLOOKUP(Сделка!$C$6,'Стороны А'!$C$2:$L$20,3,FALSE),\"\")"),
        ("Сторона Б", "=IFERROR(HLOOKUP(Сделка!$C$7,'Стороны Б'!$C$2:$L$20,3,FALSE),\"\")"),
        ("Изделие", "=IFERROR(HLOOKUP(Сделка!$C$8,'Изделия'!$C$2:$L$20,2,FALSE),\"\")"),
        ("Сумма, USD", "=IFERROR(CEILING(CEILING(Сделка!$C$17,1)*Сделка!$C$16,1),\"\")"),
    )
    for r, (label, formula) in enumerate(rows, start=3):
        ws.cell(row=r, column=1, value=label).font = Font(bold=True)
        ws.cell(row=r, column=2, value=formula)
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 40


def make_workbook(path: Path, *, filled: bool) -> None:
    wb = Workbook()
    _set_props(wb)
    ws = wb.active
    ws.title = "Сделка"
    _deal_sheet(ws, EXAMPLE_DEAL if filled else {})
    _annex_sheet(wb.create_sheet("Приложения"), EXAMPLE_ANNEXES if filled else [])
    for sheet, fields in ARCHIVES.items():
        _archive_sheet(wb.create_sheet(sheet), fields, EXAMPLE_RECORDS[sheet] if filled else [])
    _preview_sheet(wb.create_sheet("Превью"))
    path.parent.mkdir(exist_ok=True)
    wb.save(path)
    # Form sheets go in by XML patch — the same code that adds them to an older deal.
    add_transfer_sheet(path, EXAMPLE_TRANSFER if filled else None)
    add_amendment_sheet(path, EXAMPLE_AMENDMENT if filled else None)
    print("  ✓ data/" + path.name)


def main() -> None:
    print("Templates:")
    make_agreement()
    make_annex()
    make_amendment()
    make_transfer_act()
    make_cover_list()
    print("Excel:")
    make_workbook(DATA / "input_template_EMPTY.xlsx", filled=False)
    make_workbook(DATA / "input_example_FILLED.xlsx", filled=True)


if __name__ == "__main__":
    main()
