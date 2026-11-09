"""The demo templates rendered by the engine: conditional paragraphs, table rows, dates,
declension, amounts in words, one file per annex, strict variable names."""

from datetime import date

import pytest
from docx import Document

from dca.config import TEMPLATE_FILES, TEMPLATES_DIR
from dca.context import build_render_context, org_name_genitive_ru
from dca.models import AnnexLine
from dca.render import (
    active_kinds_for,
    render_all,
    render_amendment,
    render_transfer,
    render_with_context,
)
from tests.helpers import deal, docx_text, example_deal, incoming_deal


def _agreement(d) -> str:
    return docx_text(render_all(d)[f"Agreement_{d.agreement_num}.docx"])


def test_outgoing_package_files():
    names = set(render_all(example_deal()))
    assert names == {"Agreement_1-2026-1705.docx", "Annex_1_1-2026-1705.docx",
                     "Annex_2_1-2026-1705.docx", "CoverList_1-2026-1705.docx"}


def test_incoming_package_has_no_cover_list():
    assert active_kinds_for(incoming_deal()) == ["agreement", "annex"]
    assert set(render_all(incoming_deal())) == {"Agreement_1-2026-1705.docx", "Annex_1_1-2026-1705.docx",
                                                "Annex_2_1-2026-1705.docx"}


def test_single_annex_keeps_the_short_name():
    assert "Annex_1-2026-1705.docx" in render_all(deal())


def test_agreement_paragraph_depends_on_deal_kind():
    out, inc = _agreement(deal()), _agreement(incoming_deal(annexes=[]))
    assert "Сторона А передаёт, а Сторона Б принимает" in out and "получает от Стороны А" not in out
    assert "Сторона Б получает от Стороны А" in inc and "передаёт, а Сторона Б" not in inc


def test_variant_b_adds_the_party_c_clause_and_nothing_else_changes():
    b = _agreement(deal(variant="B"))
    a = _agreement(deal(variant="A", party_c_name=None, party_c_address=None, party_c_details=None))
    assert "через Сторону В: Gamma Agency, 4 Fourth Street, City G" in b
    assert "Сторона В" not in a and "Party C" not in a
    extra = [line for line in b.splitlines() if line not in a.splitlines()]
    assert len(extra) == 3                     # heading + RU + EN of the clause


def test_annex_rows_table_has_one_row_per_annex():
    text = _agreement(example_deal())
    assert "1 | 17.05.2026 | 600 | 270,00" in text
    assert "2 | 01.06.2026 | 400 | 180,00" in text
    assert "{%" not in text and "{{" not in text


def test_dates_declension_and_amount_in_words():
    text = _agreement(deal())
    assert "Город Б, «17» мая 2026 г." in text and "City B, May 17, 2026" in text
    assert "Соглашение действует до 17 июля 2026 года." in text
    assert "Директор Компании «Альфа»" in text                 # genitive of «Компания «Альфа»»
    assert "450,00 (Четыреста пятьдесят долларов США 00 центов)" in text
    assert "450.00 (Four hundred fifty US dollars 00 cents)" in text
    assert "0,45 доллара США за 1 кг" in text


@pytest.mark.parametrize("name, genitive", [
    ("Компания «Альфа»", "Компании «Альфа»"),
    ("Общество «Бета»", "Общества «Бета»"),
    ("ООО «Бета»", "ООО «Бета»"),
    ("АКЦИОНЕРНОЕ ОБЩЕСТВО «Гамма»", "АКЦИОНЕРНОГО ОБЩЕСТВА «Гамма»"),
])
def test_org_name_genitive(name, genitive):
    assert org_name_genitive_ru(name) == genitive


def test_each_annex_has_its_own_number_quantity_and_total():
    out = render_all(example_deal())
    first, second = docx_text(out["Annex_1_1-2026-1705.docx"]), docx_text(out["Annex_2_1-2026-1705.docx"])
    assert "ПРИЛОЖЕНИЕ № 1" in first and "Изделие X / Item X | 600 | 0,45 | 270,00" in first
    assert "ПРИЛОЖЕНИЕ № 2" in second and "от «01» июня 2026 г." in second
    assert "Сто восемьдесят долларов 00 центов" in second


def test_cover_list_numbering_matches_the_batches():
    text = docx_text(render_all(example_deal())["CoverList_1-2026-1705.docx"])
    assert "2. копия соглашения № В-17/2025 от 15.09.2025" in text
    assert "3. копия приложения № 1 от 20.09.2025" in text
    assert "4. копия соглашения № В-21/2025 от 01.11.2025" in text
    assert "5. карточка компании." in text


def test_special_characters_are_escaped_once():
    d = deal(party_b_name_ru="Общество «A&B <Бета>»")
    assert "Общество «A&B <Бета>»" in _agreement(d)
    assert build_render_context(d)["party_b_name_ru"] == "Общество «A&amp;B &lt;Бета&gt;»"


def test_unknown_template_variable_is_an_error(tmp_path):
    doc = Document(str(TEMPLATES_DIR / TEMPLATE_FILES["annex"]))
    doc.add_paragraph("{{party_b_nmae_ru}}")
    broken = tmp_path / "broken.docx"
    doc.save(str(broken))
    with pytest.raises(ValueError, match="party_b_nmae_ru"):
        render_with_context(broken, build_render_context(deal()))


def test_nested_typo_is_an_error_too(tmp_path):
    doc = Document(str(TEMPLATES_DIR / TEMPLATE_FILES["cover_list"]))
    doc.add_paragraph("{% for b in batches %}{{b.agreemnt_num}}{% endfor %}")
    broken = tmp_path / "broken.docx"
    doc.save(str(broken))
    with pytest.raises(ValueError, match="agreemnt_num"):
        render_with_context(broken, build_render_context(deal()))


def test_annex_numbers_that_collapse_into_one_file_name_are_refused():
    d = deal(annexes=[AnnexLine(annex_num="A/B", annex_date=date(2026, 5, 17), quantity_kg=500),
                      AnnexLine(annex_num="A:B", annex_date=date(2026, 5, 18), quantity_kg=500)])
    with pytest.raises(ValueError, match="same file name"):
        render_all(d)


def test_transfer_act_lays_boxes_out_line_by_line():
    [(name, data)] = render_transfer(example_deal()).items()
    text = docx_text(data)
    assert name == "TransferAct_1-2026-1705.docx"
    assert "АКТ ПЕРЕДАЧИ № 00017 от 25.05.2026" in text
    assert "1. Передал / Handed over by | Компания «Альфа»" in text
    assert "7. Основание / Basis | прил. №1 от 17.05.2026, №2 от 01.06.2026" in text


def test_transfer_text_that_does_not_fit_is_refused_with_every_problem():
    t = example_deal().transfer.model_copy(update={"place_from_ru": "очень длинное место " * 5,
                                                   "note_ru": "строка\n" * 4})
    with pytest.raises(ValueError) as e:
        render_transfer(deal(transfer=t))
    assert "box 5 (place from)" in str(e.value) and "box 4 (note)" in str(e.value)


def test_transfer_needs_a_filled_sheet():
    with pytest.raises(ValueError, match="«Передача» is incomplete"):
        render_transfer(deal(transfer={"transfer_num": "1"}))


def test_amendment_numbers_only_the_clauses_that_change():
    text = docx_text(render_amendment(example_deal()))
    assert "2. Количество: было 1 000 кг, стало 1 200 кг." in text
    assert "Цена:" not in text
    assert "3. Общая сумма: было 450,00, стало 540,00 (Пятьсот сорок долларов США 00 центов)." in text
    assert "4. Остальные условия соглашения не меняются." in text


def test_amendment_with_price_and_quantity():
    a = example_deal().amendment.model_copy(update={"new_price_per_unit": 0.5})
    text = docx_text(render_amendment(deal(amendment=a)))
    assert "2. Цена: было 0,45 доллара США за 1 кг, стало 0,5 доллара США за 1 кг." in text
    assert "3. Количество:" in text and "4. Общая сумма: было 450,00, стало 600,00" in text
