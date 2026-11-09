"""Reading the input Excel: archives, batches, annexes, form sheets, operator-facing errors (AC3)."""

from datetime import date

import pytest
from openpyxl import load_workbook

from dca.excel_io import DealInputError, load_deal
from tests.helpers import copy_example, set_annexes, set_archive_value, set_deal_values


@pytest.fixture
def xlsx(tmp_path):
    return copy_example(tmp_path / "deal")


def test_archives_and_batches_are_looked_up_by_key(xlsx):
    d = load_deal(xlsx)
    assert d.party_a_name_ru == "Компания «Альфа»" and d.party_a_folder_name == "Компания А"
    assert d.party_b_name_eng == "Beta Ltd" and d.item_name_ru == "Изделие X"
    assert (d.batch_agreement_num_1, d.batch_agreement_date_1) == ("В-17/2025", date(2025, 9, 15))
    assert d.batch_quantity_kg_1 == 5000 and d.batch_use_kg_2 == 300
    assert d.batch_annex_num_2 is None


def test_form_sheets_are_read(xlsx):
    d = load_deal(xlsx)
    assert d.transfer.transfer_num == "00017" and d.transfer.transfer_date == date(2026, 5, 25)
    assert d.amendment.new_quantity_kg == 1200


def test_the_command_decides_the_kind(xlsx):
    assert load_deal(xlsx, deal_kind="incoming").deal_kind == "incoming"


def test_unknown_archive_key_lists_the_available_ones(xlsx):
    set_deal_values(xlsx, {"party_b_key": "nobody"})
    with pytest.raises(ValueError, match=r"party_b_key='nobody' not found in archive. Available keys: \['beta', 'alpha-b'\]"):
        load_deal(xlsx)


def test_missing_key_is_named_in_russian_context(xlsx):
    set_deal_values(xlsx, {"party_a_key": None})
    with pytest.raises(ValueError, match="party_a_key is required"):
        load_deal(xlsx)


def test_all_problems_in_one_message_with_sheet_and_row(xlsx):
    set_deal_values(xlsx, {"place_ru": None, "price_per_unit": -1})
    set_archive_value(xlsx, "Изделия", "item_name_eng", "X", None)
    with pytest.raises(DealInputError) as e:
        load_deal(xlsx)
    text = str(e.value)
    assert text.startswith("в Excel не хватает данных или они неверные:")
    assert "«Место исполнения (RU)» (place_ru): не заполнено — лист «Сделка», строка" in text
    assert "price_per_unit" in text and "неверное значение" in text
    assert "item_name_eng: не заполнено — лист «Изделия», запись «X»" in text


def test_renamed_variable_in_column_a_is_refused(xlsx):
    wb = load_workbook(xlsx)
    ws = wb["Сделка"]
    row = next(r for r in ws.iter_rows() if r[0].value == "annex_date")
    row[0].value = "Annex_date"
    wb.save(xlsx)
    with pytest.raises(ValueError, match=r"Annex_date \(should be annex_date\)"):
        load_deal(xlsx)


def test_service_field_on_the_sheet_is_refused(xlsx):
    wb = load_workbook(xlsx)
    ws = wb["Сделка"]
    ws.append(["deal_kind", "", "incoming"])
    wb.save(xlsx)
    with pytest.raises(ValueError, match="service field"):
        load_deal(xlsx)


def test_annex_table(xlsx):
    set_annexes(xlsx, [("A-1", "17.05.2026", 250.0), (7, date(2026, 6, 1), 750)])
    annexes = load_deal(xlsx).annexes
    assert [(a.annex_num, a.annex_date, a.quantity_kg) for a in annexes] == [
        ("A-1", date(2026, 5, 17), 250.0), ("7", date(2026, 6, 1), 750.0)]


def test_empty_annex_table_means_one_annex(xlsx):
    set_annexes(xlsx, [])
    d = load_deal(xlsx)
    assert d.annexes == [] and len(d.resolved_annexes()) == 1


def test_annex_data_under_a_broken_header_is_refused(xlsx):
    wb = load_workbook(xlsx)
    wb["Приложения"]["A1"] = "number"
    wb.save(xlsx)
    with pytest.raises(ValueError, match="has data but no header row"):
        load_deal(xlsx)


def test_numbers_typed_as_text_and_dates_as_text(xlsx):
    set_deal_values(xlsx, {"agreement_num": 12345, "agreement_date": "17.05.2026"})
    d = load_deal(xlsx)
    assert d.agreement_num == "12345" and d.agreement_date == date(2026, 5, 17)


def test_unknown_name_on_a_form_sheet_is_refused(xlsx):
    wb = load_workbook(xlsx)
    ws = wb["Передача"]
    for row in ws.iter_rows():
        if row[0].value == "note_ru":
            row[0].value = "notes_ru"
    wb.save(xlsx)
    with pytest.raises(ValueError, match="Unknown variable 'notes_ru' on sheet «Передача»"):
        load_deal(xlsx)


def test_text_number_on_a_form_sheet(xlsx):
    set_deal_values(xlsx, {"new_quantity_kg": "1 200,5"}, sheet="Изменение")
    assert load_deal(xlsx).amendment.new_quantity_kg == 1200.5
