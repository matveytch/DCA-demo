"""Deal model: defaults and business rules (docs/requirements.md, AC3–AC5, AC7)."""

from datetime import date

import pytest
from pydantic import ValidationError

from dca.models import AmendmentInfo, AnnexLine, Deal, TransferInfo
from tests.helpers import deal, example_deal, incoming_deal


def test_example_loads_with_its_parts():
    d = example_deal()
    assert (d.deal_kind, d.variant, d.agreement_num) == ("outgoing", "B", "1-2026-1705")
    assert [a.quantity_kg for a in d.resolved_annexes()] == [600, 400]
    assert d.active_batches() == [1, 2]


def test_defaults_annex_and_due_date():
    d = deal(annex_num=None, annex_date=None, due_date=None)
    assert d.annex_num == "1"                         # part of 1-2026-1705 before the first hyphen
    assert d.annex_date == d.agreement_date
    assert d.due_date == date(2026, 7, 17)            # agreement date + 2 months
    assert deal(agreement_num="NOHYPHEN", annex_num=None).annex_num == "NOHYPHEN"


def test_single_annex_comes_from_deal_fields():
    [a] = deal().resolved_annexes()
    assert (a.annex_num, a.annex_date, a.quantity_kg) == ("1", date(2026, 5, 17), 1000)


@pytest.mark.parametrize("field, value, message", [
    ("unit_of_measure", "t", "unit_of_measure 't' пока не поддерживается"),
    ("currency", "EUR", "currency 'EUR' пока не поддерживается"),
])
def test_only_kg_and_usd(field, value, message):
    with pytest.raises(ValidationError, match=message):
        deal(**{field: value})


def test_outgoing_needs_variant_incoming_does_not():
    with pytest.raises(ValidationError, match="requires variant"):
        deal(variant=None)
    assert incoming_deal().variant is None


def test_variant_b_needs_all_of_party_c():
    with pytest.raises(ValidationError, match="party_c_address, party_c_details"):
        deal(variant="B", party_c_address=None, party_c_details=None)
    assert deal(variant="A", party_c_name=None, party_c_address=None, party_c_details=None).variant == "A"


def test_batches_without_gaps():
    with pytest.raises(ValidationError, match="gap detected"):
        deal(batch_key_1=None, batch_use_kg_1=None, batch_key_2="П-002", batch_use_kg_2=1000)


def test_selected_batch_needs_its_quantity():
    with pytest.raises(ValidationError, match="batch_use_kg_1 is empty"):
        deal(batch_use_kg_1=None)


def test_batch_quantities_compared_as_printed():
    # 500.3 + 500.2 = 1000.5 exactly, but the documents print 501 + 501 = 1002 against 1001
    with pytest.raises(ValidationError, match="501 \\+ 501 = 1002 kg does NOT match"):
        deal(quantity_kg=1000.5, batch_use_kg_1=500.3, batch_key_2="П-002", batch_use_kg_2=500.2)


def test_annex_quantities_must_sum_and_numbers_be_unique():
    with pytest.raises(ValidationError, match="600 \\+ 300 = 900 kg does NOT match"):
        deal(annexes=[AnnexLine(annex_num="1", annex_date=date(2026, 5, 17), quantity_kg=600),
                      AnnexLine(annex_num="2", annex_date=date(2026, 5, 18), quantity_kg=300)])
    with pytest.raises(ValidationError, match="duplicates: 1"):
        deal(annexes=[AnnexLine(annex_num="1", annex_date=date(2026, 5, 17), quantity_kg=500),
                      AnnexLine(annex_num="1", annex_date=date(2026, 5, 18), quantity_kg=500)])


def test_annex_number_spelled_two_ways_is_one_number():
    composed, decomposed = "й", "й"
    assert AnnexLine(annex_num=decomposed, annex_date=date(2026, 1, 1), quantity_kg=1).annex_num == composed


def test_unknown_field_is_refused():
    with pytest.raises(ValidationError, match="Extra inputs are not permitted"):
        Deal(**{**deal().model_dump(), "colour": "red"})


def test_transfer_problems():
    t = TransferInfo(transfer_num="1", transfer_date=date(2026, 5, 1), place_from_ru="А",
                     place_to_ru="Б", handed_by_ru="Орлов", received_by_ru="Орлов")
    problems = t.problems(date(2026, 5, 17))
    assert "transfer_date 01.05.2026 is before agreement_date 17.05.2026" in problems
    assert "handed_by_ru and received_by_ru must be different people" in problems
    assert TransferInfo().missing()[0] == "transfer_num is required"


def test_amendment_defaults_and_nothing_changes():
    a = AmendmentInfo(amendment_num="1", amendment_date=date(2026, 7, 1), new_quantity_kg=1000.4)
    assert a.old_quantity(1000) == 1000 and a.new_price(0.45) == 0.45
    assert a.quantity_changed(1000)                     # printed: 1000 → 1001
    assert not a.quantity_changed(1000.7)               # printed: 1001 → 1001
    same = AmendmentInfo(amendment_num="1", amendment_date=date(2026, 7, 1), new_quantity_kg=1000)
    assert any("nothing changes" in p for p in same.problems(date(2026, 5, 17), 1000, 0.45))
    early = AmendmentInfo(amendment_num="1", amendment_date=date(2026, 5, 1), new_quantity_kg=1200)
    assert any("is before agreement_date" in p for p in early.problems(date(2026, 5, 17), 1000, 0.45))
