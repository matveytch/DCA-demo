"""Invariants that must hold for any deal.

Unlike example-based tests, these sweep many inputs: whatever the operator types,
the printed figures and the document set must stay consistent.
"""

from decimal import ROUND_CEILING, Decimal
from math import ceil

import pytest

from dca.context import build_annex_contexts, build_render_context
from dca.models import AnnexLine
from dca.render import render_annex_set
from tests.helpers import deal as make_deal

PRICES = [0.01, 0.07, 0.1, 0.29, 0.45, 1.005, 2.675, 19.99, 123.456789]
QUANTITIES = [1, 3, 100, 999, 1000, 10_000, 24_750]


@pytest.mark.parametrize("price", PRICES)
@pytest.mark.parametrize("quantity", QUANTITIES)
def test_total_is_printed_price_times_quantity_rounded_up(price, quantity):
    ctx = build_render_context(make_deal(price_per_unit=price, quantity_kg=quantity))
    exact = Decimal(ctx["price_per_unit"]) * quantity          # the price exactly as printed
    assert ctx["total_amount"] == int(exact.to_integral_value(rounding=ROUND_CEILING))
    assert ctx["total_amount"] - exact < 1                       # never more than one dollar up


@pytest.mark.parametrize("price", PRICES)
def test_multi_annex_total_is_the_sum_of_annex_totals(price):
    parts = [137, 263, 600]
    deal = make_deal(price_per_unit=price, annexes=[
        AnnexLine(annex_num=str(i), annex_date="2026-05-17", quantity_kg=w)
        for i, w in enumerate(parts, 1)])
    annex_totals = [c["total_amount"] for c in build_annex_contexts(deal)]
    assert len(annex_totals) == len(parts)
    assert build_render_context(deal)["total_amount"] == sum(annex_totals)
    exact = [Decimal(repr(price)) * w for w in parts]
    assert annex_totals == [ceil(x) for x in exact]


@pytest.mark.parametrize("n", [1, 2, 3, 5])
def test_one_annex_file_per_annex(n):
    quantities = [1000 // n] * (n - 1) + [1000 - (1000 // n) * (n - 1)]
    annexes = [AnnexLine(annex_num=f"{i}", annex_date="2026-05-17", quantity_kg=w)
               for i, w in enumerate(quantities, 1)] if n > 1 else []
    assert len(render_annex_set(make_deal(annexes=annexes))) == n
