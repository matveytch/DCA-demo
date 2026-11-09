"""Build the flat template-rendering context from a Deal.

Combines:
  - All source fields from Deal (operator-entered values)
  - Computed dates in 7 format variants per source date
  - Computed amount components (str, str_ru, words, currency_word, cents)
  - Computed price/quantity phrases (RU/EN, declined)

This dict is what gets passed to docxtpl.render().

Business rules — rounding for rendered values:
  - total_amount and quantity_kg are rounded UP to whole units (math.ceil):
    21333.04 → 21334.
  - Display formats stay unchanged (".00", "00 центов", etc.) —
    only the underlying numeric value changes, not how it's printed.
"""

from __future__ import annotations

import math
from datetime import date
from decimal import ROUND_CEILING, Decimal
from typing import Any
from xml.sax.saxutils import escape as _xml_escape

from dca.models import BATCH_SLOTS, Deal
from dca.utils import declension as decl
from dca.utils import i18n


def build_render_context(deal: Deal) -> dict[str, Any]:
    """Produce the flat dict that docxtpl will render with (strings XML-escaped)."""
    return _escape_strings(_raw_render_context(deal))


def _escape_strings(ctx: dict[str, Any]) -> dict[str, Any]:
    """XML-escape every string value — ONCE, at the rendering boundary.

    docxtpl does NOT auto-escape — an un-escaped '&' is parsed as an XML entity
    start and the text after it is lost ('A&B' → 'A'). All contexts (deal, annex,
    transfer, amendment) are built raw and escaped here, so nothing is escaped
    twice and nothing added later is left raw.
    """
    return {k: _escape_value(v) for k, v in ctx.items()}


def _escape_value(v: Any) -> Any:
    """Strings escaped; lists of row dicts (annex_rows, batches) escaped inside too."""
    if isinstance(v, str):
        return _xml_escape(v)
    if isinstance(v, list):
        return [{k: _escape_value(x) for k, x in row.items()} if isinstance(row, dict) else _escape_value(row)
                for row in v]
    return v


# Legal-form words that start a Russian company name → their genitive. Only the
# leading form changes («Общество с ограниченной ответственностью «X»» →
# «Общества с ограниченной ответственностью «X»»); abbreviations (ООО, АО, ИП)
# and unknown words are left as they are.
_ORG_FORM_GENITIVE = {
    "Общество": "Общества",
    "Компания": "Компании",
    "Корпорация": "Корпорации",
    "Предприятие": "Предприятия",
    "Фирма": "Фирмы",
    "Учреждение": "Учреждения",
    "Акционерное общество": "Акционерного общества",
    "Индивидуальный предприниматель": "Индивидуального предпринимателя",
}


def org_name_genitive_ru(name: str) -> str:
    """`Компания «Альфа»` → `Компании «Альфа»` (for «Директор …»)."""
    for form in sorted(_ORG_FORM_GENITIVE, key=len, reverse=True):
        for spelled in (form, form.lower(), form.upper()):
            if name.startswith(spelled + " ") or name == spelled:
                genitive = _ORG_FORM_GENITIVE[form]
                if spelled.isupper():
                    genitive = genitive.upper()
                elif spelled.islower():
                    genitive = genitive.lower()
                return genitive + name[len(spelled):]
    return name


def _kg_str(x: float) -> str:
    """Whole kg rounded UP, NBSP thousands separator: 13453.2 → '13 454'."""
    return f"{math.ceil(x):,}".replace(",", " ")


def _raw_render_context(deal: Deal) -> dict[str, Any]:
    """The deal context before XML-escaping (see build_render_context)."""
    ctx: dict[str, Any] = {}

    # ─── Source fields straight through ─────────────────────────
    ctx.update(deal.model_dump(mode="python"))
    # «Директор Компании «…»» — the company name after «Директор» (genitive).
    ctx["party_a_name_ru_genitive"] = org_name_genitive_ru(deal.party_a_name_ru)

    # ─── Date format variants (7 per named date) ────────────────
    assert deal.annex_date is not None and deal.due_date is not None  # _fill_defaults
    _add_date_variants(ctx, "agreement_date", deal.agreement_date)
    _add_date_variants(ctx, "annex_date", deal.annex_date)
    _add_date_variants(ctx, "due_date", deal.due_date)
    _add_date_variants(ctx, "today_date", date.today())
    _add_date_variants(ctx, "letter_date", deal.letter_date or date.today())

    # ─── Quantity — INTEGER, rounded UP, NBSP thousands separator ───
    ctx["quantity_kg"] = _kg_str(deal.quantity_kg)
    ctx["quantity_kg_phrase_ru"] = decl.quantity_kg_phrase_ru(math.ceil(deal.quantity_kg))
    ctx["quantity_kg_phrase_eng"] = decl.quantity_kg_phrase_eng(math.ceil(deal.quantity_kg))

    # ─── Price — printed EXACTLY as entered (every digit) ───────
    # The reader multiplying the printed quantity by the printed price gets the
    # printed total (format_en's «%g» would keep only 6 significant digits).
    price = _price_per_kg_dec(deal)
    ctx["price_per_unit"] = _fmt_dec(price)
    ctx["price_per_unit_ru"] = _fmt_dec(price, ",")
    ctx.update(_price_phrases(price))

    # ─── Total amount — consistent with the printed quantity ────
    # total = ceil(ceil(quantity_kg) × price). Several annexes: the SUM of the
    # per-annex totals, so the agreement total equals the sum of the annexes.
    annexes = deal.resolved_annexes()
    if len(annexes) > 1:
        total_amount = sum(_total_at_price(a.quantity_kg, price) for a in annexes)
    else:
        total_amount = _total_at_price(deal.quantity_kg, price)
    ctx.update(_amount_ctx(total_amount))
    # One row per annex for `{%tr for a in annex_rows %}` tables in the agreement.
    ctx["annex_rows"] = [
        {"num": a.annex_num, "date": i18n.fmt_ru(a.annex_date), "quantity_kg": _kg_str(a.quantity_kg),
         "total": decl.total_amount_str_ru(_total_at_price(a.quantity_kg, price))}
        for a in annexes
    ]

    # ─── Batches: date variants and quantities per active slot ──
    # Template convention: {base}_{suffix}_{n} (e.g. batch_agreement_date_ru_1).
    for n in BATCH_SLOTS:
        for date_field in _BATCH_DATE_FIELDS:
            d = getattr(deal, f"{date_field}_{n}")
            for suffix in i18n.all_date_formats(date(2000, 1, 1)):
                ctx[f"{date_field}_{suffix}_{n}"] = ""
            if d is not None and getattr(deal, f"batch_key_{n}"):
                for suffix, formatted in i18n.all_date_formats(d).items():
                    ctx[f"{date_field}_{suffix}_{n}"] = formatted
        w = getattr(deal, f"batch_use_kg_{n}")
        if w is not None:
            ctx[f"batch_use_kg_{n}"] = _kg_str(w)
    ctx["batches"] = [_batch_ctx(deal, n) for n in deal.active_batches()]

    if not ctx.get("letter_num"):
        ctx["letter_num"] = "1"

    # Remaining None → "" so Jinja doesn't render the literal "None".
    for k, v in list(ctx.items()):
        if v is None:
            ctx[k] = ""
    return ctx


_BATCH_DATE_FIELDS = ("batch_agreement_date", "batch_annex_date")


def _batch_ctx(deal: Deal, n: int) -> dict[str, Any]:
    """One active batch as a dict for `{% for b in batches %}` loops in templates."""
    def ru(d: date | None) -> str:
        return i18n.fmt_ru(d) if d else ""

    return {
        "n": n,
        "key": getattr(deal, f"batch_key_{n}") or "",
        "agreement_num": getattr(deal, f"batch_agreement_num_{n}") or "",
        "agreement_date": ru(getattr(deal, f"batch_agreement_date_{n}")),
        "annex_num": getattr(deal, f"batch_annex_num_{n}") or "",
        "annex_date": ru(getattr(deal, f"batch_annex_date_{n}")),
        "party_name": getattr(deal, f"batch_party_name_ru_{n}") or "",
        "party_location": getattr(deal, f"batch_party_location_ru_{n}") or "",
        "use_kg": _kg_str(getattr(deal, f"batch_use_kg_{n}") or 0),
    }


# ─── Amount / per-annex helpers ────────────────────────────────

def _dec(x: float) -> Decimal:
    """The number exactly as entered: Excel/JSON floats via their shortest repr
    (0.07 → Decimal('0.07'), never 0.07000000000000000666…)."""
    return Decimal(repr(float(x)))


def _fmt_dec(d: Decimal, decimal_sep: str = ".") -> str:
    """All significant digits, no exponent, no trailing zeros: 123.456789 / 0.45 / 450."""
    s = format(d.normalize(), "f")
    if "." in s:
        s = s.rstrip("0").rstrip(".")
    return s.replace(".", decimal_sep)


def _price_per_kg_dec(deal: Deal) -> Decimal:
    """Price per kg, exact. Only kg is supported for now (models: _check_supported_units)."""
    return _dec(deal.price_per_unit)


def _total_for_quantity(deal: Deal, quantity_kg: float) -> int:
    """total = ceil(shown quantity × price per kg), in exact decimal arithmetic.

    Decimal, not float: 100 × 0.07 must be 7, not ceil(7.000000000000001) = 8.
    """
    return _total_at_price(quantity_kg, _price_per_kg_dec(deal))


def _total_at_price(quantity_kg: float, price_per_kg: Decimal) -> int:
    """ceil(ceil(quantity) × price) in Decimal — the one place a total is rounded."""
    shown = Decimal(math.ceil(quantity_kg))
    return int((shown * price_per_kg).to_integral_value(rounding=ROUND_CEILING))


def _amount_ctx(total_amount: int) -> dict[str, Any]:
    """All `total_amount*` template keys for a given total (str/words/cents/currency)."""
    return {
        "total_amount": total_amount,
        "total_amount_str": decl.total_amount_str(total_amount),
        "total_amount_str_ru": decl.total_amount_str_ru(total_amount),
        "total_amount_number_words_ru": decl.total_amount_number_words_ru(total_amount),
        "total_amount_number_words_eng": decl.total_amount_number_words_eng(total_amount),
        "total_amount_words_ru": decl.total_amount_words_ru(total_amount),
        "total_amount_words_eng": decl.total_amount_words_eng(total_amount),
        "currency_word_ru_for_total": decl.currency_word_ru_for_total(total_amount),
        "currency_word_eng_for_total": decl.currency_word_eng_for_total(total_amount),
        "cents_phrase_ru": decl.cents_phrase_ru(total_amount),
        "cents_phrase_eng": decl.cents_phrase_eng(total_amount),
    }


def _price_phrases(price_per_kg: Decimal) -> dict[str, str]:
    """`price_per_kg_phrase_*` for a price in USD per kg (declined in Russian)."""
    kg, kg_ru = _fmt_dec(price_per_kg), _fmt_dec(price_per_kg, ",")
    return {
        "price_per_kg_phrase_ru": f"{kg_ru} {decl.decline_ru_float(float(price_per_kg), decl.DOLLAR_FORMS)} США за 1 кг",
        "price_per_kg_phrase_eng": f"{kg} USD per 1 kg",
    }


def build_annex_contexts(deal: Deal) -> list[dict[str, Any]]:
    """One render context per annex (always ≥ 1).

    Starts from the aggregate context and, for each annex, overrides the annex
    keys: number, the date variants, the shown quantity and the full total_amount
    block (recomputed from the annex's quantity). Everything else stays shared.
    """
    base = build_render_context(deal)
    contexts: list[dict[str, Any]] = []
    for annex in deal.resolved_annexes():
        ctx = dict(base)
        ctx["quantity_kg"] = _kg_str(annex.quantity_kg)
        ctx["annex_num"] = _xml_escape(annex.annex_num)
        for suffix, formatted in i18n.all_date_formats(annex.annex_date).items():
            ctx[f"annex_date_{suffix}"] = formatted
        ctx.update(_amount_ctx(_total_for_quantity(deal, annex.quantity_kg)))
        contexts.append(ctx)
    return contexts


def build_amendment_context(deal: Deal) -> dict[str, Any]:
    """Context for the amendment: the deal context plus the «Изменение» values —
    old/new quantity, price, totals, words, clause numbers.

    Totals use the agreement's rule (ceil(ceil(quantity) × price), Decimal). Clauses
    that don't change (price, quantity) are left out and the rest numbered without
    gaps. Raises ValueError listing EVERY problem at once.
    """
    a = deal.amendment
    if a is None:
        raise ValueError("the «Изменение» sheet is empty — fill it in the deal's input Excel")
    deal_price = float(_price_per_kg_dec(deal))
    problems = a.problems(deal.agreement_date, deal.quantity_kg, deal_price)
    if problems:
        raise ValueError("«Изменение» is incomplete:\n  - " + "\n  - ".join(problems))
    assert a.amendment_date is not None and a.new_quantity_kg is not None

    ctx = _raw_render_context(deal)
    ctx["amendment_num"] = a.amendment_num
    _add_date_variants(ctx, "amendment_date", a.amendment_date)
    for prefix, quantity, price in (
        ("old_", a.old_quantity(deal.quantity_kg), _dec(a.old_price(deal_price))),
        ("new_", a.new_quantity_kg, _dec(a.new_price(deal_price))),
    ):
        ctx[f"{prefix}quantity_kg"] = _kg_str(quantity)
        ctx.update({prefix + k: v for k, v in _amount_ctx(_total_at_price(quantity, price)).items()})
        ctx.update({prefix + k: v for k, v in _price_phrases(price).items()})

    # Clause 1 (subject) is always there; price and quantity only when they change.
    ctx["price_changed"] = a.price_changed(deal_price)
    ctx["quantity_changed"] = a.quantity_changed(deal.quantity_kg)
    n = 2
    for item, shown in (("item_price", ctx["price_changed"]), ("item_quantity", ctx["quantity_changed"]),
                        ("item_total", True), ("item_rest", True)):
        ctx[item] = n if shown else ""
        n += shown
    return _escape_strings(ctx)


def _add_date_variants(ctx: dict[str, Any], base: str, d: date) -> None:
    """Add the format variants for one date under keys `<base>_<suffix>`."""
    for suffix, formatted in i18n.all_date_formats(d).items():
        ctx[f"{base}_{suffix}"] = formatted


# ─── Transfer act ───────────────────────────────────────────────


def build_transfer_context(deal: Deal) -> dict[str, Any]:
    """Context for the transfer act: the deal context plus the «Передача» fields,
    their date variants and the engine-laid-out box lines.

    Raises ValueError listing EVERY problem at once (missing fields, date logic,
    text that doesn't fit a box) so the operator fixes them in one pass.
    """
    from dca.config import TRANSFER_FIELD_LIMITS, TRANSFER_FONT_SIZE_PT, TRANSFER_LINE_BOXES
    from dca.utils.textfit import fit_lines

    t = deal.transfer
    if t is None:
        raise ValueError("the «Передача» sheet is empty — fill it in the deal's input Excel")
    problems = t.problems(deal.agreement_date)
    if problems:
        raise ValueError("«Передача» is incomplete:\n  - " + "\n  - ".join(problems))

    ctx = _raw_render_context(deal)
    for name, value in t.model_dump().items():
        ctx[name] = "" if value is None else value
    assert t.transfer_date is not None
    _add_date_variants(ctx, "transfer_date", t.transfer_date)
    ctx["transfer_annexes_ref"] = _transfer_annexes_ref(deal)

    # Party blocks are composed from the archive fields: each part starts on a new
    # form line, its own Alt+Enter breaks become ", " so it flows.
    blocks = {
        "sender": _flow(deal.party_a_name_ru, ctx.get("party_a_address_full_ru"), t.handed_by_ru),
        "receiver": _flow(deal.party_b_name_ru, ctx.get("party_b_address_full_ru"), t.received_by_ru),
        "item": f"{deal.item_name_ru} / {deal.item_name_eng}",
        "note": t.note_ru or "",
    }
    errors: list[str] = []
    for box, text in blocks.items():
        label, width, rows = TRANSFER_LINE_BOXES[box]
        try:
            lines = fit_lines(text, label=label, width_pt=width,
                              font_size_pt=TRANSFER_FONT_SIZE_PT, max_lines=rows)
        except ValueError as e:
            errors.append(str(e))
            continue
        for n, line in enumerate(lines, start=1):
            ctx[f"transfer_{box}_line_{n}"] = line
    for key, (label, width, rows) in TRANSFER_FIELD_LIMITS.items():
        try:
            fit_lines(str(ctx.get(key) or ""), label=label, width_pt=width,
                      font_size_pt=TRANSFER_FONT_SIZE_PT, max_lines=rows)
        except ValueError as e:
            errors.append(str(e))
    if errors:
        raise ValueError("Text does not fit the transfer act:\n  - " + "\n  - ".join(errors))
    return _escape_strings(ctx)


def _flow(*parts: str | None) -> str:
    """Join block parts one per line; line breaks inside a part become ', '."""
    out = []
    for part in parts:
        if part:
            pieces = (x.strip().rstrip(",").strip() for x in str(part).splitlines())
            out.append(", ".join(x for x in pieces if x))
    return "\n".join(out)


def _transfer_annexes_ref(deal: Deal) -> str:
    """Box 7 line: the annex(es) of this agreement (the agreement itself is in box 8)."""
    annexes = deal.resolved_annexes()
    if len(annexes) == 1:
        a = annexes[0]
        return f"приложение №{a.annex_num} от {i18n.fmt_ru(a.annex_date)}"
    refs = ", ".join(f"№{a.annex_num} от {i18n.fmt_ru(a.annex_date)}" for a in annexes)
    return f"прил. {refs}"
