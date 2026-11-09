"""
Russian declension helpers for the document automation system.

All phrase functions take primitive values (number, unit) and return
the fully formed phrase with correct grammatical forms.

Tested on 50+ edge cases — see test_declension.py.
"""

import re

from num2words import num2words

# ─────────────────────────────────────────────────────────────────
# Core declension rule
# ─────────────────────────────────────────────────────────────────

def decline_ru(n: int, forms: tuple[str, str, str]) -> str:
    """
    Pick the correct grammatical form for a Russian noun after a number.

    forms = (nom_sg, gen_sg, gen_pl)
      nom_sg : 1, 21, 31, ..., 101 (but not 11, 111)
      gen_sg : 2, 3, 4, 22, 23, 24, ...
      gen_pl : 0, 5-20, 25-30, 100, 11-14, etc.

    Examples:
      decline_ru(1,  ('доллар', 'доллара', 'долларов')) → 'доллар'
      decline_ru(2,  ...) → 'доллара'
      decline_ru(5,  ...) → 'долларов'
      decline_ru(11, ...) → 'долларов'  (special: teens always gen_pl)
      decline_ru(21, ...) → 'доллар'    (ends in 1, not 11)
      decline_ru(1234, ...) → 'доллара' (last two = 34, %10 = 4)
    """
    n = abs(n) % 100
    if 11 <= n <= 19:
        return forms[2]
    last_digit = n % 10
    if last_digit == 1:
        return forms[0]
    if 2 <= last_digit <= 4:
        return forms[1]
    return forms[2]


_FP_TOLERANCE = 1e-6  # floating-point cleanup threshold


def _is_effectively_integer(x: float) -> bool:
    """Handle FP imprecision: 2.01*1000 ≈ 2010.0000000000005 should be treated as 2010."""
    return abs(x - round(x)) < _FP_TOLERANCE


def decline_ru_float(amount: float, forms: tuple[str, str, str]) -> str:
    """
    Like decline_ru but handles floats. Fractional numbers always use gen_sg
    (Russian rule: «1,5 доллара», «0,5 тонны»).
    FP-tolerant: 2.01*1000 = 2010.0000000000005 → treated as 2010.
    """
    if not _is_effectively_integer(amount):
        return forms[1]  # gen_sg for any fractional
    return decline_ru(round(amount), forms)


def format_ru(x: float) -> str:
    """Format number with Russian decimal style (1234,56 → '1234,56')."""
    if _is_effectively_integer(x):
        return str(round(x))
    return ("%g" % x).replace(".", ",")


def format_en(x: float) -> str:
    """Format number with EN decimal style."""
    if _is_effectively_integer(x):
        return str(round(x))
    return "%g" % x


# ─────────────────────────────────────────────────────────────────
# Price phrases
# ─────────────────────────────────────────────────────────────────

DOLLAR_FORMS = ("доллар", "доллара", "долларов")


def price_per_ton_phrase_ru(price: float) -> str:
    """1000 → '1000 долларов США за 1 тонну'"""
    n = format_ru(price)
    word = decline_ru_float(price, DOLLAR_FORMS)
    return f"{n} {word} США за 1 тонну"


def price_per_kg_phrase_ru(price: float) -> str:
    n = format_ru(price)
    word = decline_ru_float(price, DOLLAR_FORMS)
    return f"{n} {word} США за 1 кг"


def price_per_ton_phrase_eng(price: float) -> str:
    return f"{format_en(price)} USD per 1 ton"


def price_per_kg_phrase_eng(price: float) -> str:
    return f"{format_en(price)} USD per 1 kg"


# ─────────────────────────────────────────────────────────────────
# Quantity phrases
# ─────────────────────────────────────────────────────────────────

TON_FORMS_RU = ("тонна", "тонны", "тонн")


def quantity_phrase_ru(qty: float, unit: str) -> str:
    """
    312 kg → '312 кг'
    2 t  → '2 тонны'
    5 t  → '5 тонн'
    0.5 t → '0,5 тонны'
    """
    n = format_ru(qty)
    if unit == "kg":
        return f"{n} кг"
    if unit == "t":
        word = decline_ru_float(qty, TON_FORMS_RU)
        return f"{n} {word}"
    return f"{n} {unit}"


def quantity_phrase_eng(qty: float, unit: str) -> str:
    n = format_en(qty)
    if unit == "kg":
        return f"{n} kg"
    if unit == "t":
        word = "ton" if qty == 1 else "tons"
        return f"{n} {word}"
    return f"{n} {unit}"


def quantity_kg_phrase_ru(weight_kg: float) -> str:
    return f"{format_ru(weight_kg)} кг"


def quantity_kg_phrase_eng(weight_kg: float) -> str:
    return f"{format_en(weight_kg)} kg"


def quantity_tons_phrase_ru(weight_tons: float) -> str:
    word = decline_ru_float(weight_tons, TON_FORMS_RU)
    return f"{format_ru(weight_tons)} {word}"


def quantity_tons_phrase_eng(weight_tons: float) -> str:
    word = "ton" if weight_tons == 1 else "tons"
    return f"{format_en(weight_tons)} {word}"


# ─────────────────────────────────────────────────────────────────
# Amount-in-words
# ─────────────────────────────────────────────────────────────────

def _ensure_float(x) -> float:
    """num2words integer = cents (bug); always pass float for dollars."""
    return float(x)


# ─────────────────────────────────────────────────────────────────
# Atomic amount pieces — for templates that want
# «<digit> (<words>) <currency> США <cents>» pattern
# ─────────────────────────────────────────────────────────────────

def total_amount_str(amount: float) -> str:
    """`1234.56` → `'1 234.56'` — thousands separated by NBSP, EN decimal.

    NBSP (U+00A0) prevents line wrap inside a money figure.
    """
    return f"{amount:,.2f}".replace(",", " ")


def total_amount_str_ru(amount: float) -> str:
    """`1234.56` → `'1 234,56'` — thousands separated by NBSP, RU comma decimal.

    Step 1: format with comma-as-thousands → `'1,234.56'`
    Step 2: comma → NBSP                  → `'1 234.56'`
    Step 3: dot → comma                    → `'1 234,56'`
    Order matters: comma must be replaced before dot, otherwise the dot becomes a
    comma and we lose the ability to distinguish thousands separator from decimal.
    """
    return f"{amount:,.2f}".replace(",", " ").replace(".", ",")


def total_amount_number_words_ru(amount: float) -> str:
    """
    Spelled-out cardinal of the INTEGER part only, capitalized, no currency.
    1234.56 → 'Одна тысяча двести тридцать четыре'
    1000    → 'Одна тысяча'
    1       → 'Один'
    """
    n = int(amount)  # drop cents
    w = num2words(n, lang="ru")
    return w[0].upper() + w[1:]


def total_amount_number_words_eng(amount: float) -> str:
    """1234.56 → 'One thousand two hundred thirty-four' (no «and», no currency)."""
    n = int(amount)
    w = num2words(n, lang="en").replace(", ", " ").replace(" and ", " ")
    return w[0].upper() + w[1:]


def currency_word_ru_for_total(amount: float) -> str:
    """
    Russian form of «доллар» that agrees with the WHOLE-DOLLAR part of the amount.
    1.00 → 'доллар', 2.00 → 'доллара', 5.00 → 'долларов', 1234.56 → 'доллара'
    """
    n = int(amount)  # whole dollars only — cents handled separately
    return decline_ru(n, DOLLAR_FORMS)


def currency_word_eng_for_total(amount: float) -> str:
    """English: 'dollar' if whole-part is 1, else 'dollars'."""
    return "dollar" if int(amount) == 1 else "dollars"


CENT_FORMS = ("цент", "цента", "центов")


def cents_phrase_ru(amount: float) -> str:
    """Cents portion with grammar. 1.56 → '56 центов', 1.21 → '21 цент', 1.00 → '00 центов'."""
    n = int(round(amount * 100)) % 100
    form = decline_ru(n, CENT_FORMS)
    return f"{n:02d} {form}"


def cents_phrase_eng(amount: float) -> str:
    """1.56 → '56 cents', 1.01 → '01 cent', 1.00 → '00 cents'."""
    n = int(round(amount * 100)) % 100
    word = "cent" if n == 1 else "cents"
    return f"{n:02d} {word}"


def total_amount_words_ru(amount: float) -> str:
    """
    1234.56 → 'Одна тысяча двести тридцать четыре доллара США 56 центов'

    Internals:
      - num2words returns: 'одна тысяча двести тридцать четыре доллара, пятьдесят шесть центов'
      - We rewrite cents as digits + grammatical form: '56 центов' / '01 цент'
      - We insert «США» after the dollar word
      - We capitalize first letter
    """
    raw = num2words(_ensure_float(amount), lang="ru", to="currency", currency="USD")
    # raw ≈ 'одна тысяча двести тридцать четыре доллара, пятьдесят шесть центов'

    # Split on dollar/cent boundary
    # Pattern: «<words> доллар(а|ов)?, <words> цент(а|ов)?»
    m = re.match(r"^(.*?)\s+(доллар[ао]?в?)\s*,\s*(.*?)\s+(цент[ао]?в?)\s*$", raw)
    if not m:
        return raw[0].upper() + raw[1:]  # fallback — return as-is

    dollar_words = m.group(1)
    dollar_noun = m.group(2)

    # Use digits + grammatical form for cents (legal style: "56 центов" not "пятьдесят шесть центов")
    n = int(round(amount * 100)) % 100
    cent_form = decline_ru(n, ("цент", "цента", "центов"))
    result = f"{dollar_words} {dollar_noun} США {n:02d} {cent_form}"
    return result[0].upper() + result[1:]


def total_amount_words_eng(amount: float) -> str:
    """
    1234.56 → 'One thousand two hundred thirty-four US dollars 56 cents'
    1.00    → 'One US dollar 00 cents'  (singular dollar)
    1.01    → 'One US dollar 01 cent'   (singular cent)
    2.00    → 'Two US dollars 00 cents'
    """
    raw = num2words(_ensure_float(amount), lang="en", to="currency", currency="USD")
    cleaned = raw.replace(", ", " ").replace(" and ", " ")
    # Capture singular/plural of «dollar» from num2words
    m = re.match(r"^(.+?)\s+(dollars?)\s+(.+?)\s+(cents?)$", cleaned)
    if not m:
        return cleaned[0].upper() + cleaned[1:]

    words_part = m.group(1)
    dollar_word = m.group(2)  # 'dollar' or 'dollars'
    n_cents = int(round(amount * 100)) % 100
    cent_word = "cent" if n_cents == 1 else "cents"
    result = f"{words_part} US {dollar_word} {n_cents:02d} {cent_word}"
    return result[0].upper() + result[1:]
