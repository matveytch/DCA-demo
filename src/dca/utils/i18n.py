"""Date formatting helpers (RU/EN) — month names, long formats."""

from datetime import date

# Russian month names in genitive case, lower case as Russian spelling requires
# («17» мая 2026 г.); they used to be capitalised (changed 2026-10-28 on request).
_RU_MONTHS_GEN = [
    "января", "февраля", "марта", "апреля", "мая", "июня",
    "июля", "августа", "сентября", "октября", "ноября", "декабря",
]

# English month names (capitalized)
_EN_MONTHS = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]


def fmt_iso(d: date) -> str:
    """`2026-05-17`."""
    return d.isoformat()


def fmt_ru(d: date) -> str:
    """Short RU: `17.05.2026`."""
    return d.strftime("%d.%m.%Y")


def fmt_ru_long(d: date) -> str:
    """Long RU with quotes: `«17» мая 2026 г.`"""
    month_word = _RU_MONTHS_GEN[d.month - 1]
    return f"«{d.day:02d}» {month_word} {d.year} г."


def fmt_ru_full(d: date) -> str:
    """Full RU without quotes: `17 мая 2026 года`"""
    month_word = _RU_MONTHS_GEN[d.month - 1]
    return f"{d.day} {month_word} {d.year} года"


def fmt_ru_text(d: date) -> str:
    """RU mirror of EN style: `17 мая 2026` — word month, no «г.»/«года» suffix.

    Use when RU/EN cells should render visually parallel (e.g. a transfer date).
    """
    month_word = _RU_MONTHS_GEN[d.month - 1]
    return f"{d.day} {month_word} {d.year}"


def fmt_eng(d: date) -> str:
    """Short EN: `17 May 2026`."""
    return f"{d.day} {_EN_MONTHS[d.month - 1]} {d.year}"


def fmt_eng_long(d: date) -> str:
    """US long EN: `May 17, 2026`."""
    return f"{_EN_MONTHS[d.month - 1]} {d.day}, {d.year}"


def all_date_formats(d: date) -> dict[str, str]:
    """All 7 format variants as dict keyed by suffix."""
    return {
        "iso":      fmt_iso(d),
        "ru":       fmt_ru(d),
        "ru_long":  fmt_ru_long(d),
        "ru_full":  fmt_ru_full(d),
        "ru_text":  fmt_ru_text(d),
        "eng":      fmt_eng(d),
        "eng_long": fmt_eng_long(d),
    }
