"""Smoke tests — package imports, basic helpers work."""

from datetime import date


def test_package_imports():
    """The dca package itself imports."""
    import dca

    assert dca.__version__ == "0.1.0"


def test_config_paths_resolve():
    """Path constants resolve."""
    from dca import config

    assert config.PROJECT_ROOT.exists()
    assert config.TEMPLATES_DIR.exists()
    assert config.DATA_DIR.exists()


def test_slug_basic():
    """Slug helper handles unsafe chars."""
    from dca.utils.slug import slug

    assert slug("BETA LTD") == "BETA_LTD"
    assert slug("foo/bar:baz") == "foo_bar_baz"
    assert slug("Общество «Бета»") == "Общество_«Бета»"


def test_i18n_date_formats():
    """Date formatting helpers return the expected variants."""
    from dca.utils.i18n import all_date_formats

    d = date(2026, 5, 17)
    formats = all_date_formats(d)
    assert formats["iso"] == "2026-05-17"
    assert formats["ru"] == "17.05.2026"
    assert formats["ru_long"] == "«17» мая 2026 г."
    assert formats["ru_full"] == "17 мая 2026 года"
    assert formats["eng"] == "17 May 2026"
    assert formats["eng_long"] == "May 17, 2026"


def test_declension_by_number():
    """Russian noun form by number: 1 доллар, 2 доллара, 5 долларов, 21 доллар, 11 долларов."""
    from dca.utils.declension import DOLLAR_FORMS, decline_ru, decline_ru_float

    assert [decline_ru(n, DOLLAR_FORMS) for n in (1, 2, 5, 11, 21, 22)] == [
        "доллар", "доллара", "долларов", "долларов", "доллар", "доллара"]
    assert decline_ru_float(0.45, DOLLAR_FORMS) == "доллара"


def test_amount_in_words():
    from dca.utils.declension import total_amount_words_eng, total_amount_words_ru

    assert total_amount_words_ru(450) == "Четыреста пятьдесят долларов США 00 центов"
    assert total_amount_words_eng(450) == "Four hundred fifty US dollars 00 cents"
