"""Form-line layout for fixed-height form rows (dca.utils.textfit)."""

import pytest

from dca.utils.textfit import fit_lines, text_width_pt, wrap_lines


def test_capitals_are_wider_than_lowercase():
    assert text_width_pt("BETA COMPANY", 8) > text_width_pt("beta company", 8)
    assert text_width_pt("ШИРОКИЙ", 8) > text_width_pt("широкий", 8)


def test_short_text_is_one_line():
    assert wrap_lines("City V, 3 Third Street", 250, 8) == ["City V, 3 Third Street"]


def test_wraps_between_words_within_width():
    text = "City V, 12 Third Street, office 5, postal code 000001, building B, entrance 2"
    lines = wrap_lines(text, 150, 8)
    assert len(lines) > 1
    assert " ".join(lines) == text
    assert all(text_width_pt(line, 8) <= 150 for line in lines)


def test_keeps_explicit_newlines_and_drops_blank_lines():
    assert wrap_lines("LINE ONE\n\nline two", 250, 8) == ["LINE ONE", "line two"]


def test_cuts_a_word_longer_than_the_line_with_hyphen():
    lines = wrap_lines("X" * 100, 60, 8)
    assert all(line.endswith("-") for line in lines[:-1])
    assert "".join(line.rstrip("-") for line in lines) == "X" * 100
    assert all(text_width_pt(line, 8) <= 60 for line in lines)


def test_cut_word_never_leaves_a_one_or_two_letter_tail():
    lines = wrap_lines("многофункциональный", 77, 8)
    assert lines == ["многофункциональ-", "ный"]


def test_fit_lines_pads_to_form_rows():
    assert fit_lines("A\nB", label="box", width_pt=250, font_size_pt=8, max_lines=4) == ["A", "B", "", ""]
    assert fit_lines(None, label="box", width_pt=250, font_size_pt=8, max_lines=2) == ["", ""]


def test_fit_lines_overflow_is_an_error_not_a_cut():
    with pytest.raises(ValueError, match=r"Box 2 \(received by\): text needs 3 lines, the form has 2"):
        fit_lines("a\nb\nc", label="Box 2 (received by)", width_pt=250, font_size_pt=8, max_lines=2)
