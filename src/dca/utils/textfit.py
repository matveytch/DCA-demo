"""Lay a text block out over the fixed-height lines of a form (the transfer act).

Form rows have an exact height of one line, so Word clips (does not wrap) text that
doesn't fit. Instead the engine breaks each block into lines itself and puts every
line into its own row variable. Widths are measured in points with Arial metrics
(Helvetica AFM widths for ASCII, conservative estimates for Cyrillic), so a line of
capitals is correctly treated as wider than one of lowercase letters.

Deterministic: same text → same lines. If a block needs more lines than the form
has, `fit_lines` raises instead of silently cutting text.
"""

from __future__ import annotations

# Helvetica/Arial advance widths, 1/1000 em, for printable ASCII 32..126.
_ASCII_WIDTHS = [
    278, 278, 355, 556, 556, 889, 667, 191, 333, 333, 389, 584, 278, 333, 278, 278,  # ' '../
    556, 556, 556, 556, 556, 556, 556, 556, 556, 556,                                # 0..9
    278, 278, 584, 584, 584, 556, 1015,                                              # :..@
    667, 667, 722, 722, 667, 611, 778, 722, 278, 500, 667, 556, 833,                 # A..M
    722, 778, 667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611,                 # N..Z
    278, 278, 278, 469, 556, 333,                                                    # [..`
    556, 556, 500, 556, 556, 278, 556, 556, 222, 222, 500, 222, 833,                 # a..m
    556, 556, 556, 556, 333, 500, 278, 556, 500, 722, 500, 500, 500,                 # n..z
    334, 260, 334, 584,                                                              # {..~
]
_WIDE_CYR_UPPER = set("ЖШЩЮЫМФ")
_WIDE_CYR_LOWER = set("жшщюыф")
_OTHER = {"«": 556, "»": 556, "№": 1073, "—": 1000, "–": 556, "“": 333, "”": 333, " ": 278}
_DEFAULT = 667  # anything else: assume a wide-ish glyph (errs towards wrapping earlier)
_MIN_TAIL = 3   # a cut word leaves at least this many letters for the next line


def char_width(ch: str) -> int:
    """Advance width of one character in 1/1000 em."""
    o = ord(ch)
    if 32 <= o <= 126:
        return _ASCII_WIDTHS[o - 32]
    if ch in _OTHER:
        return _OTHER[ch]
    if "А" <= ch <= "Я" or ch == "Ё":
        return 944 if ch in _WIDE_CYR_UPPER else 722
    if "а" <= ch <= "я" or ch == "ё":
        return 778 if ch in _WIDE_CYR_LOWER else 556
    return _DEFAULT


def text_width_pt(text: str, font_size_pt: float) -> float:
    return sum(char_width(c) for c in text) * font_size_pt / 1000


def wrap_lines(text: str, width_pt: float, font_size_pt: float) -> list[str]:
    """Break `text` into lines no wider than `width_pt`.

    Explicit newlines (Alt+Enter in Excel) are kept; within a line, breaks go
    between words; a single word wider than the line is cut with a hyphen.
    Blank lines are dropped (they would waste a form row).
    """
    lines: list[str] = []
    for paragraph in text.splitlines():
        current = ""
        for word in paragraph.split():
            candidate = f"{current} {word}" if current else word
            if text_width_pt(candidate, font_size_pt) <= width_pt:
                current = candidate
                continue
            if current:
                lines.append(current)
            current = ""
            while text_width_pt(word, font_size_pt) > width_pt:
                cut = _longest_prefix(word, width_pt - text_width_pt("-", font_size_pt), font_size_pt)
                if len(word) - cut < _MIN_TAIL and cut > _MIN_TAIL:
                    cut = len(word) - _MIN_TAIL  # no lone 1-2 letter tail on the next line
                lines.append(word[:cut] + "-")
                word = word[cut:]
            current = word
        if current:
            lines.append(current)
    return lines


def fit_lines(
    text: str | None, *, label: str, width_pt: float, font_size_pt: float, max_lines: int
) -> list[str]:
    """`wrap_lines`, padded with "" to exactly `max_lines`; ValueError if it doesn't fit."""
    lines = wrap_lines(text or "", width_pt, font_size_pt)
    if len(lines) > max_lines:
        raise ValueError(
            f"{label}: text needs {len(lines)} lines, the form has {max_lines}. "
            f"Shorten it. Overflowing part: {' / '.join(lines[max_lines:])!r}"
        )
    return lines + [""] * (max_lines - len(lines))


def _longest_prefix(word: str, width_pt: float, font_size_pt: float) -> int:
    n = 1
    while n < len(word) and text_width_pt(word[: n + 1], font_size_pt) <= width_pt:
        n += 1
    return n
