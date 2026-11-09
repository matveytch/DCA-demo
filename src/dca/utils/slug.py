"""Filename slugging — make safe folder/file names from arbitrary text."""

import re

# Chars that are unsafe in filenames on any major OS
_UNSAFE_CHARS = re.compile(r'[/\\:?*<>|"\x00-\x1f]')
_MULTI_SPACE = re.compile(r"\s+")


def slug(text: str, max_len: int = 60) -> str:
    """Convert text to a safe filename component.

    - Removes/replaces unsafe chars
    - Collapses whitespace runs to single underscore
    - Trims leading/trailing whitespace and dots
    - Truncates to max_len

    Preserves Cyrillic — modern OSes handle Unicode fine.
    """
    s = text.strip()
    s = _UNSAFE_CHARS.sub("_", s)
    s = _MULTI_SPACE.sub("_", s)
    s = s.strip(" .")
    if len(s) > max_len:
        s = s[:max_len].rstrip(" .")
    return s
