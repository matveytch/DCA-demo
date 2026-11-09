#!/usr/bin/env python3
"""PreToolUse hook: keep Claude Code sessions inside this repository.

Claude Code passes the tool call as JSON on stdin. The hook denies any file
tool (Read, Write, Edit, Glob, Grep, ...) whose path resolves outside the
repository, and any Bash command that names such a path: absolute paths,
`~`, `$HOME`, or `..` that climbs out of the repository.

Allowed outside the repository: /dev/null & co, Claude Code's own scratchpad
(/tmp/claude-<uid>/) and this project's Claude Code folder
(~/.claude/projects/<project>/: memory and saved tool results).

The Bash check is a text heuristic. It catches paths written in the command,
not paths a program computes at run time.
"""
from __future__ import annotations

import json
import os
import re
import sys
from pathlib import Path

ROOT = os.path.realpath(Path(__file__).resolve().parents[2])

ALLOWED_FILES = {"/dev/null", "/dev/stdin", "/dev/stdout", "/dev/stderr", "/dev/tty"}

# Claude Code names a project's folder after its path with non-alphanumerics replaced by "-".
_PROJECT_KEY = re.sub(r"[^A-Za-z0-9]", "-", str(Path(__file__).resolve().parents[2]))
ALLOWED_DIRS = [
    os.path.realpath(f"/tmp/claude-{os.getuid()}"),
    os.path.realpath(os.path.expanduser(f"~/.claude/projects/{_PROJECT_KEY}")),
]

FILE_TOOLS = {
    "Read": "file_path",
    "Write": "file_path",
    "Edit": "file_path",
    "MultiEdit": "file_path",
    "NotebookEdit": "notebook_path",
    "Glob": "path",
    "Grep": "path",
}

# Bash words: split on whitespace, quotes, shell punctuation and "=" (for --opt=/path).
_WORD = re.compile(r"[^\s'\"<>|;&()`=,]+")
_GLOB_CHARS = re.compile(r"[*?\[{]")


def _inside(path: str) -> bool:
    real = os.path.realpath(path)
    if real in ALLOWED_FILES or path in ALLOWED_FILES:
        return True
    for base in [ROOT, *ALLOWED_DIRS]:
        if real == base or real.startswith(base + os.sep):
            return True
    return False


def _resolve(raw: str, cwd: str) -> str:
    path = os.path.expanduser(os.path.expandvars(raw))
    # Only the part before the first wildcard is a real location.
    match = _GLOB_CHARS.search(path)
    if match:
        path = path[: match.start()] or "."
    return path if os.path.isabs(path) else os.path.join(cwd, path)


def _is_path_word(word: str) -> bool:
    if "://" in word:
        return False
    expanded = os.path.expanduser(os.path.expandvars(word))
    if os.path.isabs(expanded):
        return True
    return ".." in expanded.split("/")


def _bash_outside(command: str, cwd: str) -> list[str]:
    if not _inside(cwd):
        return [cwd]
    return [
        word
        for word in _WORD.findall(command)
        if _is_path_word(word) and not _inside(_resolve(word, cwd))
    ]


def check(event: dict) -> str | None:
    """Return a deny reason, or None when the call stays inside the repository."""
    tool = event.get("tool_name", "")
    params = event.get("tool_input") or {}
    cwd = event.get("cwd") or ROOT

    if tool in FILE_TOOLS:
        raw = params.get(FILE_TOOLS[tool]) or cwd
        outside = [] if _inside(_resolve(raw, cwd)) else [raw]
        pattern = params.get("pattern", "") if tool == "Glob" else ""
        if pattern and _is_path_word(pattern) and not _inside(_resolve(pattern, _resolve(raw, cwd))):
            outside.append(pattern)
    elif tool == "Bash":
        outside = _bash_outside(params.get("command", ""), cwd)
    else:
        return None

    if not outside:
        return None
    return (
        f"Blocked by .claude/hooks/stay_in_repo.py: sessions work only inside {ROOT}. "
        f"Outside path(s): {', '.join(outside)}. Ask the user to do this themselves "
        "or to start a session in the needed folder."
    )


def main() -> int:
    try:
        reason = check(json.load(sys.stdin))
    except Exception as exc:  # fail closed: an unreadable call is not allowed through
        reason = f"stay_in_repo.py could not check this call: {exc!r}"
    if reason:
        json.dump(
            {
                "hookSpecificOutput": {
                    "hookEventName": "PreToolUse",
                    "permissionDecision": "deny",
                    "permissionDecisionReason": reason,
                }
            },
            sys.stdout,
        )
    return 0


if __name__ == "__main__":
    sys.exit(main())
