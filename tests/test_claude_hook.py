"""The Claude Code hook that keeps sessions inside this repository (.claude/hooks/stay_in_repo.py)."""
from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parents[1]
HOOK = REPO / ".claude" / "hooks" / "stay_in_repo.py"
OUTSIDE = str(REPO.parent / "Data" / "master_db.xlsx")


def _run(tool: str, params: dict, cwd: Path = REPO) -> dict | None:
    event = {"tool_name": tool, "tool_input": params, "cwd": str(cwd)}
    result = subprocess.run(
        [sys.executable, str(HOOK)], input=json.dumps(event),
        capture_output=True, text=True, check=True,
    )
    return json.loads(result.stdout) if result.stdout.strip() else None


def _denied(tool: str, params: dict, cwd: Path = REPO) -> bool:
    out = _run(tool, params, cwd)
    return bool(out) and out["hookSpecificOutput"]["permissionDecision"] == "deny"


@pytest.mark.parametrize("tool, params", [
    ("Read", {"file_path": str(REPO / "CLAUDE.md")}),
    ("Edit", {"file_path": str(REPO / "src" / "dca" / "cli.py")}),
    ("Write", {"file_path": "tests/new_test.py"}),
    ("Glob", {"pattern": "**/*.py"}),
    ("Glob", {"pattern": "src/**/*.py", "path": str(REPO)}),
    ("Grep", {"pattern": "def main"}),
    ("Bash", {"command": "python3 -m pytest -q"}),
    ("Bash", {"command": f"git -C {REPO} status && git log --oneline -5"}),
    ("Bash", {"command": "ls src/../tests > /dev/null 2>&1"}),
    ("Bash", {"command": "git diff HEAD..main --numstat"}),
    ("Bash", {"command": "git push origin main  # https://github.com/x/y"}),
    ("Bash", {"command": f"cat {REPO}/README.md | head -5"}),
    ("Write", {"file_path": f"/tmp/claude-{os.getuid()}/scratch/x.txt"}),
    ("WebSearch", {"query": "anything"}),
])
def test_inside_repo_is_allowed(tool, params):
    assert _run(tool, params) is None


@pytest.mark.parametrize("tool, params", [
    ("Read", {"file_path": OUTSIDE}),
    ("Read", {"file_path": "../Data/master_db.xlsx"}),
    ("Write", {"file_path": str(Path.home() / "Desktop" / "x.txt")}),
    ("Edit", {"file_path": str(REPO.parent / "CLAUDE.md")}),
    ("Glob", {"pattern": "*", "path": str(REPO.parent)}),
    ("Glob", {"pattern": "../**/*.docx"}),
    ("Glob", {"pattern": "/home/*/Downloads/*"}),
    ("Grep", {"pattern": "x", "path": str(Path.home())}),
    ("Bash", {"command": f"ls {REPO.parent}"}),
    ("Bash", {"command": "ls .."}),
    ("Bash", {"command": "cd .. && ls"}),
    ("Bash", {"command": "cat ../Data/batch_ledger.xlsx"}),
    ("Bash", {"command": "ls ~/Downloads"}),
    ("Bash", {"command": "ls $HOME/Desktop"}),
    ("Bash", {"command": "find / -name '*.docx'"}),
    ("Bash", {"command": "python3 x.py --out=/home/someone/file.txt"}),
    ("Bash", {"command": 'echo hi > "/home/someone/file.txt"'}),
    ("Bash", {"command": "git -C .. status"}),
])
def test_outside_repo_is_denied(tool, params):
    assert _denied(tool, params)


def test_bash_from_outside_cwd_is_denied():
    assert _denied("Bash", {"command": "ls"}, cwd=REPO.parent)


def test_unreadable_input_fails_closed():
    result = subprocess.run(
        [sys.executable, str(HOOK)], input="not json",
        capture_output=True, text=True, check=True,
    )
    assert json.loads(result.stdout)["hookSpecificOutput"]["permissionDecision"] == "deny"


def test_settings_register_the_hook():
    settings = json.loads((REPO / ".claude" / "settings.json").read_text(encoding="utf-8"))
    entry = settings["hooks"]["PreToolUse"][0]
    for tool in ("Read", "Write", "Edit", "Glob", "Grep", "Bash"):
        assert tool in entry["matcher"].split("|")
    assert "stay_in_repo.py" in entry["hooks"][0]["command"]
