"""Safe writes of Excel workbooks (audit 2026-10-28 A05, red team R04/R09).

Every workbook DCA writes over an operator's file goes through here:
  1. refuse if Excel has the file open (its «~$» owner file is next to it) —
     Excel would later save its old copy over ours;
  2. save into a uniquely named temp file in the same folder;
  3. re-open it — a file that doesn't load is never put in place;
  4. copy the current file into a versions folder (last `keep` kept);
  5. os.replace — atomic on the same file system.
A failure at any step leaves the previous file untouched.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
from collections.abc import Callable, Iterable
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook


def excel_owner_files(path: Path) -> list[Path]:
    """Excel's «~$» owner files for `path` that exist (full name; long names lose 2 chars)."""
    names = {"~$" + path.name, "~$" + path.name[2:]}
    return [path.with_name(n) for n in sorted(names) if path.with_name(n).exists()]


def excel_running() -> bool:
    """Is Microsoft Excel running? Outside macOS (or if unsure) answers True — stay careful."""
    if sys.platform != "darwin":
        return True
    try:
        result = subprocess.run(["pgrep", "-x", "Microsoft Excel"], capture_output=True, timeout=5)
    except (OSError, subprocess.SubprocessError):
        return True
    return result.returncode == 0


def open_in_excel(path: Path, removed: list[Path] | None = None) -> bool:
    """True if Excel really has `path` open: its «~$» owner file exists AND Excel is running.

    Excel deletes the owner file when it closes the workbook. A crash or a forced quit
    leaves it behind, and DCA then said «открыт в Excel» although Excel was closed
    (2026-10-29). With Excel not running such a file is a leftover: it is deleted
    (and appended to `removed`), and the workbook counts as closed.
    """
    owners = excel_owner_files(path)
    if not owners:
        return False
    if excel_running():
        return True
    for owner in owners:
        try:
            owner.unlink()
        except OSError:
            return True
        if removed is not None:
            removed.append(owner)
    return False


def refuse_if_open_in_excel(path: Path) -> None:
    if open_in_excel(path):
        raise PermissionError(
            f"{path.name} открыт в Excel — закройте его и повторите команду. Если в Excel этот файл "
            f"не открыт, закройте Excel полностью (⌘Q): оставшийся служебный файл удалится сам."
        )


def save_workbook_atomic(
    wb: Any,
    path: Path,
    versions_dir: Path | None = None,
    keep: int = 10,
    required_sheets: Iterable[str] = (),
    check: Callable[[Any], None] | None = None,
) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    refuse_if_open_in_excel(path)
    fd, tmp_name = tempfile.mkstemp(prefix=f".{path.stem}-", suffix=".tmp.xlsx", dir=path.parent)
    os.close(fd)
    tmp = Path(tmp_name)
    try:
        wb.save(tmp)
        written = load_workbook(tmp)
        missing = [s for s in required_sheets if s not in written.sheetnames]
        if missing:
            raise ValueError(f"{tmp.name}: written workbook has no sheet(s) {', '.join(missing)}")
        if check is not None:
            check(written)
        if path.exists() and versions_dir is not None:
            versions_dir.mkdir(parents=True, exist_ok=True)
            stamp = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
            shutil.copy2(path, versions_dir / f"{path.stem}_{stamp}{path.suffix}")
            for old in sorted(versions_dir.glob(f"{path.stem}_*{path.suffix}"))[:-keep]:
                old.unlink()
        refuse_if_open_in_excel(path)          # opened while we were writing?
        os.replace(tmp, path)
    finally:
        if tmp.exists():
            tmp.unlink()
