"""Where the code lives vs. where the operator's data lives.

Two roots:

- **code root** — the folder with `src/`, `templates/`, `data/input_template_*.xlsx`
  (this repository). Always derived from this file's location
  (`config.PROJECT_ROOT`). Templates and the empty Excel belong here: their
  variable names must match the code.

- **data root** — the operator's folder (business records), marked by `.dca.toml`
  (or `dca.toml`), the way git marks a repo with `.git`:

      <company folder>/Outgoing/<deal>   ← outgoing deals of that company (party_a_folder_name)
      <company folder>/Incoming/<deal>   ← incoming deals (party_b_folder_name)
      <company folder>/company_profile.pdf
      Новые сделки/<num>_pending         ← dca new / new-incoming, until the first generate
      Data/                              ← master_db.xlsx, batch_ledger.xlsx, scans/

Data-root resolution is deterministic; the first hit wins:

1. explicit `--home <dir>` (CLI option);
2. `DCA_HOME` environment variable;
3. nearest ancestor of `start` (the deal folder, or cwd) that contains a marker
   (if that marker sets `data_dir`, the data root is there instead);
4. `data_dir` key of the code root's marker (code and data kept apart);
5. the code root itself.

Marker keys (both optional; an empty file is a valid marker):

    code_dir = "dca"                 # in the DATA root: where the code lives
    data_dir = "../dca-data"         # in the CODE root: where the data lives

Relative paths are resolved against the folder holding that marker; `~` is expanded.
"""

from __future__ import annotations

import os
import tomllib
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from dca.config import PROJECT_ROOT

MARKER_FILENAME = "dca.toml"
MARKER_FILENAMES = (".dca.toml", MARKER_FILENAME)   # hidden one first: keeps the root folder clean
HOME_ENV_VAR = "DCA_HOME"

DATA_DIR_NAME = "Data"
NEW_DEALS_DIR_NAME = "Новые сделки"
OUTGOING_DIR_NAME = "Outgoing"
INCOMING_DIR_NAME = "Incoming"


@dataclass(frozen=True)
class DataRoot:
    """Resolved data root + how it was found (shown by `dca doctor`)."""

    path: Path
    source: str

    @property
    def data_dir(self) -> Path:
        return self.path / DATA_DIR_NAME

    @property
    def new_deals_dir(self) -> Path:
        """Deals created by `dca new` / `new-incoming`, until generate files them."""
        return self.path / NEW_DEALS_DIR_NAME

    def company_dir(self, folder_name: str) -> Path:
        return self.path / folder_name

    def outgoing_dir(self, folder_name: str) -> Path:
        return self.company_dir(folder_name) / OUTGOING_DIR_NAME

    def incoming_dir(self, folder_name: str) -> Path:
        return self.company_dir(folder_name) / INCOMING_DIR_NAME

    def deal_dirs(self) -> list[Path]:
        """Every deal folder: «Новые сделки»/* and <company>/Outgoing|Incoming/*."""
        parents = [self.new_deals_dir]
        if self.path.is_dir():
            for company in sorted(self.path.iterdir()):
                if company.is_dir() and not company.name.startswith((".", "_")):
                    parents += [company / OUTGOING_DIR_NAME, company / INCOMING_DIR_NAME]
        return [d for p in parents if p.is_dir() for d in sorted(p.iterdir()) if d.is_dir()]

    @property
    def scans_dir(self) -> Path:
        """Operator's archive of batch scans, used by the attachment pack."""
        return self.data_dir / "scans"

    @property
    def ledger_path(self) -> Path:
        return self.data_dir / "batch_ledger.xlsx"

    @property
    def master_db_path(self) -> Path:
        """Master database of parties / items / batches (see dca.master_db)."""
        return self.data_dir / "master_db.xlsx"


def code_root() -> Path:
    return PROJECT_ROOT


def marker_in(folder: Path) -> Path | None:
    """The marker file in `folder` (`.dca.toml` preferred), or None."""
    return next((folder / n for n in MARKER_FILENAMES if (folder / n).is_file()), None)


def read_marker(folder: Path) -> dict[str, Any]:
    """Parse the marker in `folder`. Missing file → {}; invalid TOML → ValueError."""
    marker = marker_in(folder)
    if marker is None:
        return {}
    try:
        with marker.open("rb") as f:
            return tomllib.load(f)
    except tomllib.TOMLDecodeError as e:
        raise ValueError(f"{marker}: invalid TOML ({e})") from e


def find_marker_dir(start: Path) -> Path | None:
    """Nearest folder at or above `start` that contains a marker."""
    start = start.resolve()
    for folder in (start, *start.parents):
        if marker_in(folder) is not None:
            return folder
    return None


def resolve_data_root(start: Path | None = None, home: Path | None = None) -> DataRoot:
    """Find the data root (see module docstring for the order)."""
    if home is not None:
        return DataRoot(home.expanduser().resolve(), "--home")
    env = os.environ.get(HOME_ENV_VAR)
    if env:
        return DataRoot(Path(env).expanduser().resolve(), f"${HOME_ENV_VAR}")

    marker_dir = find_marker_dir(start if start is not None else Path.cwd())
    if marker_dir is not None:
        marker = marker_in(marker_dir)
        # A marker with data_dir says "the data is over there" (the code folder's marker).
        redirect = _path_key(marker_dir, "data_dir")
        if redirect is not None:
            return DataRoot(redirect, f"data_dir in {marker}")
        return DataRoot(marker_dir, f"{marker.name if marker else MARKER_FILENAME} in {marker_dir}")

    data_dir = _path_key(code_root(), "data_dir")
    if data_dir is not None:
        return DataRoot(data_dir, f"data_dir in {marker_in(code_root())}")

    if marker_in(code_root()) is not None:
        return DataRoot(code_root(), f"code folder ({MARKER_FILENAME} without data_dir)")
    return DataRoot(code_root(), f"code folder (no {MARKER_FILENAME} found)")


def _path_key(folder: Path, key: str) -> Path | None:
    """`key` from the marker in `folder` as an absolute path, or None if unset."""
    value = read_marker(folder).get(key)
    if not value:
        return None
    if not isinstance(value, str):
        raise ValueError(f"{marker_in(folder)}: {key} must be a string path")
    p = Path(value).expanduser()
    if not p.is_absolute():
        p = folder / p
    return p.resolve()
