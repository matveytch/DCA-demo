"""Data-root resolution (dca.paths): --home > $DCA_HOME > dca.toml lookup > code-root data_dir > code root."""

from pathlib import Path

import pytest

from dca import paths
from dca.paths import MARKER_FILENAME, DataRoot, code_root, resolve_data_root


@pytest.fixture(autouse=True)
def _no_env(monkeypatch):
    monkeypatch.delenv("DCA_HOME", raising=False)


@pytest.fixture
def fake_code(tmp_path, monkeypatch):
    """A code root with no dca.toml, so the fallback is predictable."""
    code = tmp_path / "code"
    code.mkdir()
    monkeypatch.setattr(paths, "PROJECT_ROOT", code)
    return code


def test_marker_found_by_walking_up_from_deal(tmp_path, fake_code):
    root = tmp_path / "data root"
    deal = root / "deals" / "1-2026-1705_1000kg"
    deal.mkdir(parents=True)
    (root / MARKER_FILENAME).write_text("")

    found = resolve_data_root(deal)
    assert found.path == root.resolve()
    assert MARKER_FILENAME in found.source


def test_nearest_marker_wins(tmp_path, fake_code):
    outer = tmp_path / "outer"
    inner = outer / "inner"
    deal = inner / "deals" / "X"
    deal.mkdir(parents=True)
    (outer / MARKER_FILENAME).write_text("")
    (inner / MARKER_FILENAME).write_text("")
    assert resolve_data_root(deal).path == inner.resolve()


def test_env_beats_marker(tmp_path, fake_code, monkeypatch):
    root = tmp_path / "root"
    (root / "deals").mkdir(parents=True)
    (root / MARKER_FILENAME).write_text("")
    monkeypatch.setenv("DCA_HOME", str(tmp_path / "elsewhere"))

    found = resolve_data_root(root / "deals")
    assert found.path == (tmp_path / "elsewhere").resolve()
    assert found.source == "$DCA_HOME"


def test_home_option_beats_env(tmp_path, fake_code, monkeypatch):
    monkeypatch.setenv("DCA_HOME", str(tmp_path / "env"))
    found = resolve_data_root(tmp_path, home=tmp_path / "opt")
    assert found.path == (tmp_path / "opt").resolve()
    assert found.source == "--home"


def test_fallback_is_code_root(tmp_path, fake_code):
    found = resolve_data_root(tmp_path / "no" / "marker" / "here")
    assert found.path == fake_code
    assert "code folder" in found.source


def test_code_root_data_dir_relative(tmp_path, fake_code):
    (fake_code / MARKER_FILENAME).write_text('data_dir = "../DCA-data"\n')
    found = resolve_data_root(tmp_path / "somewhere")
    assert found.path == (tmp_path / "DCA-data").resolve()
    assert "data_dir" in found.source


def test_code_root_data_dir_tilde(tmp_path, fake_code, monkeypatch):
    monkeypatch.setenv("HOME", str(tmp_path / "userhome"))
    (fake_code / MARKER_FILENAME).write_text('data_dir = "~/Seafile/DCA"\n')
    found = resolve_data_root(tmp_path / "somewhere")
    assert found.path == (tmp_path / "userhome" / "Seafile" / "DCA").resolve()


def test_marker_with_data_dir_redirects(tmp_path, fake_code):
    """Running from inside the code folder of a split layout still lands on the data."""
    (fake_code / MARKER_FILENAME).write_text('data_dir = "../DCA-data"\n')
    found = resolve_data_root(fake_code)
    assert found.path == (tmp_path / "DCA-data").resolve()


def test_split_layout_deal_uses_data_side_marker(tmp_path, fake_code):
    data = tmp_path / "DCA-data"
    deal = data / "deals" / "X"
    deal.mkdir(parents=True)
    (data / MARKER_FILENAME).write_text(f'code_dir = "{fake_code}"\n')
    (fake_code / MARKER_FILENAME).write_text(f'data_dir = "{data}"\n')
    assert resolve_data_root(deal).path == data.resolve()


def test_invalid_toml_is_a_clear_error(tmp_path, fake_code):
    (fake_code / MARKER_FILENAME).write_text("data_dir = [unclosed\n")
    with pytest.raises(ValueError, match="invalid TOML"):
        resolve_data_root(tmp_path / "somewhere")


def test_non_string_data_dir_is_rejected(tmp_path, fake_code):
    (fake_code / MARKER_FILENAME).write_text("data_dir = 5\n")
    with pytest.raises(ValueError, match="must be a string"):
        resolve_data_root(tmp_path / "somewhere")


def test_data_root_layout():
    root = DataRoot(Path("/x"), "test")
    assert root.new_deals_dir == Path("/x/Новые сделки")
    assert root.outgoing_dir("ME") == Path("/x/ME/Outgoing")
    assert root.incoming_dir("ME") == Path("/x/ME/Incoming")
    assert root.scans_dir == Path("/x/Data/scans")
    assert root.ledger_path == Path("/x/Data/batch_ledger.xlsx")
    assert root.master_db_path == Path("/x/Data/master_db.xlsx")


def test_real_project_has_marker():
    """The shipped code folder points its data at workspace/ (gitignored), not at itself."""
    assert (paths.code_root() / MARKER_FILENAME).is_file()
    assert paths.read_marker(paths.code_root())["data_dir"] == "workspace"
    assert "workspace/" in (paths.code_root() / ".gitignore").read_text(encoding="utf-8")


# ─── Launcher without `pip install` (2026-10-29) ───────────────

def test_tests_import_the_code_next_to_them():
    """pyproject `pythonpath = ["src"]`: a stale `pip install -e` path can't be picked up."""
    import dca
    assert Path(dca.__file__).resolve().parent == code_root().joinpath("src", "dca").resolve()


def test_launcher_runs_the_code_next_to_it_after_a_move(tmp_path):
    """`./dca` finds src/ beside itself: a moved copy runs its own code, no install, no dca on PATH."""
    import shutil
    import subprocess
    import sys

    moved = tmp_path.joinpath("Новая папка", "DCA copy", "dca")
    shutil.copytree(code_root().joinpath("src", "dca"), moved.joinpath("src", "dca"),
                    ignore=shutil.ignore_patterns("__pycache__"))
    shutil.copy2(code_root().joinpath("dca"), moved.joinpath("dca"))
    init = moved.joinpath("src", "dca", "__init__.py")
    init.write_text(init.read_text(encoding="utf-8") + '\nprint("MOVED-COPY-CODE")\n', encoding="utf-8")

    bin_dir = tmp_path.joinpath("bin")
    bin_dir.mkdir()
    bin_dir.joinpath("python3").symlink_to(sys.executable)
    env = {"PATH": f"{bin_dir}:/usr/bin:/bin", "HOME": str(tmp_path)}
    proc = subprocess.run([str(moved.joinpath("dca")), "hello"], env=env, cwd=tmp_path,
                          capture_output=True, text=True, timeout=60)
    assert proc.returncode == 0, proc.stdout + proc.stderr
    assert "MOVED-COPY-CODE" in proc.stdout and "hello" in proc.stdout


def test_launcher_is_executable_and_in_git():
    import os
    import subprocess
    launcher = code_root().joinpath("dca")
    assert os.access(launcher, os.X_OK)
    mode = subprocess.run(["git", "ls-files", "-s", "dca"], cwd=code_root(), capture_output=True, text=True).stdout
    assert mode == "" or mode.startswith("100755"), mode   # "" until the first commit
