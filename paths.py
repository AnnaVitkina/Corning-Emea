"""Resolve code + data folders for local runs and Google Colab.

Data root (input / processing / output):
  1) env CORNING_EMEA_OCEAN_DATA_ROOT, if set
  2) Google Drive shared folder, if it exists (Colab)
  3) same folder as the Python code (local default)

Code root (where *.py live):
  - local: directory of this file
  - Colab exec(): /content/Corning-Emea when present
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

ENV_DATA_ROOT = "CORNING_EMEA_OCEAN_DATA_ROOT"

COLAB_CODE_DIR = Path("/content/Corning-Emea")
COLAB_DATA_ROOT = Path(
    "/content/drive/Shareddrives/FA Ops Europe: Rate Maintenance Team "
    "/Documents/AI Adoption RMT/RMT_Corning/RMT_EMEA_Ocean"
)

CODE_DIR: Path
DATA_ROOT: Path
INPUT_DIR: Path
PROCESSING_DIR: Path
OUTPUT_DIR: Path


def resolve_code_dir() -> Path:
    """Directory that contains run_pipeline.py / process_input.py."""
    if "__file__" in globals():
        here = Path(__file__).resolve().parent
        if (here / "run_pipeline.py").exists() or (here / "paths.py").exists():
            return here

    for candidate in (
        COLAB_CODE_DIR,
        Path.cwd(),
        Path("/content/Corning-Emea"),
    ):
        if (candidate / "run_pipeline.py").exists() or (candidate / "paths.py").exists():
            return candidate.resolve()

    if "__file__" in globals():
        return Path(__file__).resolve().parent
    return Path.cwd().resolve()


def resolve_data_root(code_dir: Path | None = None) -> Path:
    """Folder that contains input/, processing/, output/."""
    override = os.environ.get(ENV_DATA_ROOT, "").strip()
    if override:
        return Path(override).expanduser().resolve()

    if COLAB_DATA_ROOT.exists():
        return COLAB_DATA_ROOT.resolve()

    base = code_dir or resolve_code_dir()
    return base.resolve()


def ensure_data_folders(data_root: Path) -> None:
    for name in ("input", "processing", "output"):
        (data_root / name).mkdir(parents=True, exist_ok=True)


def configure(code_dir: Path | None = None) -> Path:
    """Set module-level path globals. Safe to call multiple times (e.g. after Drive mount)."""
    global CODE_DIR, DATA_ROOT, INPUT_DIR, PROCESSING_DIR, OUTPUT_DIR

    CODE_DIR = (code_dir or resolve_code_dir()).resolve()
    if str(CODE_DIR) not in sys.path:
        sys.path.insert(0, str(CODE_DIR))

    DATA_ROOT = resolve_data_root(CODE_DIR)
    INPUT_DIR = DATA_ROOT / "input"
    PROCESSING_DIR = DATA_ROOT / "processing"
    OUTPUT_DIR = DATA_ROOT / "output"
    ensure_data_folders(DATA_ROOT)
    return DATA_ROOT


def ensure_dependencies() -> None:
    """Install runtime packages when missing (needed on fresh Colab runtimes)."""
    import importlib
    import subprocess

    required = [
        ("pandas", "pandas"),
        ("openpyxl", "openpyxl"),
        ("python-calamine", "python_calamine"),
    ]
    missing = []
    for pip_name, module_name in required:
        try:
            importlib.import_module(module_name)
        except ImportError:
            missing.append(pip_name)

    if not missing:
        return

    code_dir = resolve_code_dir()
    req_file = code_dir / "requirements.txt"
    print(f"Installing missing packages: {', '.join(missing)}")
    if req_file.exists():
        cmd = [sys.executable, "-m", "pip", "install", "-q", "-r", str(req_file)]
    else:
        cmd = [sys.executable, "-m", "pip", "install", "-q", *missing]
    subprocess.check_call(cmd)
    # Clear failed import caches so pandas can see the new engine
    for mod in ("python_calamine", "pandas.io.excel", "pandas.io.excel._calamine"):
        sys.modules.pop(mod, None)


def maybe_mount_google_drive() -> bool:
    """Mount Drive on Colab when the shared data folder is not visible yet."""
    if COLAB_DATA_ROOT.exists():
        return True
    if not Path("/content").exists():
        return False
    try:
        from google.colab import drive  # type: ignore
    except ImportError:
        return False

    print("Mounting Google Drive...")
    drive.mount("/content/drive", force_remount=False)
    return COLAB_DATA_ROOT.exists()


def bootstrap() -> Path:
    """Colab/local entry setup: deps, Drive mount, then configure paths."""
    ensure_dependencies()
    maybe_mount_google_drive()
    data_root = configure()
    mode = (
        "colab-drive"
        if data_root == COLAB_DATA_ROOT.resolve() or str(COLAB_DATA_ROOT) in str(data_root)
        else "local"
    )
    if ENV_DATA_ROOT in os.environ and os.environ[ENV_DATA_ROOT].strip():
        mode = "env-override"
    print(f"Code dir : {CODE_DIR}")
    print(f"Data root: {DATA_ROOT}  [{mode}]")
    print(f"  input     : {INPUT_DIR}")
    print(f"  processing: {PROCESSING_DIR}")
    print(f"  output    : {OUTPUT_DIR}")
    return data_root


# Local imports / `python process_input.py` work without an explicit bootstrap call.
configure()
