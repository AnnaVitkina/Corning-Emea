"""End-to-end pipeline: input/ -> processing/ -> output/ result matrix.

Local:
    python run_pipeline.py

Google Colab (code in /content/Corning-Emea, data on shared Drive):
    import os
    exec(open('/content/Corning-Emea/run_pipeline.py').read())

Optional override for data root (local or Colab):
    os.environ['CORNING_EMEA_OCEAN_DATA_ROOT'] = '/path/to/RMT_EMEA_Ocean'
"""

from __future__ import annotations

import sys
from pathlib import Path


def _ensure_code_on_path() -> Path:
    """Make project imports work for both `python run_pipeline.py` and Colab exec()."""
    candidates: list[Path] = []
    if "__file__" in globals():
        candidates.append(Path(__file__).resolve().parent)
    candidates.extend(
        [
            Path("/content/Corning-Emea"),
            Path.cwd(),
        ]
    )
    for candidate in candidates:
        if (candidate / "paths.py").exists():
            resolved = candidate.resolve()
            if str(resolved) not in sys.path:
                sys.path.insert(0, str(resolved))
            return resolved
    cwd = Path.cwd().resolve()
    if str(cwd) not in sys.path:
        sys.path.insert(0, str(cwd))
    return cwd


_CODE_DIR = _ensure_code_on_path()

import paths as path_config

path_config.bootstrap()

import build_result_matrix as matrix
import process_input as ingest


def run_pipeline() -> None:
    print("=== Step 1/2: Load input file into processing/ ===")
    files = ingest.list_input_files()
    if not files:
        print(f"No supported files found in {ingest.INPUT_DIR}")
        print(f"Supported types: {', '.join(sorted(ingest.SUPPORTED_EXTENSIONS))}")
        return

    selected = ingest.ask_user_to_select(files)
    print(f"\nLoading: {selected.name}")
    df = ingest.load_to_dataframe(selected)
    print(f"Loaded DataFrame with shape {df.shape[0]} rows x {df.shape[1]} columns")

    processing_path = ingest.save_dataframe(df, selected)
    print(f"Saved processing copy to: {processing_path}")

    print("\n=== Step 2/2: Build result matrix into output/ ===")
    result = matrix.build_result_matrix(df, source=selected)
    domestic = matrix.build_domestic_matrix(result)
    result_path = matrix.save_result_matrix(result, processing_path, domestic=domestic)
    print(
        f"Saved result matrix to: {result_path} "
        f"({result.shape[0]} rows x {result.shape[1]} columns)"
    )
    if domestic is not None:
        print(
            f"Added sheet '{matrix.DOMESTIC_SHEET_NAME}' "
            f"({domestic.shape[0]} rows x {domestic.shape[1]} columns)"
        )
    print("\nPipeline complete.")


# `python run_pipeline.py` and Colab `exec(open(...).read())` both use __name__ == "__main__"
if __name__ == "__main__":
    run_pipeline()
