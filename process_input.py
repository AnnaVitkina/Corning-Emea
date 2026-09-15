"""Select a file from input/, load it into a DataFrame, save as xlsx in processing/."""

from pathlib import Path

import pandas as pd

from paths import INPUT_DIR, PROCESSING_DIR

SUPPORTED_EXTENSIONS = {".xlsx", ".xls", ".csv", ".tsv"}


def list_input_files() -> list[Path]:
    if not INPUT_DIR.exists():
        raise FileNotFoundError(f"Input folder not found: {INPUT_DIR}")

    return sorted(
        p
        for p in INPUT_DIR.iterdir()
        if p.is_file()
        and not p.name.startswith("~$")
        and p.suffix.lower() in SUPPORTED_EXTENSIONS
    )


def ask_user_to_select(files: list[Path]) -> Path:
    print("\nFiles available in input/:\n")
    for i, path in enumerate(files, start=1):
        print(f"  {i}. {path.name}")

    while True:
        choice = input(f"\nSelect a file to process [1-{len(files)}]: ").strip()
        if not choice.isdigit():
            print("Please enter a number.")
            continue
        index = int(choice)
        if 1 <= index <= len(files):
            return files[index - 1]
        print(f"Please choose a number between 1 and {len(files)}.")


def load_to_dataframe(path: Path) -> pd.DataFrame:
    suffix = path.suffix.lower()
    if suffix == ".xlsx":
        # calamine is more tolerant of invalid Excel stylesheets
        try:
            return pd.read_excel(path, engine="calamine")
        except Exception:
            return pd.read_excel(path, engine="openpyxl")
    if suffix == ".xls":
        return pd.read_excel(path)
    if suffix == ".csv":
        return pd.read_csv(path)
    if suffix == ".tsv":
        return pd.read_csv(path, sep="\t")
    raise ValueError(f"Unsupported file type: {path.suffix}")


def save_dataframe(df: pd.DataFrame, source: Path) -> Path:
    PROCESSING_DIR.mkdir(parents=True, exist_ok=True)
    output_path = PROCESSING_DIR / f"{source.stem}.xlsx"
    df.to_excel(output_path, index=False)
    return output_path


def main() -> None:
    files = list_input_files()
    if not files:
        print(f"No supported files found in {INPUT_DIR}")
        print(f"Supported types: {', '.join(sorted(SUPPORTED_EXTENSIONS))}")
        return

    selected = ask_user_to_select(files)
    print(f"\nLoading: {selected.name}")
    df = load_to_dataframe(selected)
    print(f"Loaded DataFrame with shape {df.shape[0]} rows x {df.shape[1]} columns")

    output_path = save_dataframe(df, selected)
    print(f"Saved to: {output_path}")


if __name__ == "__main__":
    main()
