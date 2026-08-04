"""
Convert selected tabs from an input Excel workbook into cleaned dataframes
and save them to the processing/ folder.

Interactive flow:
  1. Choose input file from input/
  2. Review proposed default tabs (Tab Index + Accessorials + ACTIVE tabs)
  3. Accept, change, or add tabs
  4. Write one multi-sheet workbook to processing/
"""

from __future__ import annotations

import re
import sys
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

import pandas as pd

from project_paths import INPUT_DIR, PROCESSING_DIR, ensure_workspace_dirs

EXCEL_SUFFIXES = {".xlsx", ".xls", ".xlsm"}

DEFAULT_SHEETS_ALWAYS = ("Tab Index", "Accessorials")
ACTIVE_TAB_STATUS = "ACTIVE"
TAB_INDEX_SHEET_NAME = "Tab Index"

HEADER_MARKERS = (
    "origin",
    "shipment type",
    "chargeable weight",
    "rate logic",
    "scac",
    "line charge code",
    "ansi code",
)

TAB_INDEX_HEADER_MARKERS = ("tab-name",)
PUK_FILE_PATTERN = re.compile(r"DHLPUK", re.IGNORECASE)
PUK_DESTINATION_COLUMN_PATTERN = re.compile(r"^(GB|GB_NI|GB_HI|IE|[A-Z]{2})$")

METADATA_LABELS = {
    "version number": "Version Number",
    "version date": "Version Date",
    "destination country": "Destination Country",
    "destination iso": "Destination ISO",
    "billing currency": "Billing Currency",
}


def _normalize_label(value: object) -> str:
    return _cell_text(value).lower().rstrip(":").strip()


@dataclass(frozen=True)
class SheetSelection:
    file_path: Path
    sheet_name: str

    @property
    def label(self) -> str:
        return f"{self.file_path.name} -> {self.sheet_name}"


@dataclass
class SheetMetadata:
    version_number: object = None
    version_date: object = None
    destination_country: object = None
    destination_iso: object = None
    billing_currency: object = None


def list_input_files() -> list[Path]:
    return [
        path
        for path in sorted(INPUT_DIR.iterdir())
        if path.is_file()
        and path.suffix.lower() in EXCEL_SUFFIXES
        and not path.name.startswith("~$")
    ]


def parse_selection(raw: str, max_index: int) -> list[int]:
    """Parse selections like '2,4,7-19,24-29,34,36' into zero-based indices."""
    raw = raw.strip().lower()
    if raw in {"all", "*"}:
        return list(range(max_index))

    indices: set[int] = set()
    for part in re.split(r"\s*,\s*", raw):
        if not part:
            continue
        part = part.strip()
        if "-" in part:
            start_s, end_s = part.split("-", 1)
            start = int(start_s.strip()) - 1
            end = int(end_s.strip()) - 1
            if start > end or start < 0 or end >= max_index:
                raise ValueError(f"Invalid range: {part}")
            indices.update(range(start, end + 1))
        else:
            idx = int(part) - 1
            if idx < 0 or idx >= max_index:
                raise ValueError(f"Invalid index: {part}")
            indices.add(idx)
    return sorted(indices)


def prompt_selection(title: str, items: list[str], allow_empty: bool = False) -> list[int]:
    if not items:
        return []

    print(f"\n{title}")
    for i, item in enumerate(items, start=1):
        print(f"  {i}. {item}")

    hint = "Enter numbers (e.g. 2,4,7-19,24-29,34,36) or 'all'"

    while True:
        raw = input(f"{hint}: ").strip()
        if not raw and allow_empty:
            return []
        if not raw:
            print("Please enter at least one choice.")
            continue
        try:
            chosen = parse_selection(raw, len(items))
            if chosen or allow_empty:
                return chosen
            print("Please enter at least one choice.")
        except ValueError as exc:
            print(f"Invalid input: {exc}")


def read_tab_index_sheet(file_path: Path) -> pd.DataFrame | None:
    """Read Tab Index from a workbook, including cleaned processing extracts."""
    try:
        workbook = pd.ExcelFile(file_path)
    except Exception:
        return None

    sheet_name = next(
        (name for name in workbook.sheet_names if name.strip().lower() == TAB_INDEX_SHEET_NAME.lower()),
        None,
    )
    if sheet_name is None:
        return None

    preview = pd.read_excel(file_path, sheet_name=sheet_name, header=None, nrows=10)
    header_row = next(
        (
            row_idx
            for row_idx, row in preview.iterrows()
            if any(_normalize_label(value) == "tab-name" for value in row)
        ),
        None,
    )

    if header_row is None:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
    else:
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)

    df.columns = [str(column).strip() for column in df.columns]
    return df


def active_tab_names_from_index(tab_index_df: pd.DataFrame | None) -> list[str]:
    if tab_index_df is None or tab_index_df.empty:
        return []

    status_column = next(
        (column for column in tab_index_df.columns if column.strip().lower() == "status"),
        None,
    )
    name_column = next(
        (column for column in tab_index_df.columns if column.strip().lower() == "tab-name"),
        None,
    )
    if status_column is None or name_column is None:
        return []

    active_rows = tab_index_df[
        tab_index_df[status_column].astype(str).str.strip().str.upper() == ACTIVE_TAB_STATUS
    ]
    return [_cell_text(name) for name in active_rows[name_column] if _cell_text(name)]


def propose_default_tabs(sheet_names: list[str], file_path: Path) -> list[str]:
    """Select Tab Index, Accessorials, and tabs marked ACTIVE in Tab Index."""
    by_lower = {name.lower(): name for name in sheet_names}
    selected_lower: set[str] = set()

    for exact_name in DEFAULT_SHEETS_ALWAYS:
        actual = by_lower.get(exact_name.lower())
        if actual is not None:
            selected_lower.add(actual.lower())

    tab_index_df = read_tab_index_sheet(file_path)
    for tab_name in active_tab_names_from_index(tab_index_df):
        actual = by_lower.get(tab_name.lower())
        if actual is not None:
            selected_lower.add(actual.lower())

    return [name for name in sheet_names if name.lower() in selected_lower]


def is_default_tab(sheet_name: str, default_tabs: set[str] | None = None) -> bool:
    if default_tabs is not None:
        return sheet_name in default_tabs
    return sheet_name.strip().lower() in {name.lower() for name in DEFAULT_SHEETS_ALWAYS}


def select_input_file(files: list[Path], *, auto: bool = False) -> Path:
    if not files:
        print(f"No Excel files found in: {INPUT_DIR}")
        sys.exit(1)

    if auto:
        if len(files) == 1:
            print(f"\nAuto mode: using input file {files[0].name}")
        else:
            print(f"\nAuto mode: using first input file {files[0].name}")
        return files[0]

    print("\nAvailable input files:")
    for i, path in enumerate(files, start=1):
        print(f"  {i}. {path.name}")

    indices = prompt_selection("Select file to convert:", [f.name for f in files])
    if len(indices) != 1:
        print("Please select exactly one file.")
        return select_input_file(files)
    return files[indices[0]]


def print_all_tabs(sheet_names: list[str], selected: list[str], default_tabs: set[str]) -> None:
    print("\nAll available tabs:")
    for i, name in enumerate(sheet_names, start=1):
        marker = " [selected]" if name in selected else ""
        default_marker = " [default]" if is_default_tab(name, default_tabs) else ""
        print(f"  {i}. {name}{marker}{default_marker}")


def print_current_selection(selected: list[str]) -> None:
    if selected:
        print("\nCurrent tab selection:")
        for name in selected:
            print(f"  - {name}")
    else:
        print("\nNo tabs selected yet.")


def select_tabs_interactive(file_path: Path, sheet_names: list[str]) -> list[str]:
    selected = propose_default_tabs(sheet_names, file_path)
    default_tabs = set(selected)

    while True:
        print_all_tabs(sheet_names, selected, default_tabs)
        print_current_selection(selected)
        print("\nTab selection options:")
        print("  1. Accept current selection and convert")
        print("  2. Change tabs (replace selection)")
        print("  3. Add tabs to current selection")
        print("  4. Enter custom selection (replace; e.g. 2,4,7-19,24-29,34,36)")

        choice = input("Choose option (1-4): ").strip()

        if choice == "1":
            if not selected:
                print("Select at least one tab before converting.")
                continue
            return selected

        if choice in {"2", "4"}:
            indices = prompt_selection(
                f"Choose sheet(s) for {file_path.name} (use numbers from the tab list above):",
                sheet_names,
            )
            selected = [sheet_names[i] for i in indices]
            continue

        if choice == "3":
            indices = prompt_selection(
                "Choose additional sheet(s) to add (use numbers from the tab list above):",
                sheet_names,
            )
            for idx in indices:
                tab_name = sheet_names[idx]
                if tab_name not in selected:
                    selected.append(tab_name)
            continue

        print("Invalid choice. Enter 1, 2, 3, or 4.")


def _cell_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def _format_metadata_date(value: object) -> object:
    if pd.isna(value):
        return pd.NA
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    return value


def _row_texts(row: pd.Series) -> list[str]:
    return [_cell_text(value).lower() for value in row]


def extract_sheet_metadata(df_raw: pd.DataFrame, max_scan_rows: int = 20) -> SheetMetadata:
    metadata = SheetMetadata()
    scan_limit = min(len(df_raw), max_scan_rows)

    for row_idx in range(scan_limit):
        row = df_raw.iloc[row_idx]
        for col_idx in range(len(row) - 1):
            label = _normalize_label(row.iloc[col_idx])
            if label not in METADATA_LABELS:
                continue
            candidate = row.iloc[col_idx + 1]
            if pd.isna(candidate):
                continue

            field_name = METADATA_LABELS[label]
            if field_name == "Version Number" and metadata.version_number is None:
                metadata.version_number = candidate
            elif field_name == "Version Date" and metadata.version_date is None:
                metadata.version_date = _format_metadata_date(candidate)
            elif field_name == "Destination Country" and metadata.destination_country is None:
                metadata.destination_country = candidate
            elif field_name == "Destination ISO" and metadata.destination_iso is None:
                metadata.destination_iso = candidate
            elif field_name == "Billing Currency" and metadata.billing_currency is None:
                metadata.billing_currency = candidate

    return metadata


def _normalize_headers(headers: list[object]) -> list[str]:
    cleaned: list[str] = []
    seen: dict[str, int] = {}

    for idx, header in enumerate(headers, start=1):
        value = _cell_text(header)
        if not value or value.lower() == "nan":
            value = f"column_{idx}"

        base = value
        count = seen.get(base, 0)
        if count:
            value = f"{base}_{count + 1}"
        seen[base] = count + 1
        cleaned.append(value)

    return cleaned


def _drop_empty_rows(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    non_empty_mask = df.apply(
        lambda row: any(_cell_text(value) for value in row),
        axis=1,
    )
    return df.loc[non_empty_mask].copy()


def _is_puk_destination_column(column: object) -> bool:
    return bool(PUK_DESTINATION_COLUMN_PATTERN.match(_cell_text(column).upper()))


def _drop_empty_columns(
    df: pd.DataFrame,
    *,
    preserve_columns: set[str] | None = None,
) -> pd.DataFrame:
    if df.empty:
        return df

    preserve_columns = preserve_columns or set()
    keep_columns = [
        column
        for column in df.columns
        if column in preserve_columns or any(_cell_text(value) for value in df[column])
    ]
    return df.loc[:, keep_columns].copy()


def _row_contains_marker(row: pd.Series, markers: tuple[str, ...]) -> bool:
    texts = set(_row_texts(row))
    return any(marker in texts for marker in markers)


def find_header_row_index(df_raw: pd.DataFrame, max_scan_rows: int = 30) -> int | None:
    scan_limit = min(len(df_raw), max_scan_rows)
    for row_idx in range(scan_limit):
        if _row_contains_marker(df_raw.iloc[row_idx], HEADER_MARKERS):
            return row_idx
    return None


def find_tab_index_header_row_index(df_raw: pd.DataFrame, max_scan_rows: int = 20) -> int | None:
    scan_limit = min(len(df_raw), max_scan_rows)
    for row_idx in range(scan_limit):
        if _row_contains_marker(df_raw.iloc[row_idx], TAB_INDEX_HEADER_MARKERS):
            return row_idx
    return None


def clean_tab_index_df(df_raw: pd.DataFrame) -> pd.DataFrame:
    """Parse Tab Index using the Tab-Name header row."""
    if df_raw.empty:
        return df_raw.copy()

    header_row_idx = find_tab_index_header_row_index(df_raw)
    if header_row_idx is None:
        cleaned = _drop_empty_rows(df_raw)
        cleaned.columns = _normalize_headers(list(cleaned.columns))
        return cleaned.reset_index(drop=True)

    headers = _normalize_headers(df_raw.iloc[header_row_idx].tolist())
    df = df_raw.iloc[header_row_idx + 1 :].copy()
    df.columns = headers
    df = _drop_empty_rows(df)
    df = _drop_empty_columns(df)
    return df.reset_index(drop=True)


def _attach_metadata_columns(df: pd.DataFrame, metadata: SheetMetadata) -> pd.DataFrame:
    metadata_columns: list[tuple[str, object]] = []

    if metadata.version_number is not None:
        metadata_columns.append(("Version Number", metadata.version_number))
    if metadata.version_date is not None:
        metadata_columns.append(("Version Date", metadata.version_date))
    if metadata.destination_country is not None:
        metadata_columns.append(("Destination Country", metadata.destination_country))
    if metadata.destination_iso is not None:
        metadata_columns.append(("Destination ISO", metadata.destination_iso))
    if metadata.billing_currency is not None:
        metadata_columns.append(("Billing Currency", metadata.billing_currency))

    if not metadata_columns:
        return df

    enriched = df.copy()
    for offset, (column_name, value) in enumerate(metadata_columns):
        enriched.insert(offset, column_name, value)
    return enriched


def clean_tab_df(
    df_raw: pd.DataFrame,
    sheet_name: str | None = None,
    *,
    source_file: Path | None = None,
) -> pd.DataFrame:
    """Remove preamble rows above the header and fully empty rows/columns."""
    if sheet_name and sheet_name.strip().lower() == "tab index":
        return clean_tab_index_df(df_raw)

    metadata = extract_sheet_metadata(df_raw)

    if df_raw.empty:
        return df_raw.copy()

    header_row_idx = find_header_row_index(df_raw)
    if header_row_idx is None:
        cleaned = _drop_empty_rows(df_raw)
        cleaned.columns = _normalize_headers(list(cleaned.columns))
        cleaned = cleaned.reset_index(drop=True)
        return _attach_metadata_columns(cleaned, metadata)

    headers = _normalize_headers(df_raw.iloc[header_row_idx].tolist())
    data_start = header_row_idx + 1

    df = df_raw.iloc[data_start:].copy()
    df.columns = headers
    df = _drop_empty_rows(df)
    preserve_columns: set[str] = set()
    if source_file is not None and PUK_FILE_PATTERN.search(source_file.name):
        preserve_columns = {
            str(column)
            for column in df.columns
            if _is_puk_destination_column(column)
        }
    df = _drop_empty_columns(df, preserve_columns=preserve_columns)
    df = df.reset_index(drop=True)
    return _attach_metadata_columns(df, metadata)


def tab_to_df(file_path: Path, sheet_name: str) -> pd.DataFrame:
    raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    return clean_tab_df(raw, sheet_name=sheet_name, source_file=file_path)


MAX_SHEET_NAME_LEN = 31


def sanitize_sheet_part(text: str) -> str:
    return re.sub(r"[\[\]:*?/\\]", "_", text).strip() or "Sheet"


def output_sheet_name(sheet_name: str, used: set[str]) -> str:
    base = sanitize_sheet_part(sheet_name)
    if base not in used:
        used.add(base)
        return base

    for n in range(2, 1000):
        suffix = f"_{n}"
        candidate = sanitize_sheet_part(sheet_name)[: MAX_SHEET_NAME_LEN - len(suffix)] + suffix
        if candidate not in used:
            used.add(candidate)
            return candidate

    raise RuntimeError(f"Could not create a unique sheet name for: {sheet_name}")


def collect_frames(
    file_path: Path,
    sheet_names: list[str],
) -> list[tuple[str, pd.DataFrame]]:
    frames: list[tuple[str, pd.DataFrame]] = []
    used_names: set[str] = set()

    for sheet_name in sheet_names:
        try:
            raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
            metadata = extract_sheet_metadata(raw)
            df = clean_tab_df(raw, sheet_name=sheet_name, source_file=file_path)
            removed_rows = len(raw) - len(df)
        except Exception as exc:
            print(f"Skipping {sheet_name}: could not read sheet ({exc})")
            continue

        label = output_sheet_name(sheet_name, used_names)
        frames.append((label, df))
        meta_parts: list[str] = []
        if metadata.version_number is not None:
            meta_parts.append(f"version: {metadata.version_number}")
        if metadata.destination_iso is not None:
            meta_parts.append(f"destination: {metadata.destination_iso}")
        if metadata.billing_currency is not None:
            meta_parts.append(f"currency: {metadata.billing_currency}")
        meta_summary = f"; {'; '.join(meta_parts)}" if meta_parts else ""
        print(
            f"  Loaded: {sheet_name} -> '{label}' "
            f"({len(df)} rows, {len(df.columns)} columns; removed {removed_rows} preamble/empty rows"
            f"{meta_summary})"
        )

    return frames


def save_to_processing(file_path: Path, frames: list[tuple[str, pd.DataFrame]]) -> Path:
    output_path = PROCESSING_DIR / f"{file_path.stem}_extracted.xlsx"

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, df in frames:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"  Wrote tab: {sheet_name}")

    return output_path


def run_convert(*, auto: bool = False, file_path: Path | None = None) -> Path:
    ensure_workspace_dirs()
    files = list_input_files()
    selected_file = file_path or select_input_file(files, auto=auto)

    if not selected_file.exists():
        raise FileNotFoundError(f"Input file not found: {selected_file}")

    try:
        workbook = pd.ExcelFile(selected_file)
    except Exception as exc:
        raise RuntimeError(f"Could not open {selected_file.name}: {exc}") from exc

    if auto:
        selected_tabs = propose_default_tabs(workbook.sheet_names, selected_file)
        if not selected_tabs:
            raise RuntimeError(
                f"No default tabs matched in {selected_file.name}. "
                "Run without --auto to choose tabs manually."
            )
        print(f"Auto mode: converting tabs: {', '.join(selected_tabs)}")
    else:
        selected_tabs = select_tabs_interactive(selected_file, workbook.sheet_names)

    print(f"\nConverting {len(selected_tabs)} tab(s) from {selected_file.name}...")
    frames = collect_frames(selected_file, selected_tabs)

    if not frames:
        raise RuntimeError("Nothing to save. No sheets could be loaded.")

    output_path = save_to_processing(selected_file, frames)
    print(f"\nSaved {len(frames)} sheet(s) to: {output_path}")
    return output_path


def main() -> int:
    try:
        run_convert()
        return 0
    except KeyboardInterrupt:
        print("\nOperation cancelled.")
        return 130
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
