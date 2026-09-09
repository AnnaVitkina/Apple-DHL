"""Shared Excel parsing helpers for DHL US rates."""

from __future__ import annotations

import re
import warnings
from datetime import date, datetime
from pathlib import Path

import pandas as pd

WEIGHT_RANGE_PATTERN = re.compile(r"^(\d+(?:\.\d+)?)\s*(?:to|-)\s*(\d+(?:\.\d+)?)$", re.IGNORECASE)
COUNTRY_CODE_PATTERN = re.compile(r"\(([A-Z]{2})\)\s*$")


def cell_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def rate_value(value: object) -> float | int | None:
    if pd.isna(value):
        return None
    text = cell_text(value)
    if not text or text.lower() in {"on request", "n/a", "#n/a", "-"}:
        return None
    try:
        number = float(str(value).replace(",", ""))
    except (TypeError, ValueError):
        return None
    if number.is_integer():
        return int(number)
    return number


def _bracket_upper_label(number: float) -> str | int:
    if number == int(number):
        return int(number)
    return f"{number:g}"


def weight_bracket_label(chargeable_weight: object) -> str:
    text = cell_text(chargeable_weight)
    if not text:
        return ""

    range_match = WEIGHT_RANGE_PATTERN.match(text)
    if range_match:
        upper = _bracket_upper_label(float(range_match.group(2)))
        return f"<={upper}"

    lowered = text.lower()
    if "lbs or less" in lowered or "lb or less" in lowered:
        number_match = re.search(r"(\d+(?:\.\d+)?)", text)
        if number_match:
            upper = _bracket_upper_label(float(number_match.group(1)))
            return f"<={upper}"

    try:
        upper = _bracket_upper_label(float(text.replace(",", "")))
    except ValueError:
        return ""
    return f"<={upper}"


def bracket_sort_key(label: str) -> tuple[int, float | str]:
    match = re.match(r"^<=(.+)$", label)
    if match:
        upper_text = match.group(1)
        try:
            return (0, float(upper_text))
        except ValueError:
            return (1, upper_text)
    return (2, label)


def cost_column_name(cost_group: str, bracket_label: str) -> str:
    return f"{cost_group} ({bracket_label})"


def format_display_date(value: object) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, datetime):
        return value.date().strftime("%d.%m.%Y")
    if isinstance(value, date):
        return value.strftime("%d.%m.%Y")
    text = cell_text(value)
    if not text:
        return ""
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(text, fmt).strftime("%d.%m.%Y")
        except ValueError:
            continue
    return text


def read_sheet_raw(file_path: Path, sheet_name: str) -> pd.DataFrame:
    with warnings.catch_warnings():
        warnings.filterwarnings(
            "ignore",
            message="Unknown extension is not supported and will be removed",
            category=UserWarning,
        )
        return pd.read_excel(file_path, sheet_name=sheet_name, header=None)


def locate_rate_header(sheet_df: pd.DataFrame) -> tuple[int, dict[str, int]] | None:
    for row_index, row in sheet_df.iterrows():
        column_map = {
            cell_text(value): col_index
            for col_index, value in enumerate(row.tolist())
            if cell_text(value)
        }
        if "Origin Point" in column_map:
            return row_index, column_map
    return None


def weight_column_index(column_map: dict[str, int]) -> int | None:
    for name in ("Weight (Kg)", "Weight (KG)", "Weight (Lb)", "Weight (LB)"):
        if name in column_map:
            return column_map[name]
    return None


def rate_columns_from_header_row(
    header_row: pd.Series,
    *,
    weight_col: int | None,
) -> list[tuple[int, str]]:
    columns: list[tuple[int, str]] = []
    for col_index, value in enumerate(header_row.tolist()):
        if weight_col is not None and col_index <= weight_col:
            continue
        text = cell_text(value)
        if not text:
            continue
        if text.lower().startswith("rate"):
            columns.append((col_index, text))
        elif text.replace(".", "", 1).isdigit():
            columns.append((col_index, text))
        elif text.upper().startswith("ZONE "):
            columns.append((col_index, text))
    return columns


def parse_country_zone_table(
    sheet_df: pd.DataFrame,
    *,
    table_title: str,
) -> dict[int, list[str]]:
    """Parse multi-column country/zone tables used on US support tabs."""
    title_row = None
    for row_index, row in sheet_df.iterrows():
        row_text = " ".join(cell_text(value) for value in row.tolist())
        if table_title.lower() in row_text.lower():
            title_row = row_index
            break
    if title_row is None:
        return {}

    header_row = title_row + 1
    zone_by_number: dict[int, list[str]] = {}
    for row_index in range(header_row + 1, len(sheet_df)):
        row = sheet_df.iloc[row_index]
        saw_country = False
        for block_start in range(0, len(row), 3):
            country = cell_text(row.iloc[block_start]) if block_start < len(row) else ""
            zone_text = cell_text(row.iloc[block_start + 1]) if block_start + 1 < len(row) else ""
            if country.lower().startswith("countries"):
                continue
            if not country and not zone_text:
                continue
            if country:
                saw_country = True
            if not country or not zone_text:
                continue
            try:
                zone_number = int(float(zone_text))
            except ValueError:
                continue
            zone_by_number.setdefault(zone_number, []).append(country)
        if not saw_country and row_index > header_row + 1:
            break
    return zone_by_number


def parse_country_code(country_label: str) -> str:
    match = COUNTRY_CODE_PATTERN.search(country_label.strip())
    if match:
        return match.group(1)
    return ""


def load_3c_country_zones(sheet_df: pd.DataFrame) -> dict[int, list[str]]:
    return parse_country_zone_table(
        sheet_df,
        table_title="DHL EXPRESS INTERNATIONAL THIRD COUNTRY ZONING",
    )


def load_dhl_ib_destination_countries(sheet_df: pd.DataFrame) -> list[str]:
    title_row = None
    for row_index, row in sheet_df.iterrows():
        row_text = " ".join(cell_text(value) for value in row.tolist())
        if "import) destination local zone" in row_text.lower():
            title_row = row_index
            break
    if title_row is None:
        return []

    countries: list[str] = []
    for row_index in range(title_row + 2, len(sheet_df)):
        row = sheet_df.iloc[row_index]
        row_text = " ".join(cell_text(value) for value in row.tolist())
        if "dhl express worldwide destination zone" in row_text.lower():
            break
        for block_start in range(0, len(row), 3):
            country = cell_text(row.iloc[block_start]) if block_start < len(row) else ""
            if country and "(" in country:
                countries.append(country)
    return countries


def load_index_validity(file_path: Path, sheet_name: str = "Index") -> tuple[str, str]:
    sheet_df = read_sheet_raw(file_path, sheet_name)
    valid_from = ""
    valid_to = ""
    for _, row in sheet_df.iterrows():
        label = cell_text(row.iloc[1]).lower()
        if label == "valid from":
            valid_from = format_display_date(row.iloc[2])
        elif label == "valid to":
            valid_to = format_display_date(row.iloc[2])
    return valid_from, valid_to
