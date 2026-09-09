"""
Build a US rate matrix / result workbook from DHL US rate cards.

This is a separate flow from the EU/ME/GB matrix builder in build_matrix.py.
"""

from __future__ import annotations

import re
import sys
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd

from build_matrix import CURRENCY_COLUMN, save_matrix, transport_cost_column_name
from project_paths import INPUT_DIR, OUTPUT_DIR, ensure_workspace_dirs
from us_config import (
    DHLC_DESTINATION_ZONES_SHEET,
    DHL_3C_ZONES_SHEET,
    DHL_ORIGIN_ZONES_SHEET,
    FRED_SERVICE_BY_RATE_CARD,
    INDEX_SHEET,
    DHL_IB_DESTINATION_REGION,
    US_COST_GROUP_ORDER,
    USCostGroup,
    US_RATE_TAB_CONFIGS,
    US_SHIPMENT_COLUMNS,
    us_cost_group_label,
)
from us_excel import (
    bracket_sort_key,
    cell_text,
    load_index_validity,
    locate_rate_header,
    parse_country_zone_table,
    rate_columns_from_header_row,
    rate_value,
    read_sheet_raw,
    weight_bracket_label,
    weight_column_index,
)

EXCEL_SUFFIXES = {".xlsx", ".xls", ".xlsm"}
US_FILE_PATTERN = re.compile(r"DHL US Rates", re.IGNORECASE)


def is_us_rate_file(file_path: Path | str) -> bool:
    return bool(US_FILE_PATTERN.search(Path(file_path).name))

@dataclass
class USLaneBuilder:
    lanes: dict[tuple[str, ...], dict[str, object]] = field(default_factory=dict)
    brackets: set[str] = field(default_factory=set)
    cost_group_brackets: dict[str, set[str]] = field(default_factory=dict)

    def add_rate(
        self,
        *,
        key: tuple[str, ...],
        lane_values: dict[str, object],
        cost_group: USCostGroup,
        bracket: str,
        rate: float | int,
    ) -> None:
        self.brackets.add(bracket)
        group_label = us_cost_group_label(cost_group)
        self.cost_group_brackets.setdefault(group_label, set()).add(bracket)
        cost_column = transport_cost_column_name(group_label, bracket)
        if key not in self.lanes:
            self.lanes[key] = {**lane_values, CURRENCY_COLUMN: "USD"}
        lane = self.lanes[key]
        for field_name, value in lane_values.items():
            if field_name == CURRENCY_COLUMN:
                continue
            if not lane.get(field_name) and value:
                lane[field_name] = value
        existing_rate = lane.get(cost_column)
        if existing_rate is not None and existing_rate != rate:
            raise ValueError(
                f"Conflicting rates for lane {key} at {bracket}: {existing_rate} vs {rate}"
            )
        lane[cost_column] = rate


def format_3c_zone_label(zone_number: int) -> str:
    return f"Zone {zone_number}"


def us_lane_values(
    *,
    tab_name: str,
    origin_country_region: str = "",
    destination_country_region: str = "",
    service_type_package_type: str = "",
    zone_3c: str = "",
    carrier_account_number: str = "",
) -> dict[str, object]:
    return {
        "Tab": tab_name,
        "Origin Country Region": origin_country_region,
        "Destination Country Region": destination_country_region,
        "Service Type/Package Type": service_type_package_type,
        "Zone 3C": zone_3c,
        "Carrier Account Number": carrier_account_number,
    }


def resolve_fred_service(service_type: str, package_type: str) -> str:
    normalized_service = cell_text(service_type).upper()
    normalized_package = cell_text(package_type).upper()
    if normalized_service in {"IE OR CX", "IE/CX"}:
        normalized_service = "IE"

    service_key = (normalized_service, normalized_package)
    if service_key in FRED_SERVICE_BY_RATE_CARD:
        return FRED_SERVICE_BY_RATE_CARD[service_key]
    if (normalized_service, "") in FRED_SERVICE_BY_RATE_CARD:
        return FRED_SERVICE_BY_RATE_CARD[(normalized_service, "")]
    return ""


def list_us_input_files() -> list[Path]:
    return [
        path
        for path in sorted(INPUT_DIR.iterdir())
        if path.is_file()
        and path.suffix.lower() in EXCEL_SUFFIXES
        and not path.name.startswith("~$")
        and US_FILE_PATTERN.search(path.name)
    ]


def select_us_input_file(files: list[Path], *, auto: bool = False) -> Path:
    if not files:
        raise FileNotFoundError(
            f"No DHL US rate files found in {INPUT_DIR}. Expected a workbook matching 'DHL US Rates'."
        )
    if auto or len(files) == 1:
        print(f"\nUsing US input file: {files[0].name}")
        return files[0]
    labels = [path.name for path in files]
    return files[_prompt_selection("Select US input file:", labels)]


def _prompt_selection(title: str, items: list[str]) -> int:
    print(f"\n{title}")
    for index, item in enumerate(items, start=1):
        print(f"  {index}. {item}")
    while True:
        raw = input("Enter number: ").strip()
        try:
            choice = int(raw) - 1
        except ValueError:
            print("Please enter a valid number.")
            continue
        if 0 <= choice < len(items):
            return choice
        print("Please enter a number from the list.")


def load_dhlc_destination_zones(file_path: Path) -> dict[int, list[str]]:
    sheet_df = read_sheet_raw(file_path, DHLC_DESTINATION_ZONES_SHEET)
    return parse_country_zone_table(
        sheet_df,
        table_title="DHL EXPRESS WORLDWIDE DESTINATION ZONE",
    )


def load_dhl_origin_zones(file_path: Path) -> dict[int, list[str]]:
    sheet_df = read_sheet_raw(file_path, DHL_ORIGIN_ZONES_SHEET)
    return parse_country_zone_table(
        sheet_df,
        table_title="DHL EXPRESS WORLDWIDE DESTINATION ZONE",
    )


def load_dhl_import_local_zones(file_path: Path) -> dict[int, list[str]]:
    sheet_df = read_sheet_raw(file_path, DHL_ORIGIN_ZONES_SHEET)
    return parse_country_zone_table(
        sheet_df,
        table_title="DHL EXPRESS WORLDWIDE (IMPORT) DESTINATION LOCAL ZONE",
    )


def load_3c_zone_letter_map(file_path: Path) -> dict[str, list[tuple[int, int]]]:
    sheet_df = read_sheet_raw(file_path, DHL_3C_ZONES_SHEET)
    label_row_index = None
    for row_index, row in sheet_df.iterrows():
        row_values = [cell_text(value) for value in row.tolist()]
        if "Origin Zone" in row_values and "Destination Zone" in row_values:
            label_row_index = row_index
            break
    if label_row_index is None:
        return {}

    label_row = sheet_df.iloc[label_row_index]
    number_row = sheet_df.iloc[label_row_index + 1]
    origin_zone_col = None
    for col_index, value in enumerate(label_row.tolist()):
        if cell_text(value) == "Origin Zone":
            origin_zone_col = col_index
            break
    if origin_zone_col is None:
        return {}

    dest_zone_by_col: dict[int, int] = {}
    for col_index, value in enumerate(number_row.tolist()):
        text = cell_text(value)
        if text.isdigit():
            dest_zone_by_col[col_index] = int(text)

    letter_map: dict[str, list[tuple[int, int]]] = defaultdict(list)
    for row_index in range(label_row_index + 2, len(sheet_df)):
        row = sheet_df.iloc[row_index]
        origin_text = cell_text(row.iloc[origin_zone_col])
        if not origin_text.isdigit():
            continue
        origin_zone = int(origin_text)
        for col_index, dest_zone in dest_zone_by_col.items():
            letter = cell_text(row.iloc[col_index]).upper()
            if not letter or letter == "-":
                continue
            letter_map[letter].append((origin_zone, dest_zone))
    return dict(letter_map)


def _find_rate_region_columns(sheet_df: pd.DataFrame, header_row: int) -> list[tuple[int, str]]:
    regions: list[tuple[int, str]] = []
    region_row = header_row - 1
    if region_row < 0:
        return regions
    for col_index, value in enumerate(sheet_df.iloc[region_row].tolist()):
        label = cell_text(value)
        if label:
            regions.append((col_index, label))
    return regions


def _parse_aos_accounts(sheet_df: pd.DataFrame) -> str:
    accounts: list[str] = []
    for _, row in sheet_df.iterrows():
        label = cell_text(row.iloc[14]) if len(row) > 14 else ""
        if label.lower() == "number":
            continue
        account = cell_text(row.iloc[14]) if len(row) > 14 else ""
        if account.isdigit():
            accounts.append(account)
    return "; ".join(accounts)


def process_dhlg_bbx_tab(
    *,
    tab_name: str,
    sheet_df: pd.DataFrame,
    valid_from: str,
    valid_to: str,
    aos_accounts: str = "",
) -> USLaneBuilder:
    builder = USLaneBuilder()
    located = locate_rate_header(sheet_df)
    if located is None:
        return builder
    header_row_index, column_map = located
    weight_col = weight_column_index(column_map)
    service_col = column_map.get("Service Name/Code", column_map.get("Service Type", 4))
    region_columns = _find_rate_region_columns(sheet_df, header_row_index)
    if not region_columns:
        rate_columns = rate_columns_from_header_row(
            sheet_df.iloc[header_row_index],
            weight_col=weight_col,
        )
        region_columns = [(col_index, label) for col_index, label in rate_columns]

    for row_index in range(header_row_index + 1, len(sheet_df)):
        row = sheet_df.iloc[row_index]
        service_name = cell_text(row.iloc[service_col])
        if not service_name:
            continue
        bracket = weight_bracket_label(row.iloc[weight_col]) if weight_col is not None else ""
        if not bracket:
            continue
        fred_service = resolve_fred_service(service_name, "")
        for col_index, region_label in region_columns:
            rate = rate_value(row.iloc[col_index])
            if rate is None:
                continue
            key = (
                tab_name,
                region_label,
                "US",
                service_name,
                fred_service,
                aos_accounts,
            )
            builder.add_rate(
                key=key,
                lane_values=us_lane_values(
                    tab_name=tab_name,
                    origin_country_region=region_label,
                    service_type_package_type=service_name,
                    carrier_account_number=aos_accounts,
                ),
                cost_group=USCostGroup.DHLG,
                bracket=bracket,
                rate=rate,
            )
    return builder


def _iter_dhlc_sections(sheet_df: pd.DataFrame) -> list[dict[str, object]]:
    sections: list[dict[str, object]] = []
    current_section: dict[str, object] | None = None
    for row_index, row in sheet_df.iterrows():
        row_texts = [cell_text(value) for value in row.tolist()]
        title = next((text for text in row_texts if text), "")
        if not title:
            continue
        lowered = title.lower()
        if "envelope" in lowered:
            current_section = {
                "title": title,
                "cost_group": USCostGroup.DHLC_ENVELOPE,
                "header_row": None,
            }
            sections.append(current_section)
            continue
        if "document" in lowered or "non-document" in lowered:
            current_section = {
                "title": title,
                "cost_group": USCostGroup.DHLC_DOCS,
                "header_row": None,
            }
            sections.append(current_section)
            continue
        if "Origin Point" in row_texts and current_section is not None:
            current_section["header_row"] = row_index
    return [section for section in sections if section.get("header_row") is not None]


def process_dhlc_exp_tab(
    *,
    tab_name: str,
    sheet_df: pd.DataFrame,
    destination_zones: dict[int, list[str]],
    valid_from: str,
    valid_to: str,
) -> USLaneBuilder:
    builder = USLaneBuilder()
    for section in _iter_dhlc_sections(sheet_df):
        header_row_index = int(section["header_row"])
        header_row = sheet_df.iloc[header_row_index]
        column_map = {
            cell_text(value): col_index
            for col_index, value in enumerate(header_row.tolist())
            if cell_text(value)
        }
        weight_col = weight_column_index(column_map)
        origin_country_col = column_map.get("Origin Country", 2)
        service_col = column_map.get("Service Type", 4)
        package_col = column_map.get("Package Type", 5)
        zone_columns = rate_columns_from_header_row(header_row, weight_col=weight_col)

        for row_index in range(header_row_index + 1, len(sheet_df)):
            row = sheet_df.iloc[row_index]
            if cell_text(row.iloc[column_map["Origin Point"]]) == "Origin Point":
                break
            origin_country = cell_text(row.iloc[origin_country_col]) or "US"
            service_type = cell_text(row.iloc[service_col])
            package_type = cell_text(row.iloc[package_col])
            if not service_type and not package_type:
                continue
            bracket = weight_bracket_label(row.iloc[weight_col]) if weight_col is not None else ""
            if not bracket:
                continue
            fred_service = resolve_fred_service(service_type, package_type)
            for col_index, zone_label in zone_columns:
                rate = rate_value(row.iloc[col_index])
                if rate is None:
                    continue
                try:
                    zone_number = int(float(zone_label))
                except ValueError:
                    continue
                key = (
                    tab_name,
                    origin_country,
                    str(zone_number),
                    service_type,
                    package_type,
                    fred_service,
                )
                builder.add_rate(
                    key=key,
                    lane_values=us_lane_values(
                        tab_name=tab_name,
                        origin_country_region=origin_country,
                        service_type_package_type=service_type,
                    ),
                    cost_group=section["cost_group"],
                    bracket=bracket,
                    rate=rate,
                )
    return builder


def process_dhl_3c_tab(
    *,
    tab_name: str,
    sheet_df: pd.DataFrame,
    zone_letter_map: dict[str, list[tuple[int, int]]],
    valid_from: str,
    valid_to: str,
) -> USLaneBuilder:
    builder = USLaneBuilder()
    located = locate_rate_header(sheet_df)
    if located is None:
        return builder
    header_row_index, column_map = located
    weight_col = weight_column_index(column_map)
    service_col = column_map.get("Service Type", 3)
    package_col = column_map.get("Package Type", 4)
    zone_columns = rate_columns_from_header_row(
        sheet_df.iloc[header_row_index],
        weight_col=weight_col,
    )
    for row_index in range(header_row_index + 1, len(sheet_df)):
        row = sheet_df.iloc[row_index]
        service_type = cell_text(row.iloc[service_col])
        package_type = cell_text(row.iloc[package_col])
        if not service_type:
            continue
        bracket = weight_bracket_label(row.iloc[weight_col]) if weight_col is not None else ""
        if not bracket:
            continue
        fred_service = resolve_fred_service(service_type, package_type)
        for col_index, zone_label in zone_columns:
            rate = rate_value(row.iloc[col_index])
            if rate is None:
                continue
            zone_letter = cell_text(zone_label).replace("Zone ", "").strip().upper()
            zone_pairs = zone_letter_map.get(zone_letter)
            if not zone_pairs:
                continue
            for origin_zone, destination_zone in zone_pairs:
                key = (
                    tab_name,
                    str(origin_zone),
                    str(destination_zone),
                    zone_letter,
                    service_type,
                    package_type,
                    fred_service,
                )
                builder.add_rate(
                    key=key,
                    lane_values=us_lane_values(
                        tab_name=tab_name,
                        origin_country_region=format_3c_zone_label(origin_zone),
                        destination_country_region=format_3c_zone_label(destination_zone),
                        service_type_package_type=service_type,
                        zone_3c=zone_letter,
                    ),
                    cost_group=USCostGroup.DHL_3C,
                    bracket=bracket,
                    rate=rate,
                )
    return builder


def process_dhl_ib_tab(
    *,
    tab_name: str,
    sheet_df: pd.DataFrame,
    origin_zones: dict[int, list[str]],
    valid_from: str,
    valid_to: str,
) -> USLaneBuilder:
    builder = USLaneBuilder()
    located = locate_rate_header(sheet_df)
    if located is None:
        return builder
    header_row_index, column_map = located
    weight_col = weight_column_index(column_map)
    service_col = column_map.get("Service Type", 4)
    package_col = column_map.get("Package Type", 5)
    origin_zone_numbers = sorted(origin_zones)
    origin_zone_set = set(origin_zone_numbers)
    zone_columns = rate_columns_from_header_row(
        sheet_df.iloc[header_row_index],
        weight_col=weight_col,
    )
    for row_index in range(header_row_index + 1, len(sheet_df)):
        row = sheet_df.iloc[row_index]
        service_type = cell_text(row.iloc[service_col])
        package_type = cell_text(row.iloc[package_col])
        if not service_type:
            continue
        bracket = weight_bracket_label(row.iloc[weight_col]) if weight_col is not None else ""
        if not bracket:
            continue
        fred_service = resolve_fred_service(service_type, package_type)
        for col_index, zone_label in zone_columns:
            try:
                origin_zone = int(float(zone_label))
            except ValueError:
                continue
            if origin_zone not in origin_zone_set:
                continue
            rate = rate_value(row.iloc[col_index])
            if rate is None:
                continue
            key = (
                tab_name,
                str(origin_zone),
                DHL_IB_DESTINATION_REGION,
                service_type,
                package_type,
                fred_service,
            )
            builder.add_rate(
                key=key,
                lane_values=us_lane_values(
                    tab_name=tab_name,
                    origin_country_region=format_3c_zone_label(origin_zone),
                    destination_country_region=DHL_IB_DESTINATION_REGION,
                    service_type_package_type=service_type,
                ),
                cost_group=USCostGroup.DHL_IB,
                bracket=bracket,
                rate=rate,
            )
    return builder


def _merge_builders(*builders: USLaneBuilder) -> USLaneBuilder:
    merged = USLaneBuilder()
    for builder in builders:
        merged.brackets.update(builder.brackets)
        for group_label, brackets in builder.cost_group_brackets.items():
            merged.cost_group_brackets.setdefault(group_label, set()).update(brackets)
        for key, lane in builder.lanes.items():
            if key not in merged.lanes:
                merged.lanes[key] = dict(lane)
                continue
            for column, value in lane.items():
                if column.startswith("Transport cost"):
                    existing = merged.lanes[key].get(column)
                    if existing is not None and existing != value:
                        raise ValueError(
                            f"Conflicting rates for lane {key} at {column}: {existing} vs {value}"
                        )
                    merged.lanes[key][column] = value
                elif not merged.lanes[key].get(column) and value:
                    merged.lanes[key][column] = value
    return merged


def collect_us_cost_columns(merged: USLaneBuilder) -> list[str]:
    columns: list[str] = []
    for cost_group in US_COST_GROUP_ORDER:
        group_label = us_cost_group_label(cost_group)
        brackets = merged.cost_group_brackets.get(group_label, set())
        for bracket in sorted(brackets, key=bracket_sort_key):
            columns.append(transport_cost_column_name(group_label, bracket))
    return columns


def default_us_output_path(source_file: Path) -> Path:
    stem = source_file.stem
    for suffix in ("_us_matrix", "_matrix"):
        if stem.endswith(suffix):
            stem = stem[: -len(suffix)]
    return OUTPUT_DIR / f"{stem}_matrix.xlsx"


def build_us_matrix_dataframe(file_path: Path) -> tuple[pd.DataFrame, dict[str, str]]:
    valid_from, valid_to = load_index_validity(file_path, INDEX_SHEET)
    destination_zones = load_dhlc_destination_zones(file_path)
    origin_zones = load_dhl_origin_zones(file_path)
    zone_letter_map = load_3c_zone_letter_map(file_path)
    builders: list[USLaneBuilder] = []

    for tab_config in US_RATE_TAB_CONFIGS:
        sheet_df = read_sheet_raw(file_path, tab_config.sheet_name)
        if tab_config.processor == "dhlg_bbx":
            builders.append(
                process_dhlg_bbx_tab(
                    tab_name=tab_config.sheet_name,
                    sheet_df=sheet_df,
                    valid_from=valid_from,
                    valid_to=valid_to,
                )
            )
        elif tab_config.processor == "dhlg_bbx_aos":
            aos_accounts = _parse_aos_accounts(sheet_df)
            builders.append(
                process_dhlg_bbx_tab(
                    tab_name=tab_config.sheet_name,
                    sheet_df=sheet_df,
                    valid_from=valid_from,
                    valid_to=valid_to,
                    aos_accounts=aos_accounts,
                )
            )
        elif tab_config.processor == "dhlc_exp":
            builders.append(
                process_dhlc_exp_tab(
                    tab_name=tab_config.sheet_name,
                    sheet_df=sheet_df,
                    destination_zones=destination_zones,
                    valid_from=valid_from,
                    valid_to=valid_to,
                )
            )
        elif tab_config.processor == "dhl_3c":
            builders.append(
                process_dhl_3c_tab(
                    tab_name=tab_config.sheet_name,
                    sheet_df=sheet_df,
                    zone_letter_map=zone_letter_map,
                    valid_from=valid_from,
                    valid_to=valid_to,
                )
            )
        elif tab_config.processor == "dhl_ib":
            builders.append(
                process_dhl_ib_tab(
                    tab_name=tab_config.sheet_name,
                    sheet_df=sheet_df,
                    origin_zones=origin_zones,
                    valid_from=valid_from,
                    valid_to=valid_to,
                )
            )

    merged = _merge_builders(*builders)
    ordered_brackets = sorted(merged.brackets, key=bracket_sort_key)
    cost_columns = collect_us_cost_columns(merged)
    if not merged.lanes:
        return pd.DataFrame(columns=[*US_SHIPMENT_COLUMNS, CURRENCY_COLUMN, *cost_columns]), {}

    rows = []
    for index, lane in enumerate(merged.lanes.values(), start=1):
        row = {column: lane.get(column, "") for column in US_SHIPMENT_COLUMNS}
        row["Lane #"] = index
        row[CURRENCY_COLUMN] = lane.get(CURRENCY_COLUMN, "USD")
        for cost_column in cost_columns:
            row[cost_column] = lane.get(cost_column)
        rows.append(row)

    columns = [*US_SHIPMENT_COLUMNS, CURRENCY_COLUMN, *cost_columns]
    bracket_rate_by = {bracket: "Flat" for bracket in ordered_brackets}
    return pd.DataFrame(rows)[columns], bracket_rate_by


def run_build_us_matrix(
    *,
    source_file: Path | None = None,
    output_path: Path | None = None,
    auto: bool = False,
) -> Path:
    ensure_workspace_dirs()
    file_path = source_file or select_us_input_file(list_us_input_files(), auto=auto)
    print(f"\nBuilding US matrix from {file_path.name}:")
    for tab_config in US_RATE_TAB_CONFIGS:
        print(f"  - {tab_config.sheet_name}")

    matrix_df, bracket_rate_by = build_us_matrix_dataframe(file_path)
    if output_path is None:
        output_path = default_us_output_path(file_path)
    saved_path = save_matrix(
        matrix_df,
        file_path,
        bracket_rate_by=bracket_rate_by,
        output_path=output_path,
        shipment_columns=US_SHIPMENT_COLUMNS,
    )
    print(f"\nSaved US matrix ({len(matrix_df)} lanes) to: {saved_path}")

    from us_zone_countries import export_us_zone_countries

    zone_countries_path = export_us_zone_countries(file_path)
    print(f"Saved US zone/country reference to: {zone_countries_path}")
    return saved_path


def main() -> int:
    try:
        run_build_us_matrix(auto="--auto" in sys.argv)
        return 0
    except KeyboardInterrupt:
        print("\nOperation cancelled.")
        return 130
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
