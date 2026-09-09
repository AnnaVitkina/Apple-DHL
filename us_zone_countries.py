"""Export per-tab zone-to-country reference files for DHL US rates."""

from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path

from build_us_matrix import (
    _find_rate_region_columns,
    list_us_input_files,
    load_dhl_origin_zones,
    load_dhlc_destination_zones,
    select_us_input_file,
)
from project_paths import OUTPUT_DIR, ensure_workspace_dirs
from us_config import (
    DHL_3C_ZONES_SHEET,
    DHL_IB_DESTINATION_REGION,
    DHL_ORIGIN_ZONES_SHEET,
    US_RATE_TAB_CONFIGS,
)
from us_excel import (
    load_3c_country_zones,
    load_dhl_ib_destination_countries,
    locate_rate_header,
    parse_country_code,
    read_sheet_raw,
)


US_DESTINATION_COUNTRIES = ("US",)


def _countries_to_iso(countries: list[str]) -> list[str]:
    iso_codes: list[str] = []
    for country in countries:
        code = parse_country_code(country)
        if code and code not in iso_codes:
            iso_codes.append(code)
    return iso_codes


def _iso_zone_map(zone_map: dict[str, list[str]]) -> dict[str, list[str]]:
    return {label: _countries_to_iso(countries) for label, countries in zone_map.items()}


def _parse_dhlg_region_iso_codes(region_label: str) -> list[str]:
    left = re.split(r"\s+to\s+US\b", region_label.strip(), flags=re.IGNORECASE)[0].strip()
    if not left:
        return []

    iso_codes: list[str] = []
    for token in re.split(r"[\s&/,]+", left):
        code = token.strip().upper()
        if len(code) == 2 and code.isalpha() and code != "US" and code not in iso_codes:
            iso_codes.append(code)
    return iso_codes


def _zone_sort_key(zone_label: str) -> tuple[int, str | int]:
    text = zone_label.strip()
    if text.upper().startswith("ZONE "):
        suffix = text[5:].strip()
        try:
            return (0, int(suffix))
        except ValueError:
            return (1, suffix)
    return (2, text)


def _format_zone_line(zone_label: str, countries: list[str]) -> str:
    if countries:
        return f"{zone_label} - {', '.join(countries)}"
    return f"{zone_label} -"


def _format_zone_map(zone_map: dict[str, list[str]]) -> list[str]:
    return [
        _format_zone_line(zone_label, zone_map[zone_label])
        for zone_label in sorted(zone_map, key=_zone_sort_key)
    ]


def _collect_dhlg_zones(file_path: Path, tab_name: str) -> tuple[dict[str, list[str]], dict[str, list[str]]]:
    sheet_df = read_sheet_raw(file_path, tab_name)
    located = locate_rate_header(sheet_df)
    origin: dict[str, list[str]] = {}
    if located is not None:
        for _, region_label in _find_rate_region_columns(sheet_df, located[0]):
            origin[region_label] = _parse_dhlg_region_iso_codes(region_label)
    destination = {"US": list(US_DESTINATION_COUNTRIES)}
    return origin, destination


def _collect_dhlc_zones(file_path: Path) -> tuple[dict[str, list[str]], dict[str, list[str]]]:
    origin = {"US": list(US_DESTINATION_COUNTRIES)}
    destination = _iso_zone_map({
        f"Zone {zone_number}": countries
        for zone_number, countries in sorted(load_dhlc_destination_zones(file_path).items())
    })
    return origin, destination


def _collect_3c_zones(file_path: Path) -> tuple[dict[str, list[str]], dict[str, list[str]]]:
    sheet_df = read_sheet_raw(file_path, DHL_3C_ZONES_SHEET)
    zone_countries = load_3c_country_zones(sheet_df)
    zone_map = _iso_zone_map({
        f"Zone {zone_number}": countries
        for zone_number, countries in sorted(zone_countries.items())
    })
    return zone_map, zone_map


def _collect_ib_zones(file_path: Path) -> tuple[dict[str, list[str]], dict[str, list[str]]]:
    origin = _iso_zone_map({
        f"Zone {zone_number}": countries
        for zone_number, countries in sorted(load_dhl_origin_zones(file_path).items())
    })
    origin_zones_df = read_sheet_raw(file_path, DHL_ORIGIN_ZONES_SHEET)
    destination = _iso_zone_map({
        DHL_IB_DESTINATION_REGION: load_dhl_ib_destination_countries(origin_zones_df),
    })
    return origin, destination


def collect_tab_zone_countries(file_path: Path) -> dict[str, dict[str, dict[str, list[str]]]]:
    dhlc_origin, dhlc_destination = _collect_dhlc_zones(file_path)
    origin_3c, destination_3c = _collect_3c_zones(file_path)
    ib_origin, ib_destination = _collect_ib_zones(file_path)

    tab_data: dict[str, dict[str, dict[str, list[str]]]] = {}
    for tab_config in US_RATE_TAB_CONFIGS:
        processor = tab_config.processor
        if processor in {"dhlg_bbx", "dhlg_bbx_aos"}:
            origin, destination = _collect_dhlg_zones(file_path, tab_config.sheet_name)
        elif processor == "dhlc_exp":
            origin, destination = dhlc_origin, dhlc_destination
        elif processor == "dhl_3c":
            origin, destination = origin_3c, destination_3c
        elif processor == "dhl_ib":
            origin, destination = ib_origin, ib_destination
        else:
            origin, destination = {}, {}

        tab_data[tab_config.sheet_name] = {
            "origin": origin,
            "destination": destination,
        }
    return tab_data


def format_zone_countries_text(
    file_path: Path,
    tab_data: dict[str, dict[str, dict[str, list[str]]]],
) -> str:
    lines = [
        "US ZONE / COUNTRY REFERENCE",
        f"Source: {file_path.name}",
        f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        "",
    ]

    for tab_name, zones in tab_data.items():
        lines.append(f"TAB: {tab_name}")
        lines.append("ORIGIN")
        lines.extend(_format_zone_map(zones["origin"]))
        lines.append("DESTINATION")
        lines.extend(_format_zone_map(zones["destination"]))
        lines.append("")

    return "\n".join(lines).rstrip() + "\n"


def default_zone_countries_output_path(source_file: Path) -> Path:
    stem = source_file.stem
    for suffix in ("_zone_countries", "_us_matrix", "_matrix"):
        if stem.endswith(suffix):
            stem = stem[: -len(suffix)]
    return OUTPUT_DIR / f"{stem}_zone_countries.txt"


def export_us_zone_countries(
    file_path: Path,
    *,
    output_path: Path | None = None,
) -> Path:
    ensure_workspace_dirs()
    tab_data = collect_tab_zone_countries(file_path)
    if output_path is None:
        output_path = default_zone_countries_output_path(file_path)

    output_path.write_text(
        format_zone_countries_text(file_path, tab_data),
        encoding="utf-8",
    )
    return output_path


def run_export_us_zone_countries(
    *,
    source_file: Path | None = None,
    output_path: Path | None = None,
    auto: bool = False,
) -> Path:
    ensure_workspace_dirs()
    file_path = source_file or select_us_input_file(list_us_input_files(), auto=auto)
    saved_path = export_us_zone_countries(file_path, output_path=output_path)
    print(f"\nSaved US zone/country reference to: {saved_path}")
    return saved_path


if __name__ == "__main__":
    import sys

    raise SystemExit(0 if run_export_us_zone_countries(auto="--auto" in sys.argv) else 1)
