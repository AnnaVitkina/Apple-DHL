"""Build postal code zone txt files for DHLPUK workbooks from the GB_Zoning tab."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from build_matrix import cell_text, is_puk_rate_file
from project_paths import INPUT_DIR, OUTPUT_DIR, ensure_workspace_dirs

GB_ZONING_SHEET = "GB_Zoning"
POSTAL_CODE_ZONES_SUFFIX = "_postal_code_zones.txt"
PUK_ZONE_ORDER = ("GB", "GB_NI", "GB_HI", "IE")


@dataclass(frozen=True)
class PostalCodeZone:
    name: str
    country: str
    postal_codes: tuple[str, ...]

    def as_text_block(self) -> str:
        postal_value = ", ".join(self.postal_codes)
        return (
            f"Name: {self.name}\n"
            f"Country: {self.country}\n"
            f"Postal Code: {postal_value}"
        )


def matching_input_workbook(extracted_path: Path) -> Path | None:
    stem = extracted_path.stem
    if stem.endswith("_extracted"):
        stem = stem[: -len("_extracted")]

    for directory in (INPUT_DIR, extracted_path.parent):
        for suffix in (".xlsx", ".xlsm", ".xls"):
            candidate = directory / f"{stem}{suffix}"
            if candidate.is_file():
                return candidate
    return None


def load_gb_zoning_sheet(workbook_path: Path) -> pd.DataFrame | None:
    workbook = pd.ExcelFile(workbook_path)
    if GB_ZONING_SHEET not in workbook.sheet_names:
        return None
    return pd.read_excel(workbook_path, sheet_name=GB_ZONING_SHEET, header=None)


def _find_gb_zoning_header_row(df_raw: pd.DataFrame) -> tuple[int, dict[str, int]] | None:
    scan_limit = min(len(df_raw), 30)
    for row_idx in range(scan_limit):
        labels = [cell_text(value) for value in df_raw.iloc[row_idx].tolist()]
        normalized = {label.casefold(): index for index, label in enumerate(labels) if label}
        if {"country code", "postal code", "lane code"}.issubset(normalized):
            return row_idx, {
                "country": normalized["country code"],
                "postal": normalized["postal code"],
                "lane": normalized["lane code"],
            }
    return None


def build_postal_code_zones_from_gb_zoning(df_raw: pd.DataFrame) -> list[PostalCodeZone]:
    header = _find_gb_zoning_header_row(df_raw)
    if header is None:
        return []

    header_row_idx, columns = header
    grouped: dict[str, tuple[str, list[str]]] = {}

    for row_idx in range(header_row_idx + 1, len(df_raw)):
        row = df_raw.iloc[row_idx]
        country = cell_text(row.iloc[columns["country"]])
        postal_code = cell_text(row.iloc[columns["postal"]])
        zone_name = cell_text(row.iloc[columns["lane"]])
        if not zone_name or not postal_code:
            continue

        if zone_name not in grouped:
            grouped[zone_name] = (country, [])
        elif grouped[zone_name][0] and country and grouped[zone_name][0] != country:
            raise ValueError(
                f"GB_Zoning zone {zone_name} has conflicting countries: "
                f"{grouped[zone_name][0]} vs {country}"
            )
        grouped[zone_name][1].append(postal_code)

    zones: list[PostalCodeZone] = []
    ordered_names = [
        *PUK_ZONE_ORDER,
        *sorted(name for name in grouped if name not in PUK_ZONE_ORDER),
    ]
    seen: set[str] = set()
    for zone_name in ordered_names:
        if zone_name in seen or zone_name not in grouped:
            continue
        seen.add(zone_name)
        country, postal_codes = grouped[zone_name]
        zones.append(
            PostalCodeZone(
                name=zone_name,
                country=country or "GB",
                postal_codes=tuple(postal_codes),
            )
        )
    return zones


def default_postal_zones_path(source_file: Path) -> Path:
    stem = source_file.stem.replace("_extracted", "")
    return OUTPUT_DIR / f"{stem}{POSTAL_CODE_ZONES_SUFFIX}"


def write_postal_code_zones_txt(zones: list[PostalCodeZone], output_path: Path) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    blocks = [zone.as_text_block() for zone in zones]
    output_path.write_text("\n\n".join(blocks) + ("\n" if blocks else ""), encoding="utf-8")
    return output_path


def run_build_postal_code_zones(
    *,
    extracted_file: Path,
    output_path: Path | None = None,
) -> Path | None:
    if not is_puk_rate_file(extracted_file):
        return None

    workbook_path = matching_input_workbook(extracted_file)
    if workbook_path is None:
        print("  Postal code zones: no matching input workbook found for GB_Zoning")
        return None

    df_raw = load_gb_zoning_sheet(workbook_path)
    if df_raw is None:
        print(f"  Postal code zones: no '{GB_ZONING_SHEET}' sheet in {workbook_path.name}")
        return None

    zones = build_postal_code_zones_from_gb_zoning(df_raw)
    if not zones:
        print(f"  Postal code zones: no zones parsed from {workbook_path.name}")
        return None

    saved_path = write_postal_code_zones_txt(
        zones,
        output_path or default_postal_zones_path(extracted_file),
    )
    print(f"  Postal code zones: wrote {len(zones)} zone(s) to {saved_path.name}")
    for zone in zones:
        print(f"    - {zone.name}: {len(zone.postal_codes)} postal code(s)")
    return saved_path
