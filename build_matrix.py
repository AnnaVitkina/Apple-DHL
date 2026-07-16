"""
Build a matrix workbook from extracted DHL processing dataframes.

Matrix layout:
  - Shipment info (lane details)
  - Transport cost columns with multi-row headers (weight brackets)
"""

from __future__ import annotations

import re
import sys
from datetime import date, datetime
from dataclasses import dataclass
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from project_paths import INPUT_DIR, OUTPUT_DIR, PROCESSING_DIR, ensure_workspace_dirs

EXTRACTED_GLOB = "*_extracted.xlsx"
ISO_COLUMN_PATTERN = re.compile(r"^[A-Z]{2}$")
WEIGHT_RANGE_PATTERN = re.compile(r"^(\d+)\s*(?:to|-)\s*(\d+)$", re.IGNORECASE)

SHIPMENT_COLUMNS = (
    "Lane #",
    "Tab",
    "Origin country",
    "Shipment Type",
    "Service",
    "Service level",
    "Destination country",
    "Destination Postal Code",
    "Carrier Account number",
    "Business Segment",
    "Category",
    "Shipping Condition",
    "Valid from",
    "Valid to",
)

CURRENCY_COLUMN = "Currency"
TAB_INDEX_SHEET = "Tab Index"
RETURN_DESCRIPTION_PATTERN = re.compile(r"shipments\s+back", re.IGNORECASE)
RETURN_CATEGORY = "return"

TRANSPORT_COST_GROUP = "Transport cost"
SFS_TRANSPORT_COST_GROUP = "Transport cost (SFS)"
SFS_TAB_NAME = "SFS"
EXCLUDED_RATE_TABS = {"tab index", "accessorials"}
BILLING_BS_GLOB = "billing-bs*"
CARRIER_ACCOUNT_SPLIT_PATTERN = re.compile(r"[,/;]+")
CARRIER_ACCOUNT_TOKEN_PATTERN = re.compile(r"\b[A-Za-z0-9]+\b")
BUSINESS_SEGMENT_ABBREVIATIONS: dict[str, str] = {
    "finished goods": "FG",
    "fg": "FG",
    "applecare": "AC",
    "apple care": "AC",
    "finished goods special billing": "FGSB",
    "applecare special billing": "ACSB",
    "secureship": "SS",
    "secureship special billing": "SSSB",
    "manufacturing": "MFG",
    "seed lab": "SL",
    "corporate": "CORP",
    "apple retail": "AR",
    "charter - special billing account": "CHARTER",
    "special billing": "SB",
}

FIXED_SOURCE_COLUMNS = {
    "Version Number",
    "Version Date",
    "Destination Country",
    "Destination ISO",
    "Billing Currency",
    "Origin",
    "Shipment Type",
    "Service",
    "Service Level",
    "Chargeable Weight",
    "Rate Logic",
    "Lane Code",
}

COST_GROUP_ROW = 1
COST_NAME_ROW = 2
APPLY_IF_ROW = 3
RATE_BY_ROW = 4
COLUMN_HEADER_ROW = 5
DATA_START_ROW = 6

HEADER_FILL = PatternFill("solid", fgColor="D9D9D9")
TRANSPORT_GROUP_FILL = PatternFill("solid", fgColor="9BC2E6")
TRANSPORT_COST_FILL = PatternFill("solid", fgColor="BDD7EE")
COST_META_FILL = PatternFill("solid", fgColor="F2F2F2")
BOLD = Font(bold=True)
NORMAL = Font()
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
RATE_NUMBER_FORMAT = "#,##0.00"
THIN_BORDER = Border(
    left=Side(style="thin", color="B4B4B4"),
    right=Side(style="thin", color="B4B4B4"),
    top=Side(style="thin", color="B4B4B4"),
    bottom=Side(style="thin", color="B4B4B4"),
)

EU_COUNTRIES = frozenset({
    "AT", "BE", "BG", "HR", "CY", "CZ", "DK", "EE", "FI", "FR", "DE", "GR",
    "HU", "IE", "IT", "LV", "LT", "LU", "MT", "NL", "PL", "PT", "RO", "SK",
    "SI", "ES", "SE",
})
ES_WPX_TERRITORIES = frozenset({"ES", "IC", "AD"})
DEPOT_TO_COUNTRY = {
    "LEJ": "DE",
    "CGO": "CN",
    "TYN": "CN",
    "CTU": "CN",
    "SHA": "CN",
    "PVG": "CN",
    "NKG": "CN",
    "SZX": "CN",
    "HKG": "HK",
    "7101": "IE",
}
LU_DEPOT_PATTERN = re.compile(r"^LU\d+$", re.IGNORECASE)
TAB_PREFIX_TO_COUNTRY = {
    "CN_": "CN",
    "VN_": "VN",
}
NSR_SERVICE_LEVELS = ("DOM", "ECX", "WPX")
SERVICE_LEVEL_TOKEN_PATTERN = re.compile(r"[A-Z0-9]+")
ISO_TOKEN_PATTERN = re.compile(r"\b[A-Z]{2}\b")
SERVICE_LEVEL_TO_EXP: dict[str, str] = {
    "DOM": "EXP_DOM",
    "GROUND": "EXP_DOM",
    "WPX": "EXP_WW_NONDOC",
    "ECX": "EXP_WW_EU",
    "DES": "EXP_EC_DOM",
    "ESU": "EXP_EC_EU",
    "ESI": "EXP_EC_NONDOC",
    "DOX": "EXP_WW_DOC",
    "BTC": "EXP_B2C",
    "DOT": "EXP_DOM_1200",
    "SDX": "SAMEDAY",
    "BBX": "EXP_BREAKBULK",
    "TDY": "EXP_1200_NONDOC",
    "TDT": "EXP_1200_DOC",
    "TDE": "EXP_900_NONDOC",
    "TDK": "EXP_900_DOC",
    "DOK": "EXP_DOM_900",
}
UNCONDITIONAL_EXP_SERVICES = frozenset({
    "EXP_DOM_1200",
    "EXP_1200_NONDOC",
    "EXP_1200_DOC",
    "EXP_900_NONDOC",
    "EXP_900_DOC",
    "EXP_DOM_900",
    "EXP_BREAKBULK",
    "SAMEDAY",
})


def cell_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def rate_value(value: object) -> float | int | None:
    if pd.isna(value):
        return None
    text = cell_text(value)
    if not text or text.lower() in {"on request", "n/a", "#n/a"}:
        return None
    try:
        number = float(str(value).replace(",", ""))
    except (TypeError, ValueError):
        return None
    if number.is_integer():
        return int(number)
    return number


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


def rate_logic_to_rate_by(rate_logic: object) -> str:
    text = cell_text(rate_logic).lower()
    if text == "dn rate":
        return "Flat"
    if text == "per-kg rate":
        return "p/unit"
    return ""


def weight_bracket_label(chargeable_weight: object) -> str:
    text = cell_text(chargeable_weight)
    if not text:
        return ""

    range_match = WEIGHT_RANGE_PATTERN.match(text)
    if range_match:
        upper = int(range_match.group(2))
        return f"<={upper}"

    try:
        upper = int(float(text))
    except ValueError:
        return text
    return f"<={upper}"


def bracket_sort_key(label: str) -> tuple[int, int | str]:
    match = re.match(r"^<=(\d+)$", label)
    if match:
        return (0, int(match.group(1)))
    return (1, label)


def transport_cost_column_name(bracket_label: str) -> str:
    return f"{TRANSPORT_COST_GROUP} ({bracket_label})"


def sfs_transport_cost_column_name(bracket_label: str) -> str:
    return f"{SFS_TRANSPORT_COST_GROUP} ({bracket_label})"


def is_sfs_tab(tab_name: str) -> bool:
    return tab_name.strip().upper() == SFS_TAB_NAME


def is_standard_weight_bracket_cost_column(column_name: str) -> bool:
    prefix = f"{TRANSPORT_COST_GROUP} ("
    return column_name.startswith(prefix) and not column_name.startswith(f"{SFS_TRANSPORT_COST_GROUP} (")


def is_sfs_weight_bracket_cost_column(column_name: str) -> bool:
    return column_name.startswith(f"{SFS_TRANSPORT_COST_GROUP} (")


def is_eu_country(country_code: str) -> bool:
    return country_code.upper() in EU_COUNTRIES


def parse_service_level_tokens(service_level: object) -> list[str]:
    text = cell_text(service_level).upper()
    if not text:
        return []

    tokens: list[str] = []
    for match in SERVICE_LEVEL_TOKEN_PATTERN.findall(text):
        if match == "NSR":
            tokens.extend(NSR_SERVICE_LEVELS)
        else:
            tokens.append(match)
    return tokens


def _country_codes_in_text(value: object) -> list[str]:
    return ISO_TOKEN_PATTERN.findall(cell_text(value).upper())


def resolve_lane_country_code(
    country_value: object,
    *,
    tab_name: str = "",
    tab_index_lookup: dict[str, TabIndexInfo] | None = None,
) -> str:
    text = cell_text(country_value)
    if not text or text.lower() == "any":
        return ""

    for code in _country_codes_in_text(text):
        if ISO_COLUMN_PATTERN.match(code):
            return code

    for part in re.split(r"[,/;\s]+", text.upper()):
        if not part:
            continue
        if part in DEPOT_TO_COUNTRY:
            return DEPOT_TO_COUNTRY[part]
        if LU_DEPOT_PATTERN.match(part):
            return "LU"
        if ISO_COLUMN_PATTERN.match(part):
            return part

    tab_name = cell_text(tab_name)
    if ISO_COLUMN_PATTERN.match(tab_name):
        return tab_name

    if tab_index_lookup:
        tab_info = tab_index_lookup.get(tab_name)
        if tab_info is not None and tab_info.origins:
            resolved = resolve_lane_country_code(
                tab_info.origins,
                tab_name=tab_name,
            )
            if resolved:
                return resolved

    for prefix, country_code in TAB_PREFIX_TO_COUNTRY.items():
        if tab_name.upper().startswith(prefix):
            return country_code

    return ""


def spl_service_for_lane(origin: str, destination: str) -> str:
    if origin == "RU" or destination == "RU":
        return "CITYLINE (DHL EXP RU)"
    return "SPRINTLINE (DHL EXP NL)"


def _dom_exp_service_applicable(origin: str, destination: str) -> bool:
    return bool(origin) and origin == destination


def _ecx_exp_service_applicable(origin: str, destination: str) -> bool:
    if not origin or not destination:
        return False
    return origin != destination and is_eu_country(origin) and is_eu_country(destination)


def _wpx_exp_service_applicable(origin: str, destination: str) -> bool:
    if not destination:
        return False
    if not origin:
        return not is_eu_country(destination)
    if origin == destination:
        return False
    if not is_eu_country(origin) or not is_eu_country(destination):
        return True
    return origin in ES_WPX_TERRITORIES or destination in ES_WPX_TERRITORIES


def _esi_exp_service_applicable(origin: str, destination: str) -> bool:
    if not destination:
        return False
    if not origin:
        return not is_eu_country(destination)
    return origin != destination and (not is_eu_country(origin) or not is_eu_country(destination))


def is_exp_service_applicable(exp_service: str, origin: str, destination: str) -> bool:
    origin = origin.upper()
    destination = destination.upper()

    if exp_service == "EXP_DOM":
        return _dom_exp_service_applicable(origin, destination)
    if exp_service == "EXP_WW_EU":
        return _ecx_exp_service_applicable(origin, destination)
    if exp_service == "EXP_WW_NONDOC":
        return _wpx_exp_service_applicable(origin, destination)
    if exp_service == "EXP_EC_DOM":
        return _dom_exp_service_applicable(origin, destination)
    if exp_service == "EXP_EC_EU":
        return _ecx_exp_service_applicable(origin, destination)
    if exp_service == "EXP_EC_NONDOC":
        return _esi_exp_service_applicable(origin, destination)
    if exp_service == "EXP_WW_DOC":
        return (
            _dom_exp_service_applicable(origin, destination)
            or _ecx_exp_service_applicable(origin, destination)
            or _wpx_exp_service_applicable(origin, destination)
        )
    if exp_service == "EXP_B2C":
        return origin == "RU" and destination == "RU"
    if exp_service in UNCONDITIONAL_EXP_SERVICES:
        return True
    return False


def resolve_service_column(
    service_level: object,
    origin: str,
    destination: str,
) -> str:
    applicable: list[str] = []
    seen: set[str] = set()

    for token in parse_service_level_tokens(service_level):
        if token == "SPL":
            exp_service = spl_service_for_lane(origin, destination)
        else:
            exp_service = SERVICE_LEVEL_TO_EXP.get(token)
            if exp_service is None:
                continue

        if exp_service in seen:
            continue
        if is_exp_service_applicable(exp_service, origin, destination):
            applicable.append(exp_service)
            seen.add(exp_service)

    return "/".join(applicable)


def apply_service_column(
    matrix_df: pd.DataFrame,
    tab_index_lookup: dict[str, TabIndexInfo] | None = None,
) -> pd.DataFrame:
    if matrix_df.empty:
        return matrix_df

    result = matrix_df.copy()
    for index, row in result.iterrows():
        origin = resolve_lane_country_code(
            row.get("Origin country"),
            tab_name=cell_text(row.get("Tab")),
            tab_index_lookup=tab_index_lookup,
        )
        destination = resolve_lane_country_code(row.get("Destination country"))
        result.at[index, "Service"] = resolve_service_column(
            row.get("Service level"),
            origin,
            destination,
        )
    return result


def destination_iso_columns(df: pd.DataFrame) -> list[str]:
    return [
        str(column)
        for column in df.columns
        if str(column) not in FIXED_SOURCE_COLUMNS and ISO_COLUMN_PATTERN.match(str(column))
    ]


def is_excluded_matrix_tab(sheet_name: str) -> bool:
    return sheet_name.strip().lower() in EXCLUDED_RATE_TABS


def is_rate_tab(
    sheet_name: str,
    df: pd.DataFrame,
    tab_index_lookup: dict[str, TabIndexInfo] | None = None,
) -> bool:
    if is_excluded_matrix_tab(sheet_name):
        return False
    if tab_index_lookup:
        tab_info = tab_index_lookup.get(sheet_name)
        if tab_info is None or not tab_info.is_active:
            return False
    required = {"Origin", "Chargeable Weight", "Rate Logic"}
    if not required.issubset(df.columns):
        return False
    return bool(destination_iso_columns(df))


def list_extracted_files() -> list[Path]:
    return sorted(
        [
            path
            for path in PROCESSING_DIR.glob(EXTRACTED_GLOB)
            if path.is_file() and not path.name.startswith("~$")
        ],
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )


def prompt_selection(title: str, items: list[str]) -> int:
    print(f"\n{title}")
    for index, item in enumerate(items, start=1):
        print(f"  {index}. {item}")

    while True:
        raw = input("Enter number: ").strip()
        if not raw.isdigit():
            print("Please enter a valid number.")
            continue
        choice = int(raw)
        if 1 <= choice <= len(items):
            return choice - 1
        print("Number is out of range. Try again.")


def select_extracted_file(files: list[Path], *, auto: bool = False) -> Path:
    if not files:
        print(f"No extracted files found in: {PROCESSING_DIR}")
        sys.exit(1)

    if auto or len(files) == 1:
        print(f"\nUsing extracted file: {files[0].name}")
        return files[0]

    labels = [path.name for path in files]
    return files[prompt_selection("Select extracted file to build matrix from:", labels)]


def load_rate_tabs(file_path: Path) -> list[tuple[str, pd.DataFrame]]:
    workbook = pd.ExcelFile(file_path)
    tab_index_lookup = build_tab_index_lookup(load_tab_index(file_path))
    loaded: list[tuple[str, pd.DataFrame]] = []

    for sheet_name in workbook.sheet_names:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        if not is_rate_tab(sheet_name, df, tab_index_lookup):
            continue
        loaded.append((sheet_name, df))

    return loaded


@dataclass(frozen=True)
class TabIndexInfo:
    tab_name: str
    is_active: bool
    has_return_lanes: bool
    billing_accounts: str
    origins: str


def normalize_billing_accounts(value: object) -> str:
    text = cell_text(value).replace("\n", " ")
    parts = [part.strip() for part in re.split(r"\s*;\s*", text) if part.strip()]
    return "; ".join(parts)


def billing_bs_file_path() -> Path | None:
    matches = sorted(INPUT_DIR.glob(BILLING_BS_GLOB))
    return matches[0] if matches else None


def normalize_business_segment_name(name: object) -> str:
    return re.sub(r"\s+", " ", cell_text(name).lower())


def business_segment_abbreviation(segment_name: object) -> str:
    normalized = normalize_business_segment_name(segment_name)
    if not normalized:
        return ""

    mapped = BUSINESS_SEGMENT_ABBREVIATIONS.get(normalized)
    if mapped:
        return mapped

    raw = cell_text(segment_name)
    if raw.isupper() and 1 < len(raw) <= 8:
        return raw

    words = re.findall(r"[a-z0-9]+", normalized)
    if words:
        return "".join(word[0] for word in words).upper()
    return raw


def load_billing_bs_lookup(file_path: Path | None = None) -> dict[str, str]:
    path = file_path or billing_bs_file_path()
    if path is None or not path.exists():
        return {}

    lookup: dict[str, str] = {}
    for line in path.read_text(encoding="utf-8", errors="replace").splitlines():
        if not line.strip():
            continue
        parts = line.split("\t")
        if len(parts) < 2:
            continue
        account_number = cell_text(parts[0])
        business_segment = cell_text(parts[1])
        if account_number and business_segment:
            lookup[account_number] = business_segment
    return lookup


def parse_carrier_account_numbers(
    value: object,
    billing_bs_lookup: dict[str, str] | None = None,
) -> list[str]:
    text = cell_text(value)
    if not text:
        return []

    tokens = CARRIER_ACCOUNT_TOKEN_PATTERN.findall(text)
    if billing_bs_lookup is None:
        return [
            token
            for part in CARRIER_ACCOUNT_SPLIT_PATTERN.split(text)
            for token in [re.split(r"[\s(]+", part.strip())[0].strip()]
            if token
        ]

    accounts: list[str] = []
    seen: set[str] = set()
    for token in tokens:
        if token not in billing_bs_lookup or token in seen:
            continue
        accounts.append(token)
        seen.add(token)
    return accounts


def resolve_business_segment(
    carrier_accounts: object,
    billing_bs_lookup: dict[str, str],
) -> str:
    abbreviations: list[str] = []
    seen: set[str] = set()

    for account_number in parse_carrier_account_numbers(
        carrier_accounts,
        billing_bs_lookup,
    ):
        segment_name = billing_bs_lookup.get(account_number)
        if not segment_name:
            continue
        abbreviation = business_segment_abbreviation(segment_name)
        if abbreviation and abbreviation not in seen:
            abbreviations.append(abbreviation)
            seen.add(abbreviation)

    return "/".join(abbreviations)


def apply_business_segment_column(
    matrix_df: pd.DataFrame,
    billing_bs_lookup: dict[str, str],
) -> pd.DataFrame:
    if matrix_df.empty or not billing_bs_lookup:
        return matrix_df

    result = matrix_df.copy()
    for index, row in result.iterrows():
        result.at[index, "Business Segment"] = resolve_business_segment(
            row.get("Carrier Account number"),
            billing_bs_lookup,
        )
    return result


def load_tab_index(file_path: Path) -> pd.DataFrame | None:
    workbook = pd.ExcelFile(file_path)
    for sheet_name in workbook.sheet_names:
        if sheet_name.strip().lower() == TAB_INDEX_SHEET.lower():
            return pd.read_excel(file_path, sheet_name=sheet_name)
    return None


def build_tab_index_lookup(tab_index_df: pd.DataFrame | None) -> dict[str, TabIndexInfo]:
    if tab_index_df is None or tab_index_df.empty:
        return {}

    lookup: dict[str, TabIndexInfo] = {}
    for _, row in tab_index_df.iterrows():
        tab_name = cell_text(row.get("Tab-Name"))
        if not tab_name:
            continue
        description = cell_text(row.get("Description"))
        status = cell_text(row.get("Status")).upper()
        lookup[tab_name] = TabIndexInfo(
            tab_name=tab_name,
            is_active=status == "ACTIVE",
            has_return_lanes=bool(RETURN_DESCRIPTION_PATTERN.search(description)),
            billing_accounts=normalize_billing_accounts(row.get("Billing Account(s)")),
            origins=cell_text(row.get("Origin(s)")),
        )
    return lookup


def apply_tab_index_metadata(
    matrix_df: pd.DataFrame,
    tab_index_lookup: dict[str, TabIndexInfo],
) -> pd.DataFrame:
    if matrix_df.empty or not tab_index_lookup:
        return matrix_df

    result = matrix_df.copy()
    if "Category" not in result.columns:
        result["Category"] = ""

    for index, row in result.iterrows():
        tab_name = cell_text(row.get("Tab"))
        tab_info = tab_index_lookup.get(tab_name)
        if tab_info is None:
            continue
        if tab_info.billing_accounts:
            result.at[index, "Carrier Account number"] = tab_info.billing_accounts

    return result


def append_return_lanes(
    matrix_df: pd.DataFrame,
    tab_index_lookup: dict[str, TabIndexInfo],
) -> pd.DataFrame:
    return_tabs = {
        tab_name
        for tab_name, tab_info in tab_index_lookup.items()
        if tab_info.has_return_lanes
    }
    if not return_tabs or matrix_df.empty:
        return matrix_df

    outbound = matrix_df[matrix_df["Category"].fillna("").astype(str).str.strip() == ""].copy()
    return_rows: list[dict[str, object]] = []

    for _, row in outbound.iterrows():
        if cell_text(row.get("Tab")) not in return_tabs:
            continue

        duplicate = row.to_dict()
        origin = cell_text(row.get("Origin country"))
        destination = cell_text(row.get("Destination country"))
        duplicate["Origin country"] = destination
        duplicate["Destination country"] = origin
        duplicate["Category"] = RETURN_CATEGORY
        return_rows.append(duplicate)

    if not return_rows:
        return matrix_df

    return pd.concat([matrix_df, pd.DataFrame(return_rows)], ignore_index=True)


def finalize_lane_numbers(matrix_df: pd.DataFrame) -> pd.DataFrame:
    if matrix_df.empty:
        return matrix_df
    result = matrix_df.copy()
    if "Category" in result.columns:
        result["Category"] = result["Category"].fillna("").astype(str).str.strip()
        result.loc[result["Category"].eq("nan"), "Category"] = ""
    result["Lane #"] = range(1, len(result) + 1)
    return result


def enrich_matrix_with_tab_index(
    matrix_df: pd.DataFrame,
    tab_index_df: pd.DataFrame | None,
    *,
    billing_bs_lookup: dict[str, str] | None = None,
) -> pd.DataFrame:
    tab_index_lookup = build_tab_index_lookup(tab_index_df)
    if billing_bs_lookup is None:
        billing_bs_lookup = load_billing_bs_lookup()
    result = apply_tab_index_metadata(matrix_df, tab_index_lookup)
    result = apply_business_segment_column(result, billing_bs_lookup)
    result = append_return_lanes(result, tab_index_lookup)
    result = apply_service_column(result, tab_index_lookup)
    return finalize_lane_numbers(result)


def collect_weight_brackets(
    rate_tabs: list[tuple[str, pd.DataFrame]],
    *,
    sfs_only: bool = False,
) -> tuple[list[str], dict[str, str]]:
    brackets: set[str] = set()
    bracket_rate_by: dict[str, str] = {}

    for tab_name, source_df in rate_tabs:
        if is_sfs_tab(tab_name) != sfs_only:
            continue
        for _, row in source_df.iterrows():
            bracket = weight_bracket_label(row.get("Chargeable Weight"))
            if not bracket:
                continue
            brackets.add(bracket)
            rate_by = rate_logic_to_rate_by(row.get("Rate Logic"))
            if rate_by:
                existing = bracket_rate_by.get(bracket)
                if existing and existing != rate_by:
                    scope = "SFS" if sfs_only else "Transport cost"
                    raise ValueError(
                        f"{scope} weight bracket {bracket} has conflicting rate logic: "
                        f"{existing} vs {rate_by}"
                    )
                bracket_rate_by[bracket] = rate_by

    ordered = sorted(brackets, key=bracket_sort_key)
    return ordered, bracket_rate_by


def collect_transport_cost_columns(
    rate_tabs: list[tuple[str, pd.DataFrame]],
) -> tuple[list[str], list[str], dict[str, str], dict[str, str]]:
    weight_brackets, bracket_rate_by = collect_weight_brackets(rate_tabs, sfs_only=False)
    sfs_brackets, sfs_bracket_rate_by = collect_weight_brackets(rate_tabs, sfs_only=True)
    standard_columns = [transport_cost_column_name(bracket) for bracket in weight_brackets]
    sfs_columns = [sfs_transport_cost_column_name(bracket) for bracket in sfs_brackets]
    return standard_columns, sfs_columns, bracket_rate_by, sfs_bracket_rate_by


def lane_key(
    origin: str,
    destination: str,
    shipment_type: str,
    stream_id: str,
    *,
    tab_name: str = "",
) -> tuple[str, ...]:
    if is_sfs_tab(tab_name):
        return (tab_name, origin, destination, shipment_type, stream_id)
    return (origin, destination, shipment_type, stream_id)


def _is_first_weight_bracket(chargeable_weight: object) -> bool:
    return weight_bracket_label(chargeable_weight) == "<=1"


def iter_rate_blocks(source_df: pd.DataFrame) -> list[pd.DataFrame]:
    """Split a tab into consecutive rate blocks that restart at weight bracket 1."""
    blocks: list[list[pd.Series]] = []
    current_block: list[pd.Series] = []

    for _, row in source_df.iterrows():
        if current_block and _is_first_weight_bracket(row.get("Chargeable Weight")):
            blocks.append(current_block)
            current_block = []
        current_block.append(row)

    if current_block:
        blocks.append(current_block)

    return [pd.DataFrame(block_rows) for block_rows in blocks]


def block_stream_id(block_df: pd.DataFrame) -> str:
    first_row = block_df.iloc[0]
    return cell_text(first_row.get("Service Level")) or cell_text(first_row.get("Shipment Type"))


def build_shipment_and_cost_rows(rate_tabs: list[tuple[str, pd.DataFrame]]) -> pd.DataFrame:
    standard_columns, sfs_columns, _, _ = collect_transport_cost_columns(rate_tabs)
    transport_columns = [*standard_columns, *sfs_columns]
    lanes: dict[tuple[str, ...], dict[str, object]] = {}

    for tab_name, source_df in rate_tabs:
        iso_columns = destination_iso_columns(source_df)

        for block_df in iter_rate_blocks(source_df):
            if block_df.empty:
                continue

            stream_id = block_stream_id(block_df)

            for _, source_row in block_df.iterrows():
                bracket = weight_bracket_label(source_row.get("Chargeable Weight"))
                if not bracket:
                    continue

                origin = cell_text(source_row.get("Origin"))
                shipment_type = cell_text(source_row.get("Shipment Type"))
                use_sfs_cost = is_sfs_tab(tab_name)
                cost_column = (
                    sfs_transport_cost_column_name(bracket)
                    if use_sfs_cost
                    else transport_cost_column_name(bracket)
                )

                for iso_column in iso_columns:
                    rate = rate_value(source_row.get(iso_column))
                    if rate is None:
                        continue

                    key = lane_key(
                        origin,
                        iso_column,
                        shipment_type,
                        stream_id,
                        tab_name=tab_name,
                    )
                    if key not in lanes:
                        lanes[key] = {
                            "Tab": tab_name,
                            "Origin country": origin,
                            "Shipment Type": shipment_type,
                            "Service": "",
                            "Service level": stream_id,
                            "Destination country": iso_column,
                            "Destination Postal Code": "",
                            "Carrier Account number": "",
                            "Business Segment": "",
                            "Category": "",
                            "Shipping Condition": "",
                            "Valid from": format_display_date(source_row.get("Version Date")),
                            "Valid to": "",
                            CURRENCY_COLUMN: cell_text(source_row.get("Billing Currency")).upper(),
                            **{column: None for column in transport_columns},
                        }

                    lane = lanes[key]
                    if not lane["Valid from"]:
                        lane["Valid from"] = format_display_date(source_row.get("Version Date"))
                    if not lane[CURRENCY_COLUMN]:
                        lane[CURRENCY_COLUMN] = cell_text(source_row.get("Billing Currency")).upper()

                    existing_rate = lane.get(cost_column)
                    if existing_rate is not None and existing_rate != rate:
                        raise ValueError(
                            f"Conflicting rates for lane {key} at {bracket}: {existing_rate} vs {rate}"
                        )
                    lane[cost_column] = rate

    if not lanes:
        return pd.DataFrame(columns=[*SHIPMENT_COLUMNS, CURRENCY_COLUMN, *transport_columns])

    result = pd.DataFrame(list(lanes.values()))
    result["Lane #"] = range(1, len(result) + 1)
    columns = [*SHIPMENT_COLUMNS, CURRENCY_COLUMN, *transport_columns]
    return result[columns]


def _style_header_cell(
    cell,
    *,
    bold: bool = True,
    center: bool = False,
    fill: PatternFill = HEADER_FILL,
) -> None:
    cell.font = BOLD if bold else NORMAL
    cell.fill = fill
    cell.alignment = CENTER if center else LEFT
    cell.border = THIN_BORDER


def _style_cost_meta_cell(cell, *, fill: PatternFill = COST_META_FILL) -> None:
    cell.font = NORMAL
    cell.fill = fill
    cell.alignment = LEFT
    cell.border = THIN_BORDER


def _write_merged_header_cell(
    worksheet,
    row_index: int,
    start_col: int,
    end_col: int,
    value: str,
    *,
    fill: PatternFill,
    bold: bool = True,
    center: bool = True,
) -> None:
    if end_col > start_col:
        worksheet.merge_cells(
            start_row=row_index,
            start_column=start_col,
            end_row=row_index,
            end_column=end_col,
        )
    cell = worksheet.cell(row_index, start_col, value)
    _style_header_cell(cell, bold=bold, center=center, fill=fill)


def _write_merged_meta_cell(
    worksheet,
    row_index: int,
    start_col: int,
    end_col: int,
    value: str,
    *,
    fill: PatternFill = COST_META_FILL,
) -> None:
    if end_col > start_col:
        worksheet.merge_cells(
            start_row=row_index,
            start_column=start_col,
            end_row=row_index,
            end_column=end_col,
        )
    cell = worksheet.cell(row_index, start_col, value)
    _style_cost_meta_cell(cell, fill=fill)


def _column_width_for_header(header: str) -> float:
    return min(42.0, max(14.0, len(header) * 0.9 + 4))


def _write_transport_cost_group(
    worksheet,
    *,
    start_col: int,
    end_col: int,
    group_title: str,
    cost_columns: list[str],
    column_prefix: str,
    bracket_rate_by: dict[str, str],
    include_currency: bool = False,
) -> None:
    if not cost_columns:
        return

    _write_merged_header_cell(
        worksheet,
        COST_GROUP_ROW,
        start_col,
        end_col,
        group_title,
        fill=TRANSPORT_GROUP_FILL,
    )

    if include_currency:
        currency_name_cell = worksheet.cell(COST_NAME_ROW, start_col, CURRENCY_COLUMN)
        _style_header_cell(currency_name_cell, fill=TRANSPORT_COST_FILL, center=True)

    data_start_col = start_col + (1 if include_currency else 0)
    for offset, cost_column in enumerate(cost_columns, start=data_start_col):
        bracket_label = cost_column.removeprefix(column_prefix).removesuffix(")")
        cell = worksheet.cell(COST_NAME_ROW, offset, bracket_label)
        _style_header_cell(cell, fill=TRANSPORT_COST_FILL, center=True)

    _write_merged_meta_cell(
        worksheet,
        APPLY_IF_ROW,
        start_col,
        end_col,
        "Applies if invoiced by Carrier",
    )

    if include_currency:
        currency_rate_by_cell = worksheet.cell(RATE_BY_ROW, start_col)
        _style_cost_meta_cell(currency_rate_by_cell)

    for offset, cost_column in enumerate(cost_columns, start=data_start_col):
        bracket_label = cost_column.removeprefix(column_prefix).removesuffix(")")
        rate_by = bracket_rate_by.get(bracket_label, "")
        cell = worksheet.cell(
            RATE_BY_ROW,
            offset,
            f"Rate by: {rate_by}\nRegular rule" if rate_by else "",
        )
        _style_cost_meta_cell(cell)


def write_matrix_sheet(
    workbook: Workbook,
    matrix_df: pd.DataFrame,
    *,
    bracket_rate_by: dict[str, str],
    sfs_bracket_rate_by: dict[str, str] | None = None,
    sheet_name: str = "Rate card",
) -> None:
    worksheet = workbook.active
    worksheet.title = sheet_name
    sfs_bracket_rate_by = sfs_bracket_rate_by or {}

    standard_columns = [
        column for column in matrix_df.columns if is_standard_weight_bracket_cost_column(column)
    ]
    sfs_columns = [column for column in matrix_df.columns if is_sfs_weight_bracket_cost_column(column)]
    shipment_count = len(SHIPMENT_COLUMNS)
    currency_col = shipment_count + 1
    standard_start_col = shipment_count + 2
    standard_end_col = shipment_count + 1 + len(standard_columns)
    sfs_start_col = standard_end_col + 1
    sfs_end_col = standard_end_col + len(sfs_columns)

    if standard_columns:
        _write_transport_cost_group(
            worksheet,
            start_col=currency_col,
            end_col=standard_end_col,
            group_title=f"Grouped cost: {TRANSPORT_COST_GROUP}",
            cost_columns=standard_columns,
            column_prefix=f"{TRANSPORT_COST_GROUP} (",
            bracket_rate_by=bracket_rate_by,
            include_currency=True,
        )

    if sfs_columns:
        _write_transport_cost_group(
            worksheet,
            start_col=sfs_start_col,
            end_col=sfs_end_col,
            group_title=f"Grouped cost: {SFS_TRANSPORT_COST_GROUP}",
            cost_columns=sfs_columns,
            column_prefix=f"{SFS_TRANSPORT_COST_GROUP} (",
            bracket_rate_by=sfs_bracket_rate_by,
            include_currency=False,
        )

    for col_index, header in enumerate(SHIPMENT_COLUMNS, start=1):
        cell = worksheet.cell(COLUMN_HEADER_ROW, col_index, header)
        _style_header_cell(cell, bold=header == "Lane #")

    if standard_columns:
        currency_header = worksheet.cell(COLUMN_HEADER_ROW, currency_col, CURRENCY_COLUMN)
        _style_header_cell(currency_header, center=True, fill=TRANSPORT_COST_FILL)

        for offset, cost_column in enumerate(standard_columns, start=standard_start_col):
            bracket_label = cost_column.removeprefix(f"{TRANSPORT_COST_GROUP} (").removesuffix(")")
            rate_by = bracket_rate_by.get(bracket_label, "Rate")
            header_cell = worksheet.cell(COLUMN_HEADER_ROW, offset, rate_by)
            _style_header_cell(header_cell, center=True, fill=TRANSPORT_COST_FILL)

    if sfs_columns:
        for offset, cost_column in enumerate(sfs_columns, start=sfs_start_col):
            bracket_label = cost_column.removeprefix(f"{SFS_TRANSPORT_COST_GROUP} (").removesuffix(")")
            rate_by = sfs_bracket_rate_by.get(bracket_label, "Rate")
            header_cell = worksheet.cell(COLUMN_HEADER_ROW, offset, rate_by)
            _style_header_cell(header_cell, center=True, fill=TRANSPORT_COST_FILL)

    all_cost_columns = [*standard_columns, *sfs_columns]
    for matrix_index, (_, row) in enumerate(matrix_df.iterrows()):
        excel_row = DATA_START_ROW + matrix_index

        for col_index, header in enumerate(SHIPMENT_COLUMNS, start=1):
            cell = worksheet.cell(excel_row, col_index, row.get(header))
            cell.alignment = LEFT
            cell.border = THIN_BORDER

        if standard_columns or sfs_columns:
            currency_cell = worksheet.cell(excel_row, currency_col, row.get(CURRENCY_COLUMN))
            currency_cell.alignment = CENTER
            currency_cell.border = THIN_BORDER

        for cost_column in all_cost_columns:
            offset = (
                currency_col + 1 + standard_columns.index(cost_column)
                if cost_column in standard_columns
                else sfs_start_col + sfs_columns.index(cost_column)
            )
            value = row.get(cost_column)
            cell = worksheet.cell(excel_row, offset)
            cell.border = THIN_BORDER
            if value is not None and value != "":
                cell.value = value
                cell.number_format = RATE_NUMBER_FORMAT
                cell.alignment = CENTER

    for col_index, header in enumerate(SHIPMENT_COLUMNS, start=1):
        worksheet.column_dimensions[get_column_letter(col_index)].width = _column_width_for_header(header)

    if standard_columns or sfs_columns:
        worksheet.column_dimensions[get_column_letter(currency_col)].width = 12.0

    for offset, cost_column in enumerate(standard_columns, start=standard_start_col):
        worksheet.column_dimensions[get_column_letter(offset)].width = _column_width_for_header(cost_column)

    for offset, cost_column in enumerate(sfs_columns, start=sfs_start_col):
        worksheet.column_dimensions[get_column_letter(offset)].width = _column_width_for_header(cost_column)

    worksheet.freeze_panes = worksheet.cell(DATA_START_ROW, 1)
    worksheet.sheet_view.showGridLines = False


def save_matrix(
    matrix_df: pd.DataFrame,
    source_file: Path,
    *,
    bracket_rate_by: dict[str, str],
    sfs_bracket_rate_by: dict[str, str] | None = None,
    output_path: Path | None = None,
) -> Path:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    if output_path is None:
        output_path = OUTPUT_DIR / f"{source_file.stem.replace('_extracted', '')}_matrix.xlsx"

    workbook = Workbook()
    write_matrix_sheet(
        workbook,
        matrix_df,
        bracket_rate_by=bracket_rate_by,
        sfs_bracket_rate_by=sfs_bracket_rate_by,
    )
    workbook.save(output_path)
    return output_path


def run_build_matrix(
    *,
    source_file: Path | None = None,
    output_path: Path | None = None,
    auto: bool = False,
) -> Path:
    ensure_workspace_dirs()
    files = list_extracted_files()
    file_path = source_file
    if file_path is None:
        file_path = select_extracted_file(files, auto=auto)

    rate_tabs = load_rate_tabs(file_path)
    if not rate_tabs:
        raise RuntimeError(
            f"No active rate tabs found in {file_path.name}. "
            "Expected ACTIVE tabs from Tab Index with Origin, Chargeable Weight, "
            f"and destination ISO columns (excluding {', '.join(sorted(EXCLUDED_RATE_TABS))})."
        )

    print(f"\nBuilding matrix from {file_path.name}:")
    for tab_name, tab_df in rate_tabs:
        print(f"  - {tab_name}: {len(tab_df)} source rows")

    _, _, bracket_rate_by, sfs_bracket_rate_by = collect_transport_cost_columns(rate_tabs)
    matrix_df = build_shipment_and_cost_rows(rate_tabs)
    tab_index_df = load_tab_index(file_path)
    billing_bs_lookup = load_billing_bs_lookup()
    matrix_df = enrich_matrix_with_tab_index(
        matrix_df,
        tab_index_df,
        billing_bs_lookup=billing_bs_lookup,
    )

    return_tab_names = [
        tab_name
        for tab_name, tab_info in build_tab_index_lookup(tab_index_df).items()
        if tab_info.has_return_lanes
    ]
    if return_tab_names:
        print(f"  Return lanes added for tabs: {', '.join(sorted(return_tab_names))}")
    if billing_bs_lookup:
        billing_bs_path = billing_bs_file_path()
        print(
            f"  Business Segment lookup: {len(billing_bs_lookup)} accounts "
            f"from {billing_bs_path.name if billing_bs_path else 'billing-bs'}"
        )
    if sfs_bracket_rate_by:
        print(
            f"  SFS transport cost brackets: "
            f"{', '.join(sfs_bracket_rate_by.keys())}"
        )

    saved_path = save_matrix(
        matrix_df,
        file_path,
        bracket_rate_by=bracket_rate_by,
        sfs_bracket_rate_by=sfs_bracket_rate_by,
        output_path=output_path,
    )
    print(f"\nSaved matrix ({len(matrix_df)} lanes) to: {saved_path}")
    print(f"  Weight brackets: {', '.join(bracket_rate_by.keys())}")
    return saved_path


def main() -> int:
    try:
        run_build_matrix()
        return 0
    except KeyboardInterrupt:
        print("\nOperation cancelled.")
        return 130
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
