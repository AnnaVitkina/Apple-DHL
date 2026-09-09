"""Configuration for the DHL US rates pipeline."""

from __future__ import annotations

from dataclasses import dataclass
from enum import Enum

US_FILE_PATTERN = r"DHL US Rates"
US_EXTRACTED_SUFFIX = "_us_extracted"
INDEX_SHEET = "Index"
DHLC_DESTINATION_ZONES_SHEET = "DHLC - Destination Zones"
DHL_3C_ZONES_SHEET = "DHL 3C - Zones"
DHL_ORIGIN_ZONES_SHEET = "DHL - Origin Zones"
DHL_IB_DESTINATION_REGION = "Zone IB"

CURRENCY_COLUMN = "Currency"

US_SHIPMENT_COLUMNS = (
    "Lane #",
    "Tab",
    "Origin Country Region",
    "Destination Country Region",
    "Service Type/Package Type",
    "Zone 3C",
    "Carrier Account Number",
)

class USCostGroup(str, Enum):
    DHLG = "Transport cost (DHLG)"
    DHLC_ENVELOPE = "Transport cost (DHLC - Envelope)"
    DHLC_DOCS = "Transport cost (DHLC - Documents and Non-documents)"
    DHL_3C = "Transport cost (DHL 3C)"
    DHL_IB = "Transport cost (DHL IB)"


US_COST_GROUP_ORDER: tuple[USCostGroup, ...] = (
    USCostGroup.DHLG,
    USCostGroup.DHLC_ENVELOPE,
    USCostGroup.DHLC_DOCS,
    USCostGroup.DHL_3C,
    USCostGroup.DHL_IB,
)


def us_cost_group_label(cost_group: USCostGroup) -> str:
    prefix = "Transport cost ("
    value = cost_group.value
    if value.startswith(prefix) and value.endswith(")"):
        return value[len(prefix) : -1]
    return value


@dataclass(frozen=True)
class USTabConfig:
    sheet_name: str
    cost_groups: tuple[USCostGroup, ...]
    processor: str


DHLG_TABS: tuple[USTabConfig, ...] = (
    USTabConfig(
        "DHLG - US_TD_BBX   WWX",
        (USCostGroup.DHLG,),
        "dhlg_bbx",
    ),
    USTabConfig(
        "DHLG - US_TD_BBX   AOS",
        (USCostGroup.DHLG,),
        "dhlg_bbx_aos",
    ),
)

DHLC_TABS: tuple[USTabConfig, ...] = (
    USTabConfig("DHLC - US_TD_Exp WW", (USCostGroup.DHLC_ENVELOPE, USCostGroup.DHLC_DOCS), "dhlc_exp"),
    USTabConfig("DHLC - US_TD_Exp 9am", (USCostGroup.DHLC_DOCS,), "dhlc_exp"),
    USTabConfig("DHLC - US_TD_Exp 12pm", (USCostGroup.DHLC_DOCS,), "dhlc_exp"),
)

DHL_3C_TABS: tuple[USTabConfig, ...] = (
    USTabConfig("DHL 3C_US_TD_Exp WW", (USCostGroup.DHL_3C,), "dhl_3c"),
    USTabConfig("DHL 3C_US_TD_Exp WW 9am", (USCostGroup.DHL_3C,), "dhl_3c"),
    USTabConfig("DHL 3C_US_TD_Exp WW 1030", (USCostGroup.DHL_3C,), "dhl_3c"),
    USTabConfig("DHL 3C_US_TD_Exp WW 12pm", (USCostGroup.DHL_3C,), "dhl_3c"),
)

DHL_IB_TABS: tuple[USTabConfig, ...] = (
    USTabConfig("DHL IB_US_TD_Exp", (USCostGroup.DHL_IB,), "dhl_ib"),
    USTabConfig("DHL IB_US_TD_Exp 1030am", (USCostGroup.DHL_IB,), "dhl_ib"),
    USTabConfig("DHL IB_US_TD_Exp 12pm", (USCostGroup.DHL_IB,), "dhl_ib"),
)

US_RATE_TAB_CONFIGS: tuple[USTabConfig, ...] = (
    *DHLG_TABS,
    *DHLC_TABS,
    *DHL_3C_TABS,
    *DHL_IB_TABS,
)

US_SUPPORT_SHEETS = (
    INDEX_SHEET,
    DHLC_DESTINATION_ZONES_SHEET,
    DHL_3C_ZONES_SHEET,
    DHL_ORIGIN_ZONES_SHEET,
)

FRED_SERVICE_BY_RATE_CARD: dict[tuple[str, str], str] = {
    ("TIME DEFINITE US-TD-BBX", ""): "EXP_BREAKBULK",
    ("PB", ""): "EXP_ENVELOPE",
    ("IE", "PKG89"): "EXP_WW_DOC",
    ("CX", "PKG89"): "EXP_WW_DOC",
    ("IE", "PKG90"): "EXP_WW_NONDOC",
    ("CX", "PKG90"): "EXP_WW_NONDOC",
    ("9A", "PKG89"): "EXP_900_DOC",
    ("9A", "PKG90"): "EXP_900_NONDOC",
    ("PN", "PKG89"): "EXP_1200_DOC",
    ("PN", "PKG90"): "EXP_1200_NONDOC",
    ("PA", "PKG89"): "EXP_1030_DOC",
    ("PA", "PKG90"): "EXP_1030_NONDOC",
}
