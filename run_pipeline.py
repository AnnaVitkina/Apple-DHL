"""
FSC — end-to-end rate pipeline.

Steps:
  1. Select RA and FSC input files
  2. Convert Rate card / FSC tabs to processing workbooks
  3. Match fuel surcharge values by origin and destination city
  4. Save formatted Rate card workbook to output/

Usage (local):
  python run_pipeline.py
  python run_pipeline.py --auto
  python run_pipeline.py --ra-file "input/RA/file.xlsx" --fsc-file "input/fsc/file.xlsx"

Usage (Google Colab):
  from google.colab import drive
  drive.mount("/content/drive")

  import sys
  sys.path.insert(0, "/content/Apple-FSC")

  from run_pipeline import run_pipeline
  run_pipeline()

  # Optional: skip file-selection prompts
  # import os
  # os.environ["FSC_AUTO"] = "1"
  # run_pipeline(auto=True)

  # Optional: override Drive data folder
  # os.environ["FSC_DRIVE_BASE"] = (
  #     "/content/drive/Shareddrives/FA Ops Europe: Rate Maintenance Team "
  #     "/Documents/AI Adoption RMT/RMT_Apple/RMT_FSC"
  # )
"""

from __future__ import annotations

import argparse
import os
import sys
from dataclasses import dataclass
from pathlib import Path

_CODE_DIR = Path(os.environ.get("FSC_ROOT", "/content/Apple-FSC")).resolve()
try:
    _CODE_DIR = Path(__file__).resolve().parent
except NameError:
    pass
if str(_CODE_DIR) not in sys.path:
    sys.path.insert(0, str(_CODE_DIR))

_PIPELINE_MODULES = (
    "project_paths",
    "excel_formatting",
    "convert_to_processing",
)


def _bootstrap_paths():
    for module_name in _PIPELINE_MODULES:
        sys.modules.pop(module_name, None)

    import project_paths

    project_paths.configure_paths_from_env()
    return project_paths


_project_paths = _bootstrap_paths()
configure_paths_from_env = _project_paths.configure_paths_from_env
print_path_config = _project_paths.print_path_config

from convert_to_processing import ConvertResult, run_convert


@dataclass(frozen=True)
class PipelineResult:
    convert: ConvertResult


def run_pipeline(
    *,
    auto: bool = False,
    ra_file: Path | None = None,
    fsc_file: Path | None = None,
    fsc_sheet: str | None = None,
    output_path: Path | None = None,
) -> PipelineResult:
    configure_paths_from_env()
    print_path_config()

    print("\n=== Step 1/3: Load RA and FSC input files ===")
    convert_result = run_convert(
        auto=auto,
        ra_file=ra_file,
        fsc_file=fsc_file,
        fsc_sheet=fsc_sheet,
        output_path=output_path,
    )

    print("\n=== Step 2/3: Processing workbooks saved ===")
    print(f"  RA processing:  {convert_result.ra_processing_path}")
    print(f"  FSC processing: {convert_result.fsc_processing_path}")

    print("\n=== Step 3/3: Formatted output saved ===")
    print(f"  Output workbook: {convert_result.output_path}")
    print(
        f"  Matched lanes:   {convert_result.matched_lane_count} / "
        f"{convert_result.total_lane_count}"
    )

    print("\n=== Pipeline complete ===")
    return PipelineResult(convert=convert_result)


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run the FSC end-to-end rate pipeline.")
    parser.add_argument(
        "--auto",
        action="store_true",
        help="Use the first available RA/FSC files and skip interactive prompts.",
    )
    parser.add_argument(
        "--ra-file",
        type=Path,
        default=None,
        help="Optional RA input workbook path.",
    )
    parser.add_argument(
        "--fsc-file",
        type=Path,
        default=None,
        help="Optional FSC input workbook path.",
    )
    parser.add_argument(
        "--fsc-sheet",
        default=None,
        help="Optional FSC sheet name.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Optional formatted output workbook path.",
    )
    return parser.parse_args()


def main() -> int:
    try:
        args = _parse_args()
        run_pipeline(
            auto=args.auto,
            ra_file=args.ra_file,
            fsc_file=args.fsc_file,
            fsc_sheet=args.fsc_sheet,
            output_path=args.output,
        )
        return 0
    except KeyboardInterrupt:
        print("\nOperation cancelled.")
        return 130
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


def _env_flag(name: str) -> bool:
    return os.environ.get(name, "").strip().lower() in {"1", "true", "yes", "on"}


def _running_in_notebook() -> bool:
    if "colab_kernel_launcher" in Path(sys.argv[0]).name:
        return True
    if any(arg == "-f" for arg in sys.argv):
        return True
    return "ipykernel" in sys.modules or "IPython" in sys.modules


if __name__ == "__main__":
    if _running_in_notebook():
        run_pipeline(
            auto=_env_flag("FSC_AUTO"),
            ra_file=Path(os.environ["FSC_RA_FILE"]) if os.environ.get("FSC_RA_FILE") else None,
            fsc_file=Path(os.environ["FSC_FSC_FILE"]) if os.environ.get("FSC_FSC_FILE") else None,
            fsc_sheet=os.environ.get("FSC_FSC_SHEET"),
            output_path=Path(os.environ["FSC_OUTPUT"]) if os.environ.get("FSC_OUTPUT") else None,
        )
    else:
        raise SystemExit(main())
