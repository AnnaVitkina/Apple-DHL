"""
DHL — end-to-end rate pipeline.

Steps (EU/ME/GB):
  1. Select input file and tabs, convert to processing/
  2. Build matrix workbook in output/

Steps (US):
  1. Select DHL US input workbook from input/
  2. Build US matrix workbook in output/

Usage (local):
  python run_pipeline.py
  python run_pipeline.py --auto
  python run_pipeline.py --us
  python run_pipeline.py --us --auto
  python run_pipeline.py --convert-only
  python run_pipeline.py --matrix-only
  python run_pipeline.py --input "input/DHL US Rates v40 20260326 (1).xlsx"

Usage (Google Colab):
  from google.colab import drive
  drive.mount("/content/drive")

  import os
  os.environ["DHL_AUTO"] = "1"          # optional: default tabs, no prompts
  os.environ["DHL_US"] = "1"             # optional: run US flow
  # os.environ["DHL_MATRIX_ONLY"] = "1" # optional: skip conversion step

  exec(open("/content/Apple-DHL/run_pipeline.py", encoding="utf-8").read())

Code lives in /content/Apple-DHL.
Input/processing/output are read from the shared Drive RMT_DHL folder when mounted.
"""

from __future__ import annotations

import argparse
import os
import sys
from dataclasses import dataclass
from pathlib import Path

_CODE_DIR = Path(os.environ.get("DHL_CODE_DIR", "/content/Apple-DHL")).resolve()
try:
    _CODE_DIR = Path(__file__).resolve().parent
except NameError:
    pass
if str(_CODE_DIR) not in sys.path:
    sys.path.insert(0, str(_CODE_DIR))

_PIPELINE_MODULES = (
    "project_paths",
    "build_matrix",
    "build_postal_code_zones",
    "convert_to_processing",
    "us_config",
    "us_excel",
    "build_us_matrix",
    "us_zone_countries",
)


def _bootstrap_paths() -> None:
    for module_name in _PIPELINE_MODULES:
        sys.modules.pop(module_name, None)

    import project_paths

    project_paths.configure_paths_from_env()
    return project_paths


_project_paths = _bootstrap_paths()
configure_paths_from_env = _project_paths.configure_paths_from_env
print_path_config = _project_paths.print_path_config

from build_matrix import run_build_matrix
from build_us_matrix import (
    is_us_rate_file,
    list_us_input_files,
    run_build_us_matrix,
    select_us_input_file,
)
from convert_to_processing import list_input_files, run_convert, select_input_file


@dataclass(frozen=True)
class PipelineResult:
    input_path: Path | None
    processing_path: Path | None
    output_path: Path | None


def _resolve_input_file(
    *,
    auto: bool,
    input_path: Path | None,
    us: bool,
) -> Path:
    if input_path is not None:
        return input_path
    if us:
        return select_us_input_file(list_us_input_files(), auto=auto)
    return select_input_file(list_input_files(), auto=auto)


def _should_run_us_flow(
    *,
    us: bool,
    input_path: Path | None,
) -> bool:
    if us or _env_flag("DHL_US"):
        return True
    return input_path is not None and is_us_rate_file(input_path)


def run_pipeline(
    *,
    auto: bool = False,
    us: bool = False,
    convert_only: bool = False,
    matrix_only: bool = False,
    input_path: Path | None = None,
    processing_path: Path | None = None,
    output_path: Path | None = None,
) -> PipelineResult:
    if convert_only and matrix_only:
        raise ValueError("Use only one of --convert-only or --matrix-only.")
    if convert_only and (us or _should_run_us_flow(us=us, input_path=input_path)):
        raise ValueError("US rates do not use the convert step. Run without --convert-only.")

    configure_paths_from_env()
    print_path_config()

    saved_input_path = input_path
    saved_processing_path = processing_path
    saved_output_path: Path | None = None
    run_us = _should_run_us_flow(us=us, input_path=input_path)

    if run_us:
        step_label = "Step 1/1" if matrix_only else "Step 1/1: Build US matrix"
        print(f"\n=== {step_label} ===")
        source_file = _resolve_input_file(auto=auto, input_path=input_path, us=True)
        saved_input_path = source_file
        saved_output_path = run_build_us_matrix(
            source_file=source_file,
            output_path=output_path,
            auto=auto,
        )
        print("\n=== Pipeline complete ===")
        print(f"  Input workbook:  {saved_input_path}")
        print(f"  Output workbook: {saved_output_path}")
        return PipelineResult(
            input_path=saved_input_path,
            processing_path=None,
            output_path=saved_output_path,
        )

    if not matrix_only:
        print("\n=== Step 1/2: Convert input to processing ===")
        if saved_input_path is None:
            saved_input_path = _resolve_input_file(auto=auto, input_path=None, us=False)
        if is_us_rate_file(saved_input_path):
            print(f"\nDetected US rate file: {saved_input_path.name}")
            print("Re-running in US mode (no conversion step).")
            saved_output_path = run_build_us_matrix(
                source_file=saved_input_path,
                output_path=output_path,
                auto=auto,
            )
            print("\n=== Pipeline complete ===")
            print(f"  Input workbook:      {saved_input_path}")
            print(f"  Output workbook:     {saved_output_path}")
            return PipelineResult(
                input_path=saved_input_path,
                processing_path=None,
                output_path=saved_output_path,
            )

        saved_processing_path = run_convert(auto=auto, file_path=saved_input_path)

    if not convert_only:
        step_label = "Step 2/2" if not matrix_only else "Step 1/1"
        print(f"\n=== {step_label}: Build matrix ===")
        saved_output_path = run_build_matrix(
            source_file=saved_processing_path,
            output_path=output_path,
            auto=auto or matrix_only,
        )

    print("\n=== Pipeline complete ===")
    if saved_input_path is not None:
        print(f"  Input workbook:      {saved_input_path}")
    if saved_processing_path is not None:
        print(f"  Processing workbook: {saved_processing_path}")
    if saved_output_path is not None:
        print(f"  Output workbook:     {saved_output_path}")

    return PipelineResult(
        input_path=saved_input_path,
        processing_path=saved_processing_path,
        output_path=saved_output_path,
    )


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run the DHL end-to-end rate pipeline.")
    parser.add_argument(
        "--auto",
        action="store_true",
        help="Use default tabs and skip interactive prompts where possible.",
    )
    parser.add_argument(
        "--us",
        action="store_true",
        help="Run the DHL US rates flow (reads input directly, writes to output/).",
    )
    parser.add_argument(
        "--convert-only",
        action="store_true",
        help="Only run input -> processing conversion.",
    )
    parser.add_argument(
        "--matrix-only",
        action="store_true",
        help="Only build matrix from the latest extracted processing file.",
    )
    parser.add_argument(
        "--input",
        type=Path,
        default=None,
        help="Optional input workbook. US files are auto-detected.",
    )
    parser.add_argument(
        "--processing",
        type=Path,
        default=None,
        help="Optional extracted processing workbook for --matrix-only.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Optional output matrix workbook path.",
    )
    return parser.parse_args()


def main() -> int:
    try:
        args = _parse_args()
        run_pipeline(
            auto=args.auto,
            us=args.us,
            convert_only=args.convert_only,
            matrix_only=args.matrix_only,
            input_path=args.input,
            processing_path=args.processing,
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
            auto=_env_flag("DHL_AUTO"),
            us=_env_flag("DHL_US"),
            convert_only=_env_flag("DHL_CONVERT_ONLY"),
            matrix_only=_env_flag("DHL_MATRIX_ONLY"),
        )
    else:
        raise SystemExit(main())
