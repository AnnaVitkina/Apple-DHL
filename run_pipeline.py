"""
DHL — end-to-end rate pipeline.

Steps:
  1. Select input file and tabs, convert to processing/
  2. Build matrix workbook in output/

Usage (local):
  python run_pipeline.py
  python run_pipeline.py --auto
  python run_pipeline.py --convert-only
  python run_pipeline.py --matrix-only

Usage (Google Colab):
  from google.colab import drive
  drive.mount("/content/drive")

  import os
  os.environ["DHL_AUTO"] = "1"          # optional: default tabs, no prompts
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
from convert_to_processing import run_convert


@dataclass(frozen=True)
class PipelineResult:
    processing_path: Path | None
    output_path: Path | None


def run_pipeline(
    *,
    auto: bool = False,
    convert_only: bool = False,
    matrix_only: bool = False,
    processing_path: Path | None = None,
    output_path: Path | None = None,
) -> PipelineResult:
    if convert_only and matrix_only:
        raise ValueError("Use only one of --convert-only or --matrix-only.")

    configure_paths_from_env()
    print_path_config()
    saved_processing_path = processing_path
    saved_output_path: Path | None = None

    if not matrix_only:
        print("\n=== Step 1/2: Convert input to processing ===")
        saved_processing_path = run_convert(auto=auto)

    if not convert_only:
        step_label = "Step 2/2" if not matrix_only else "Step 1/1"
        print(f"\n=== {step_label}: Build matrix ===")
        saved_output_path = run_build_matrix(
            source_file=saved_processing_path,
            output_path=output_path,
            auto=auto or matrix_only,
        )

    print("\n=== Pipeline complete ===")
    if saved_processing_path is not None:
        print(f"  Processing workbook: {saved_processing_path}")
    if saved_output_path is not None:
        print(f"  Output workbook:     {saved_output_path}")

    return PipelineResult(
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
            convert_only=args.convert_only,
            matrix_only=args.matrix_only,
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
            convert_only=_env_flag("DHL_CONVERT_ONLY"),
            matrix_only=_env_flag("DHL_MATRIX_ONLY"),
        )
    else:
        raise SystemExit(main())
