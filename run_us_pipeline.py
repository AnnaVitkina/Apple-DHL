"""
DHL US rates pipeline entry point.

Delegates to run_pipeline.py with US mode enabled.

Usage:
  python run_us_pipeline.py
  python run_us_pipeline.py --auto
  python run_us_pipeline.py --input "input/DHL US Rates v40 20260326 (1).xlsx"
"""

from __future__ import annotations

import sys
from pathlib import Path

_CODE_DIR = Path(__file__).resolve().parent
if str(_CODE_DIR) not in sys.path:
    sys.path.insert(0, str(_CODE_DIR))

from run_pipeline import main, run_pipeline, _env_flag, _running_in_notebook


def main_us() -> int:
    if _running_in_notebook():
        run_pipeline(auto=_env_flag("DHL_AUTO") or _env_flag("DHL_US_AUTO"), us=True)
        return 0
    return main()


if __name__ == "__main__":
    if _running_in_notebook():
        run_pipeline(auto=_env_flag("DHL_AUTO") or _env_flag("DHL_US_AUTO"), us=True)
    else:
        # Re-parse with US mode forced by injecting --us when launched directly.
        if "--us" not in sys.argv:
            sys.argv.append("--us")
        raise SystemExit(main())
