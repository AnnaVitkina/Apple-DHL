"""Project folder paths for FSC rate conversion."""

from __future__ import annotations

import os
from pathlib import Path

_DEFAULT_ROOT = Path(__file__).resolve().parent

ROOT = _DEFAULT_ROOT
INPUT_RA_DIR = ROOT / "input" / "RA"
INPUT_FSC_DIR = ROOT / "input" / "fsc"
PROCESSING_DIR = ROOT / "processing"
OUTPUT_DIR = ROOT / "output"

# Default Google Drive location for RMT FSC data in Colab.
COLAB_DRIVE_BASE = Path(
    "/content/drive/Shareddrives/FA Ops Europe: Rate Maintenance Team "
    "/Documents/AI Adoption RMT/RMT_Apple/RMT_FSC"
)


def _path_from_env(name: str) -> Path | None:
    value = os.environ.get(name)
    if not value:
        return None
    return Path(value).expanduser()


def _colab_drive_base() -> Path | None:
    candidates: list[Path] = []
    env_base = _path_from_env("FSC_DRIVE_BASE")
    if env_base is not None:
        candidates.append(env_base)
    candidates.append(COLAB_DRIVE_BASE)

    for candidate in candidates:
        if (candidate / "input" / "RA").is_dir() or (candidate / "input" / "fsc").is_dir():
            return candidate
    return None


def configure_paths(
    *,
    root: Path | str | None = None,
    input_ra_dir: Path | str | None = None,
    input_fsc_dir: Path | str | None = None,
    processing_dir: Path | str | None = None,
    output_dir: Path | str | None = None,
) -> None:
    """Override data folder locations."""
    global ROOT, INPUT_RA_DIR, INPUT_FSC_DIR, PROCESSING_DIR, OUTPUT_DIR

    if root is not None:
        ROOT = Path(root).expanduser().resolve()
    if input_ra_dir is not None:
        INPUT_RA_DIR = Path(input_ra_dir).expanduser().resolve()
    if input_fsc_dir is not None:
        INPUT_FSC_DIR = Path(input_fsc_dir).expanduser().resolve()
    if processing_dir is not None:
        PROCESSING_DIR = Path(processing_dir).expanduser().resolve()
    if output_dir is not None:
        OUTPUT_DIR = Path(output_dir).expanduser().resolve()


def configure_paths_from_env() -> None:
    """Apply FSC_* environment variables and Colab Drive defaults when available."""
    root = _path_from_env("FSC_ROOT")
    input_ra_dir = _path_from_env("FSC_INPUT_RA_DIR")
    input_fsc_dir = _path_from_env("FSC_INPUT_FSC_DIR")
    processing_dir = _path_from_env("FSC_PROCESSING_DIR")
    output_dir = _path_from_env("FSC_OUTPUT_DIR")

    if input_ra_dir is None and input_fsc_dir is None and processing_dir is None and output_dir is None:
        drive_base = _colab_drive_base()
        if drive_base is not None:
            input_ra_dir = drive_base / "input" / "RA"
            input_fsc_dir = drive_base / "input" / "fsc"
            processing_dir = drive_base / "processing"
            output_dir = drive_base / "output"

    if any(path is not None for path in (root, input_ra_dir, input_fsc_dir, processing_dir, output_dir)):
        configure_paths(
            root=root or ROOT,
            input_ra_dir=input_ra_dir or (root / "input" / "RA" if root is not None else INPUT_RA_DIR),
            input_fsc_dir=input_fsc_dir or (root / "input" / "fsc" if root is not None else INPUT_FSC_DIR),
            processing_dir=processing_dir or (root / "processing" if root is not None else PROCESSING_DIR),
            output_dir=output_dir or (root / "output" if root is not None else OUTPUT_DIR),
        )


def ensure_workspace_dirs() -> None:
    INPUT_RA_DIR.mkdir(parents=True, exist_ok=True)
    INPUT_FSC_DIR.mkdir(parents=True, exist_ok=True)
    PROCESSING_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def print_path_config() -> None:
    print("FSC paths:")
    print(f"  Root:       {ROOT}")
    print(f"  Input RA:   {INPUT_RA_DIR}")
    print(f"  Input FSC:  {INPUT_FSC_DIR}")
    print(f"  Processing: {PROCESSING_DIR}")
    print(f"  Output:     {OUTPUT_DIR}")
