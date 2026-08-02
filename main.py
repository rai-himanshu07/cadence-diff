"""QC Tool entry point: `conda run -n py311 python main.py`."""

from pathlib import Path

from qc_tool.ui.app import run_app

# Only `__main__`: a spawned QC worker re-imports this file as `__mp_main__`.
if __name__ == "__main__":
    run_app(Path(__file__).parent / "data")
