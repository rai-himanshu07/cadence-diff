"""QC Tool entry point: `conda run -n py311 python main.py`."""

from pathlib import Path

from qc_tool.ui.app import run_app

if __name__ in {"__main__", "__mp_main__"}:  # __mp_main__: NiceGUI multiprocessing
    run_app(Path(__file__).parent / "data")
