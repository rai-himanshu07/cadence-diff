"""`python -m qc_tool` entry point."""

from qc_tool.cli import main

if __name__ in {"__main__", "__mp_main__"}:  # __mp_main__: NiceGUI multiprocessing
    raise SystemExit(main())
