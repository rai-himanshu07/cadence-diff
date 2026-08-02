"""`python -m qc_tool` entry point."""

from qc_tool.cli import main

# Only `__main__`: a spawned QC worker re-imports this file as `__mp_main__`.
if __name__ == "__main__":
    raise SystemExit(main())
