"""Helper entry point to launch the Timesheet timer app without installation.

The project uses a ``src`` layout, so the ``timesheet_app`` package is located
under the ``src`` directory.  Running ``python -m timesheet_app`` from the
repository root will therefore fail unless that directory is added to
``PYTHONPATH``.  This lightweight wrapper ensures the path is configured and
then delegates to the real application entry point.
"""

from __future__ import annotations

from pathlib import Path
import sys


def _ensure_src_on_path() -> None:
    project_root = Path(__file__).resolve().parent
    src_dir = project_root / "src"
    src_dir_str = str(src_dir)
    if src_dir.is_dir() and src_dir_str not in sys.path:
        # Prepend so the local sources are always preferred over any installed
        # package with the same name.
        sys.path.insert(0, src_dir_str)


def main() -> None:
    _ensure_src_on_path()

    from timesheet_app.app import main as app_main

    app_main()


if __name__ == "__main__":
    main()
