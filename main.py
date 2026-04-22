"""Entry point: initialize logging and launch the tkinter GUI."""

import os
import sys


def _ensure_project_root_on_path() -> None:
    here = os.path.dirname(os.path.abspath(__file__))
    if here not in sys.path:
        sys.path.insert(0, here)


def main() -> None:
    _ensure_project_root_on_path()

    from utils.logger import init_logger
    from ui.main_window import launch

    init_logger()
    launch()


if __name__ == "__main__":
    main()
