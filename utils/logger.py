"""Unified logger wrapper for CLI and GUI callers."""

import logging
import sys

_INITIALIZED = False


def init_logger(level: int = logging.INFO) -> None:
    global _INITIALIZED
    if _INITIALIZED:
        return

    root = logging.getLogger()
    root.setLevel(level)

    handler = logging.StreamHandler(sys.stdout)
    handler.setLevel(level)
    formatter = logging.Formatter(
        fmt="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    handler.setFormatter(formatter)
    root.addHandler(handler)
    _INITIALIZED = True


def get_logger(name: str) -> logging.Logger:
    if not _INITIALIZED:
        init_logger()
    return logging.getLogger(name)
