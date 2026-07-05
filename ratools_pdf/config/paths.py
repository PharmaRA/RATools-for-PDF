import os
import sys
from pathlib import Path


def _source_app_dir() -> str:
    return str(Path(__file__).resolve().parents[2])


def get_app_dir() -> str:
    if getattr(sys, "frozen", False):
        return str(Path(sys.executable).resolve().parent)
    return _source_app_dir()


def get_resource_dir() -> str:
    if getattr(sys, "frozen", False):
        meipass_dir = getattr(sys, "_MEIPASS", None)
        if meipass_dir:
            return str(Path(meipass_dir).resolve())
        return get_app_dir()
    return _source_app_dir()


def get_resource_path(*parts: str) -> str:
    return os.path.join(get_resource_dir(), *parts)
