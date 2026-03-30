import os
import sys
from pathlib import Path


def get_app_dir() -> str:
    if getattr(sys, "frozen", False):
        return str(Path(sys.executable).resolve().parent)
    return str(Path(__file__).resolve().parent)


def get_resource_path(*parts: str) -> str:
    return os.path.join(get_app_dir(), *parts)
