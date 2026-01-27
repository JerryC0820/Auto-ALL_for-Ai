import os
import sys


def base_dir() -> str:
    return os.path.abspath(getattr(sys, "_MEIPASS", os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")))


def app_dir() -> str:
    if bool(getattr(sys, "frozen", False)):
        return os.path.dirname(os.path.abspath(sys.executable))
    return base_dir()


def resource_path(*parts: str) -> str:
    return os.path.join(base_dir(), *parts)
