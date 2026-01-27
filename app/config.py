import json
import os
from . import resources

DEFAULT_SETTINGS_FILENAME = "default_settings.json"
SETTINGS_FILENAME = "_mini_fish_settings.json"


def default_settings_path() -> str:
    return os.path.join(resources.app_dir(), DEFAULT_SETTINGS_FILENAME)


def settings_path() -> str:
    return os.path.join(resources.app_dir(), SETTINGS_FILENAME)


def load_settings(fallback: dict | None = None) -> dict:
    data = dict(fallback or {})
    path = settings_path()
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                data.update(json.load(f))
        except Exception:
            pass
    return data


def save_settings(data: dict) -> bool:
    try:
        with open(settings_path(), "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False
