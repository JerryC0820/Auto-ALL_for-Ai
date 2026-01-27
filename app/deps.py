import os
import shutil


def find_ahk_exe() -> str:
    candidates = [
        shutil.which("AutoHotkey.exe"),
        shutil.which("AutoHotkeyU64.exe"),
        shutil.which("AutoHotkeyU32.exe"),
        r"C:\\Program Files\\AutoHotkey\\AutoHotkey.exe",
        r"C:\\Program Files\\AutoHotkey\\AutoHotkeyU64.exe",
        r"C:\\Program Files\\AutoHotkey\\AutoHotkeyU32.exe",
        r"C:\\Program Files\\AutoHotkey\\v2\\AutoHotkey.exe",
        r"C:\\Program Files (x86)\\AutoHotkey\\AutoHotkey.exe",
        r"C:\\Program Files (x86)\\AutoHotkey\\AutoHotkeyU32.exe",
    ]
    for c in candidates:
        if c and os.path.exists(c):
            return c
    return ""
