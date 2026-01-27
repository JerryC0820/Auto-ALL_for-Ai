import os
import sys
import json
import time
import re
import shutil
import zipfile
import tempfile
import threading
import queue
import subprocess
import urllib.request
import webbrowser
import winreg
import tkinter as tk
from tkinter import font as tkfont
from tkinter import ttk, filedialog, messagebox

APP_NAME = "牛马神器"
APP_VERSION = "4.0.25"
INSTALLER_ICON_REL = os.path.join("assets", "date1_appicon", "black.png")
UPDATE_PRODUCT_KEY = "niuma_shenqi"
GITEE_RELEASE_URL = "https://gitee.com/api/v5/repos/chen-bin98/Auto-ALL_for-Ai/releases/latest"
GITHUB_RELEASE_URL = "https://api.github.com/repos/JerryC0820/Auto-ALL_for-Ai/releases/latest"
DEFAULT_SETTINGS_URL_GITEE = "https://gitee.com/chen-bin98/Auto-ALL_for-Ai/raw/main/default_settings.json"
DEFAULT_SETTINGS_URL_GITHUB = "https://raw.githubusercontent.com/JerryC0820/Auto-ALL_for-Ai/main/default_settings.json"
AHK_INSTALLER_NAME = "AutoHotkey_2.0.19_setup.exe"
AHK_INSTALLER_URL_GITEE = "https://gitee.com/chen-bin98/Auto-ALL_for-Ai/raw/main/AutoHotkey_2.0.19_setup.exe"
AHK_INSTALLER_URL_GITHUB = "https://raw.githubusercontent.com/JerryC0820/Auto-ALL_for-Ai/main/AutoHotkey_2.0.19_setup.exe"
SELENIUM_MANAGER_NAME = "selenium-manager.exe"
SELENIUM_MANAGER_URL_GITEE = "https://gitee.com/chen-bin98/Auto-ALL_for-Ai/raw/main/assets/selenium-manager.exe"
SELENIUM_MANAGER_URL_GITHUB = "https://raw.githubusercontent.com/JerryC0820/Auto-ALL_for-Ai/main/assets/selenium-manager.exe"
DOWNLOAD_TIMEOUT = 20
CHUNK_SIZE = 1024 * 512

DEFAULT_REQUIRED_MB = 300
DEFAULT_START_MENU_FOLDER = APP_NAME
FILE_ASSOC_EXTS = [".html", ".htm", ".url"]
FILE_ASSOC_PROGID = "NiuMaHTML"
CONTEXT_MENU_KEY = "OpenWithNiuMa"
UNINSTALL_REG_KEY = f"Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\{APP_NAME}"

COLOR_PAGE_BG = "#F2F2F2"
COLOR_PANEL_BG = "#FFFFFF"
COLOR_BORDER = "#D9D9D9"
COLOR_TEXT_SUB = "#555555"
COLOR_TEXT_HINT = "#333333"

FONT_FAMILY = "Microsoft YaHei UI"
FONT_BASE_SIZE = 12
FONT_TITLE_SIZE = 17
FONT_SMALL_SIZE = 11

TEXT = {
    "window_title": "安装 - {app}（User）",
    "title_agreement": "许可协议",
    "title_install_dir": "选择目标位置",
    "title_start_menu": "选择开始菜单文件夹",
    "title_tasks": "选择附加任务",
    "title_ready": "准备安装",
    "title_progress": "正在安装",
    "title_finish": "安装完成",
    "subtitle_agreement": "继续安装前请阅读下列重要信息。",
    "subtitle_install_dir": "您想将 {app} 安装在什么地方？",
    "subtitle_start_menu": "您想在哪里放置程序的快捷方式？",
    "subtitle_tasks": "您想要安装程序执行哪些附加任务？",
    "subtitle_ready": "安装程序已准备好将 {app} 安装到您的电脑中。",
    "subtitle_progress": "安装程序正在安装 {app} 到您的电脑中，请稍等。",
    "subtitle_finish": "{app} 安装完成",
    "btn_back": "上一步",
    "btn_next": "下一步",
    "btn_install": "安装",
    "btn_cancel": "取消",
    "btn_finish": "完成",
    "agree_yes": "我同意此协议(&A)",
    "agree_no": "我不同意此协议(&D)",
    "agreement_hint": "请仔细阅读下列许可协议。您在继续安装前必须同意这些协议条款。",
    "agreement_terms": "• 牛马神器 服务条款 (",
    "agreement_privacy": "• 牛马神器 用户隐私协议 (",
    "agreement_tail": "请仔细阅读并确认是否同意协议条款。",
    "terms_url": "https://gitee.com/chen-bin98/Auto-ALL_for-Ai/blob/main/TERMS.md",
    "privacy_url": "https://gitee.com/chen-bin98/Auto-ALL_for-Ai/blob/main/PRIVACY.md",
    "install_path_label": "安装目录：",
    "browse": "浏览(&R)...",
    "space_hint": "至少需要 {size} MB 的可用磁盘空间{note}。",
    "space_note_est": "（估算）",
    "start_menu_label": "开始菜单文件夹：",
    "start_menu_skip": "不创建开始菜单文件夹(&D)",
    "task_desktop": "创建桌面快捷方式(&D)",
    "task_file_menu": "将“通过牛马神器打开”加入 Windows 资源管理器文件右键菜单",
    "task_dir_menu": "将“通过牛马神器打开”加入 Windows 资源管理器目录右键菜单",
    "task_file_assoc": "将牛马神器注册为支持文件类型的默认编辑器",
    "source_gitee": "Gitee(国内)",
    "source_github": "GitHub(国外)",
    "task_group_download": "高级 / 下载加速",
    "task_group_download_hint": "仅影响下载速度，不影响安装后的运行。",
    "context_menu_label": "通过牛马神器打开",
    "ready_summary_title": "请确认以下安装选项：",
    "ready_label_path": "目标位置：",
    "ready_label_menu": "开始菜单文件夹：",
    "ready_label_tasks": "附加任务：",
    "ready_none": "无",
    "progress_status_default": "等待开始安装",
    "progress_download": "正在下载 {name} ({source})",
    "progress_download_percent": "正在下载... {percent}% ({done:.1f}MB/{total:.1f}MB)",
    "progress_download_bytes": "正在下载... {done:.1f}MB",
    "detail_download": "当前下载文件：{name}",
    "detail_install": "正在安装文件：{name}",
    "progress_extract": "正在解压安装包...",
    "progress_copy": "正在复制文件...",
    "progress_settings": "正在初始化运行环境...",
    "progress_ahk": "正在安装 AutoHotkey...",
    "progress_shortcut": "正在创建快捷方式...",
    "progress_context": "正在写入系统集成...",
    "progress_unregister": "正在清理临时文件...",
    "progress_done": "安装完成",
    "progress_failed": "安装失败",
    "finish_run": "运行 {app}",
    "finish_desc": "安装程序已在您的电脑中安装了 {app}。",
    "warning_need_path": "请选择安装目录",
    "warning_non_empty": "目标目录非空，是否继续安装？",
    "warning_cancel": "正在安装中，确定要退出吗？",
    "error_install": "安装失败: {msg}",
    "error_fetch_release": "无法获取最新安装包，请检查网络",
    "dialog_title": "提示",
    "dialog_error": "错误",
    "cancelled": "安装已取消",
    "cancelled_status": "正在取消...",
    "step_label": "当前步骤：{step}",
    "error_bad_package": "安装包结构异常",
    "step_download": "下载",
    "step_extract": "解压",
    "step_copy": "复制",
    "step_settings": "初始化",
    "step_integration": "系统集成",
    "detail_prepare_dir": "正在创建目录：{name}",
    "detail_prepare_file": "正在生成文件：{name}",
    "detail_prepare_copy": "正在准备文件：{name}",
    "detail_prepare_download": "正在下载文件：{name}",
    "uninstall_cmd_name": "卸载{app}.cmd",
    "uninstall_shortcut_name": "卸载 {app}.lnk",
}

WINDOW_TITLE = TEXT["window_title"].format(app=APP_NAME)
SOURCE_LABELS_UI = {"gitee": TEXT["source_gitee"], "github": TEXT["source_github"]}


def _fetch_json(url: str):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=10) as resp:
        data = resp.read()
    return json.loads(data.decode("utf-8", errors="ignore"))


def _resource_path(relative_path: str) -> str:
    base = getattr(sys, "_MEIPASS", os.path.abspath("."))
    return os.path.join(base, relative_path)


def _apply_window_icon(root: tk.Tk):
    try:
        icon_path = _resource_path(INSTALLER_ICON_REL)
        if not os.path.exists(icon_path):
            return
        icon_img = tk.PhotoImage(file=icon_path)
        root.iconphoto(True, icon_img)
        root._icon_img = icon_img
    except Exception:
        pass


def _pick_asset(assets):
    best = None
    best_score = -999
    for asset in assets or []:
        name = asset.get("name") or ""
        low = name.lower()
        if not low.endswith(".zip"):
            continue
        score = 0
        if UPDATE_PRODUCT_KEY in low:
            score += 3
        if "package" in low:
            score += 2
        if "source" in low:
            score -= 3
        if score > best_score:
            best = asset
            best_score = score
    return best


def _iter_release_sources(prefer_source: str):
    prefer = "github" if prefer_source == "github" else "gitee"
    primary = (prefer, GITHUB_RELEASE_URL if prefer == "github" else GITEE_RELEASE_URL)
    secondary = ("gitee", GITEE_RELEASE_URL) if prefer == "github" else ("github", GITHUB_RELEASE_URL)
    return [primary, secondary]


def _iter_preferred_urls(prefer_source: str, gitee_url: str, github_url: str):
    if prefer_source == "github":
        return [github_url, gitee_url]
    return [gitee_url, github_url]


def fetch_latest_release(prefer_source: str = "gitee"):
    for source, url in _iter_release_sources(prefer_source):
        try:
            data = _fetch_json(url)
            assets = data.get("assets") or []
            asset = _pick_asset(assets)
            if not asset:
                continue
            download_url = asset.get("browser_download_url") or asset.get("url")
            if not download_url:
                continue
            version = data.get("tag_name") or data.get("name") or ""
            return {
                "source": source,
                "version": version,
                "name": asset.get("name") or "",
                "url": download_url,
                "size": int(asset.get("size") or 0),
            }
        except Exception:
            continue
    raise RuntimeError(TEXT["error_fetch_release"])


def _download_file(url: str, dest: str, progress_cb=None, cancel_cb=None):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=DOWNLOAD_TIMEOUT) as resp:
        total = int(resp.headers.get("Content-Length") or 0)
        done = 0
        with open(dest, "wb") as f:
            while True:
                if cancel_cb and cancel_cb():
                    raise RuntimeError(TEXT["cancelled"])
                chunk = resp.read(CHUNK_SIZE)
                if not chunk:
                    break
                f.write(chunk)
                done += len(chunk)
                if progress_cb:
                    progress_cb(done, total)


def _looks_mojibake(name: str) -> bool:
    if not name:
        return False
    has_non_ascii = any(ord(ch) > 127 for ch in name)
    has_cjk = any("\u4e00" <= ch <= "\u9fff" for ch in name)
    return has_non_ascii and not has_cjk


def _score_name(name: str) -> int:
    score = 0
    if APP_NAME in name:
        score += 10
    if "niuma" in name.lower():
        score += 5
    score += sum(1 for ch in name if "\u4e00" <= ch <= "\u9fff")
    if "\ufffd" in name:
        score -= 10
    if any(ord(ch) < 32 for ch in name):
        score -= 10
    return score


def _fix_mojibake_name(name: str) -> str:
    if not _looks_mojibake(name):
        return name
    try:
        raw = name.encode("cp437")
    except UnicodeEncodeError:
        return name
    candidates = []
    for enc in ("utf-8", "gbk"):
        try:
            cand = raw.decode(enc)
        except UnicodeDecodeError:
            continue
        if cand and cand != name:
            candidates.append(cand)
    if not candidates:
        return name
    best = max(candidates, key=_score_name)
    return best if _score_name(best) > 0 else name


def _fix_extracted_names(root_dir: str):
    for root, dirs, files in os.walk(root_dir, topdown=False):
        for name in files + dirs:
            fixed = _fix_mojibake_name(name)
            if fixed == name:
                continue
            old_path = os.path.join(root, name)
            new_path = os.path.join(root, fixed)
            if os.path.exists(new_path):
                continue
            try:
                os.rename(old_path, new_path)
            except OSError:
                pass


def _find_package_root(root_dir: str):
    for root, dirs, files in os.walk(root_dir):
        if "_internal" not in dirs:
            continue
        exe_candidates = [f for f in files if f.lower().endswith(".exe")]
        if not exe_candidates:
            continue
        if len(exe_candidates) == 1:
            return root, exe_candidates[0]
        for exe_name in exe_candidates:
            if APP_NAME in exe_name or "niuma" in exe_name.lower():
                return root, exe_name
        return root, exe_candidates[0]
    return "", ""


def _copy_package(src_root: str, exe_name: str, install_dir: str):
    os.makedirs(install_dir, exist_ok=True)
    internal_src = os.path.join(src_root, "_internal")
    if os.path.isdir(internal_src):
        shutil.copytree(internal_src, os.path.join(install_dir, "_internal"), dirs_exist_ok=True)
    assets_src = os.path.join(src_root, "assets")
    if os.path.isdir(assets_src):
        shutil.copytree(assets_src, os.path.join(install_dir, "assets"), dirs_exist_ok=True)
    exe_name_fixed = _fix_mojibake_name(exe_name)
    exe_src = os.path.join(src_root, exe_name)
    if exe_name_fixed != exe_name:
        alt_src = os.path.join(src_root, exe_name_fixed)
        if os.path.exists(alt_src):
            exe_name = exe_name_fixed
            exe_src = alt_src
    exe_dst = os.path.join(install_dir, exe_name)
    shutil.copy2(exe_src, exe_dst)
    return exe_dst


def _copy_tree_with_progress(src_root: str, dst_root: str, progress_cb=None, cancel_cb=None):
    for root, dirs, files in os.walk(src_root):
        rel = os.path.relpath(root, src_root)
        dest_dir = dst_root if rel == "." else os.path.join(dst_root, rel)
        os.makedirs(dest_dir, exist_ok=True)
        for name in files:
            if cancel_cb and cancel_cb():
                raise RuntimeError(TEXT["cancelled"])
            src_path = os.path.join(root, name)
            dst_path = os.path.join(dest_dir, name)
            shutil.copy2(src_path, dst_path)
            if progress_cb:
                rel_path = name if rel == "." else os.path.join(rel, name)
                progress_cb(rel_path)


def _download_default_settings(install_dir: str, prefer_source: str = "gitee"):
    settings_path = os.path.join(install_dir, "default_settings.json")
    for url in _iter_preferred_urls(prefer_source, DEFAULT_SETTINGS_URL_GITEE, DEFAULT_SETTINGS_URL_GITHUB):
        try:
            _download_file(url, settings_path)
            return True
        except Exception:
            continue
    return False


def _write_hotkeys_template(dest_path: str):
    try:
        with open(dest_path, "w", encoding="utf-8") as f:
            f.write("; AutoHotkey hotkeys will be generated on first run.\n")
    except Exception:
        pass


def _download_selenium_manager(dest_path: str, prefer_source: str = "gitee"):
    for url in _iter_preferred_urls(prefer_source, SELENIUM_MANAGER_URL_GITEE, SELENIUM_MANAGER_URL_GITHUB):
        try:
            _download_file(url, dest_path)
            return True
        except Exception:
            continue
    return False


def _prepare_runtime_files(install_dir: str, prefer_source: str, detail_cb=None, cancel_cb=None):
    def _detail(msg: str):
        if detail_cb:
            detail_cb(msg)

    def _check_cancel():
        if cancel_cb and cancel_cb():
            raise RuntimeError(TEXT["cancelled"])

    for name in ("_mini_fish_cache", "_mini_fish_icons", "_mini_fish_profile"):
        _detail(TEXT["detail_prepare_dir"].format(name=name))
        os.makedirs(os.path.join(install_dir, name), exist_ok=True)
        _check_cancel()

    default_path = os.path.join(install_dir, "default_settings.json")
    if not os.path.exists(default_path):
        _detail(TEXT["detail_prepare_download"].format(name="default_settings.json"))
        _download_default_settings(install_dir, prefer_source)
        _check_cancel()

    settings_path = os.path.join(install_dir, "_mini_fish_settings.json")
    if not os.path.exists(settings_path):
        _detail(TEXT["detail_prepare_file"].format(name="_mini_fish_settings.json"))
        try:
            if os.path.exists(default_path):
                shutil.copy2(default_path, settings_path)
            else:
                with open(settings_path, "w", encoding="utf-8") as f:
                    f.write("{}")
        except Exception:
            pass
        _check_cancel()

    hotkeys_path = os.path.join(install_dir, "_mini_fish_hotkeys.ahk")
    if not os.path.exists(hotkeys_path):
        _detail(TEXT["detail_prepare_file"].format(name="_mini_fish_hotkeys.ahk"))
        _write_hotkeys_template(hotkeys_path)
        _check_cancel()

    se_dest = os.path.join(install_dir, SELENIUM_MANAGER_NAME)
    if not os.path.exists(se_dest):
        asset_path = os.path.join(install_dir, "assets", SELENIUM_MANAGER_NAME)
        if os.path.exists(asset_path):
            _detail(TEXT["detail_prepare_copy"].format(name=SELENIUM_MANAGER_NAME))
            try:
                shutil.copy2(asset_path, se_dest)
            except Exception:
                pass
        else:
            _detail(TEXT["detail_prepare_download"].format(name=SELENIUM_MANAGER_NAME))
            _download_selenium_manager(se_dest, prefer_source)


def _find_ahk_exe():
    candidates = [
        shutil.which("AutoHotkey.exe"),
        shutil.which("AutoHotkeyU64.exe"),
        shutil.which("AutoHotkeyU32.exe"),
        r"C:\Program Files\AutoHotkey\AutoHotkey.exe",
        r"C:\Program Files\AutoHotkey\AutoHotkeyU64.exe",
        r"C:\Program Files\AutoHotkey\AutoHotkeyU32.exe",
        r"C:\Program Files\AutoHotkey\v2\AutoHotkey.exe",
        r"C:\Program Files (x86)\AutoHotkey\AutoHotkey.exe",
        r"C:\Program Files (x86)\AutoHotkey\AutoHotkeyU32.exe",
    ]
    for c in candidates:
        if c and os.path.exists(c):
            return c
    return ""


def _download_ahk_installer(dest_path: str, prefer_source: str = "gitee"):
    for url in _iter_preferred_urls(prefer_source, AHK_INSTALLER_URL_GITEE, AHK_INSTALLER_URL_GITHUB):
        try:
            _download_file(url, dest_path)
            return True
        except Exception:
            continue
    return False


def _ensure_ahk_installed(temp_dir: str, prefer_source: str = "gitee"):
    if _find_ahk_exe():
        return True
    installer_path = os.path.join(temp_dir, AHK_INSTALLER_NAME)
    if not os.path.exists(installer_path):
        ok = _download_ahk_installer(installer_path, prefer_source)
        if not ok:
            return False
    try:
        subprocess.run([installer_path], check=False)
    except Exception:
        return False
    time.sleep(1.0)
    return bool(_find_ahk_exe())

def _create_first_run_flag(install_dir: str):
    try:
        flag_path = os.path.join(install_dir, "_mini_fish_first_run.flag")
        with open(flag_path, "w", encoding="utf-8") as f:
            f.write("1")
    except Exception:
        pass


def _ps_escape(value: str) -> str:
    return (value or "").replace("'", "''")


def _open_url(url: str):
    try:
        webbrowser.open(url)
    except Exception:
        pass


def _create_shortcut(link_path: str, target_path: str, working_dir: str = "", icon_path: str = ""):
    if not link_path or not target_path:
        return
    work_dir = working_dir or os.path.dirname(target_path)
    icon_path = icon_path or target_path
    cmd = (
        "$WScriptShell = New-Object -ComObject WScript.Shell;"
        f"$Shortcut = $WScriptShell.CreateShortcut('{_ps_escape(link_path)}');"
        f"$Shortcut.TargetPath = '{_ps_escape(target_path)}';"
        f"$Shortcut.WorkingDirectory = '{_ps_escape(work_dir)}';"
        f"$Shortcut.IconLocation = '{_ps_escape(icon_path)}';"
        "$Shortcut.Save()"
    )
    subprocess.run(["powershell", "-NoProfile", "-Command", cmd], check=False, creationflags=0x08000000)


def _create_desktop_shortcut(exe_path: str):
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    if not os.path.isdir(desktop):
        return
    link_path = os.path.join(desktop, f"{APP_NAME}.lnk")
    _create_shortcut(link_path, exe_path, os.path.dirname(exe_path), exe_path)


def _start_menu_base_dir() -> str:
    base = os.environ.get("APPDATA") or os.path.join(os.path.expanduser("~"), "AppData", "Roaming")
    return os.path.join(base, "Microsoft", "Windows", "Start Menu", "Programs")


def _resolve_start_menu_dir(folder_name: str) -> str:
    name = (folder_name or "").strip() or DEFAULT_START_MENU_FOLDER
    if os.path.isabs(name):
        return name
    return os.path.join(_start_menu_base_dir(), name)


def _write_uninstall_script(install_dir: str, exe_path: str, start_menu_dir: str) -> str:
    script_path = os.path.join(install_dir, TEXT["uninstall_cmd_name"].format(app=APP_NAME))
    install_dir = install_dir.replace('"', '')
    start_menu_dir = start_menu_dir.replace('"', '')
    script = f"""@echo off
setlocal
set "APP_NAME={APP_NAME}"
set "INSTALL_DIR={install_dir}"
set "START_MENU={start_menu_dir}"
set "DESKTOP=%USERPROFILE%\\Desktop"

del /f /q "%DESKTOP%\\{APP_NAME}.lnk" >nul 2>nul
rmdir /s /q "%START_MENU%" >nul 2>nul

reg delete "HKCU\\Software\\Classes\\*\\shell\\{CONTEXT_MENU_KEY}" /f >nul 2>nul
reg delete "HKCU\\Software\\Classes\\Directory\\shell\\{CONTEXT_MENU_KEY}" /f >nul 2>nul
reg delete "HKCU\\Software\\Classes\\{FILE_ASSOC_PROGID}" /f >nul 2>nul
for %%E in (.html .htm .url) do (
  for /f "tokens=3*" %%A in ('reg query "HKCU\\Software\\Classes\\%%E" /ve 2^>nul ^| findstr /i "REG_"') do (
    if /i "%%B"=="{FILE_ASSOC_PROGID}" reg delete "HKCU\\Software\\Classes\\%%E" /ve /f >nul 2>nul
  )
  reg delete "HKCU\\Software\\Classes\\%%E\\OpenWithProgids" /v {FILE_ASSOC_PROGID} /f >nul 2>nul
)
reg delete "HKCU\\{UNINSTALL_REG_KEY}" /f >nul 2>nul

start "" cmd /c "ping 127.0.0.1 -n 3 >nul & rmdir /s /q \"%INSTALL_DIR%\""
endlocal
"""
    with open(script_path, "w", encoding="utf-8") as f:
        f.write(script)
    return script_path


def _register_uninstall_entry(install_dir: str, exe_path: str, uninstall_script: str):
    uninstall_cmd = f'cmd.exe /c "{uninstall_script}"'
    try:
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, UNINSTALL_REG_KEY) as key:
            winreg.SetValueEx(key, "DisplayName", 0, winreg.REG_SZ, APP_NAME)
            winreg.SetValueEx(key, "DisplayVersion", 0, winreg.REG_SZ, APP_VERSION)
            winreg.SetValueEx(key, "Publisher", 0, winreg.REG_SZ, APP_NAME)
            winreg.SetValueEx(key, "InstallLocation", 0, winreg.REG_SZ, install_dir)
            winreg.SetValueEx(key, "DisplayIcon", 0, winreg.REG_SZ, exe_path)
            winreg.SetValueEx(key, "UninstallString", 0, winreg.REG_SZ, uninstall_cmd)
            winreg.SetValueEx(key, "NoModify", 0, winreg.REG_DWORD, 1)
            winreg.SetValueEx(key, "NoRepair", 0, winreg.REG_DWORD, 1)
    except OSError:
        pass


def _register_context_menu(exe_path: str, for_directory: bool = False):
    base = r"Software\\Classes\\Directory\\shell" if for_directory else r"Software\\Classes\\*\\shell"
    key_path = base + r"\\" + CONTEXT_MENU_KEY
    command = f"\"{exe_path}\" \"%1\""
    try:
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, TEXT["context_menu_label"])
            winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, exe_path)
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path + r"\\command") as key:
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, command)
    except OSError:
        pass


def _register_file_associations(exe_path: str):
    try:
        progid_key = f"Software\\Classes\\{FILE_ASSOC_PROGID}"
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, progid_key) as key:
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, f"{APP_NAME} HTML Document")
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, progid_key + r"\\DefaultIcon") as key:
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, exe_path)
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, progid_key + r"\\shell\\open\\command") as key:
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, f"\"{exe_path}\" \"%1\"")
        for ext in FILE_ASSOC_EXTS:
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, f"Software\\Classes\\{ext}") as key:
                winreg.SetValueEx(key, "", 0, winreg.REG_SZ, FILE_ASSOC_PROGID)
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, f"Software\\Classes\\{ext}\\OpenWithProgids") as key:
                winreg.SetValueEx(key, FILE_ASSOC_PROGID, 0, winreg.REG_NONE, b"")
    except OSError:
        pass


def _apply_system_integration(exe_path: str, file_menu: bool, dir_menu: bool, file_assoc: bool):
    if file_menu:
        _register_context_menu(exe_path, for_directory=False)
    if dir_menu:
        _register_context_menu(exe_path, for_directory=True)
    if file_assoc:
        _register_file_associations(exe_path)


class InstallerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self._init_fonts()
        self.title(WINDOW_TITLE)
        self.geometry("880x660")
        self.minsize(880, 660)
        self.resizable(False, False)
        _apply_window_icon(self)
        self._load_brand_icon()

        self.queue = queue.Queue()
        self.installing = False
        self.install_done = False
        self.cancel_requested = False
        self.installed_exe_path = ""
        self.release_info = None

        self.agree_var = tk.StringVar(value="no")
        self.path_var = tk.StringVar(value=self._default_install_dir())
        self.start_menu_var = tk.StringVar(value=DEFAULT_START_MENU_FOLDER)
        self.no_start_menu_var = tk.BooleanVar(value=False)
        self.desktop_var = tk.BooleanVar(value=True)
        self.file_menu_var = tk.BooleanVar(value=False)
        self.dir_menu_var = tk.BooleanVar(value=False)
        self.file_assoc_var = tk.BooleanVar(value=False)
        self.source_var = tk.StringVar(value=SOURCE_LABELS_UI["gitee"])
        self.run_after_var = tk.BooleanVar(value=True)

        self.status_var = tk.StringVar(value=TEXT["progress_status_default"])
        self.detail_var = tk.StringVar(value="")
        self.progress_var = tk.DoubleVar(value=0.0)

        self.step = 1

        self.container = tk.Frame(self, bg=COLOR_PAGE_BG)
        self.container.pack(fill="both", expand=True)
        self.container.rowconfigure(0, weight=1)
        self.container.columnconfigure(0, weight=1)

        self.page_container = tk.Frame(self.container, bg=COLOR_PAGE_BG)
        self.page_container.grid(row=0, column=0, sticky="nsew")

        self.frames = {}
        self._build_step1()
        self._build_step2()
        self._build_step3()
        self._build_step4()
        self._build_step5()
        self._build_step6()
        self._build_step7()

        self._build_footer()
        self._show_step(1)

        self.after(200, self._prefetch_release_info)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _init_fonts(self):
        try:
            base = tkfont.nametofont("TkDefaultFont")
            base.configure(family=FONT_FAMILY, size=FONT_BASE_SIZE)
            text = tkfont.nametofont("TkTextFont")
            text.configure(family=FONT_FAMILY, size=FONT_BASE_SIZE)
            heading = tkfont.nametofont("TkHeadingFont")
            heading.configure(family=FONT_FAMILY, size=FONT_BASE_SIZE, weight="bold")
            fixed = tkfont.nametofont("TkFixedFont")
            fixed.configure(family="Consolas", size=FONT_SMALL_SIZE)
            self.option_add("*Font", base)
            style = ttk.Style()
            try:
                style.theme_use("vista")
            except Exception:
                pass
            style.configure("TButton", font=(FONT_FAMILY, FONT_BASE_SIZE))
            style.configure("TCheckbutton", font=(FONT_FAMILY, FONT_BASE_SIZE))
            style.configure("TRadiobutton", font=(FONT_FAMILY, FONT_BASE_SIZE))
            style.configure("TLabel", font=(FONT_FAMILY, FONT_BASE_SIZE))
            style.configure("TEntry", font=(FONT_FAMILY, FONT_BASE_SIZE))
            style.configure("TCombobox", font=(FONT_FAMILY, FONT_BASE_SIZE))
            style.configure("TLabelframe", background=COLOR_PANEL_BG)
            style.configure("TLabelframe.Label", background=COLOR_PANEL_BG, font=(FONT_FAMILY, FONT_BASE_SIZE))
        except Exception:
            pass

    def _load_brand_icon(self):
        self.brand_icon = None
        try:
            icon_path = _resource_path(INSTALLER_ICON_REL)
            if os.path.exists(icon_path):
                img = tk.PhotoImage(file=icon_path)
                target = 72
                if img.width() > target:
                    factor = max(1, int(img.width() / target))
                    img = img.subsample(factor, factor)
                self.brand_icon = img
        except Exception:
            self.brand_icon = None

    def _default_install_dir(self):
        base = os.environ.get("LOCALAPPDATA") or os.path.join(os.path.expanduser("~"), "AppData", "Local")
        return os.path.join(base, "Programs", APP_NAME)

    def _prefetch_release_info(self):
        def _worker():
            try:
                self.release_info = fetch_latest_release(self._get_prefer_source())
            except Exception:
                self.release_info = None
            self.after(0, self._update_space_label)
        threading.Thread(target=_worker, daemon=True).start()

    def _estimate_required_mb(self):
        if self.release_info and self.release_info.get("size"):
            size_mb = int(self.release_info["size"] / 1024 / 1024)
            return max(size_mb + 200, DEFAULT_REQUIRED_MB), ""
        return DEFAULT_REQUIRED_MB, TEXT["space_note_est"]

    def _update_space_label(self):
        if not hasattr(self, "space_label"):
            return
        size_mb, note = self._estimate_required_mb()
        self.space_label.configure(text=TEXT["space_hint"].format(size=size_mb, note=note))

    def _create_page(self, title_key: str, subtitle_key: str):
        frame = tk.Frame(self.page_container, bg=COLOR_PAGE_BG)
        header = tk.Frame(frame, bg=COLOR_PAGE_BG)
        header.pack(fill="x", padx=24, pady=(10, 6))

        title = tk.Label(header, text=TEXT[title_key], font=(FONT_FAMILY, FONT_TITLE_SIZE, "bold"), bg=COLOR_PAGE_BG)
        title.grid(row=0, column=0, sticky="w")
        subtitle = tk.Label(
            header,
            text=TEXT[subtitle_key].format(app=APP_NAME),
            font=(FONT_FAMILY, FONT_BASE_SIZE),
            fg=COLOR_TEXT_SUB,
            bg=COLOR_PAGE_BG,
        )
        subtitle.grid(row=1, column=0, sticky="w", pady=(2, 0))

        if self.brand_icon is not None:
            icon_label = tk.Label(header, image=self.brand_icon, bg=COLOR_PAGE_BG)
            icon_label.grid(row=0, column=1, rowspan=2, sticky="e")

        header.columnconfigure(0, weight=1)

        body_outer = tk.Frame(frame, bg=COLOR_PAGE_BG)
        body_outer.pack(fill="both", expand=True, padx=24, pady=(0, 8))
        body = tk.Frame(body_outer, bg=COLOR_PANEL_BG, highlightbackground=COLOR_BORDER, highlightthickness=1)
        body.pack(fill="both", expand=True)
        content = tk.Frame(body, bg=COLOR_PANEL_BG)
        content.pack(fill="both", expand=True, padx=18, pady=14)
        return frame, content

    def _build_step1(self):
        frame, body = self._create_page("title_agreement", "subtitle_agreement")
        self.frames[1] = frame

        hint = tk.Label(body, text=TEXT["agreement_hint"], font=(FONT_FAMILY, FONT_BASE_SIZE), bg=COLOR_PANEL_BG)
        hint.pack(anchor="w", pady=(0, 8))

        text_frame = tk.Frame(body, bg=COLOR_PANEL_BG)
        text_frame.pack(fill="both", expand=True)
        scrollbar = ttk.Scrollbar(text_frame)
        scrollbar.pack(side="right", fill="y")
        text = tk.Text(
            text_frame,
            height=12,
            wrap="word",
            yscrollcommand=scrollbar.set,
            bg=COLOR_PANEL_BG,
            highlightbackground=COLOR_BORDER,
            highlightthickness=1,
            bd=0,
        )
        text.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=text.yview)
        text.configure(spacing1=2, spacing3=2, font=(FONT_FAMILY, FONT_BASE_SIZE))

        text.insert("end", TEXT["agreement_terms"])
        self._insert_link(text, TEXT["terms_url"], TEXT["terms_url"])
        text.insert("end", ")\n")
        text.insert("end", TEXT["agreement_privacy"])
        self._insert_link(text, TEXT["privacy_url"], TEXT["privacy_url"])
        text.insert("end", ")\n\n")
        text.insert("end", TEXT["agreement_tail"])
        text.configure(state="disabled")

        radio_frame = tk.Frame(body, bg=COLOR_PANEL_BG)
        radio_frame.pack(fill="x", pady=10)
        ttk.Radiobutton(
            radio_frame,
            text=TEXT["agree_yes"],
            value="yes",
            variable=self.agree_var,
            command=self._update_nav,
        ).pack(anchor="w")
        ttk.Radiobutton(
            radio_frame,
            text=TEXT["agree_no"],
            value="no",
            variable=self.agree_var,
            command=self._update_nav,
        ).pack(anchor="w", pady=(2, 0))

    def _insert_link(self, widget: tk.Text, label: str, url: str):
        start = widget.index("end-1c")
        widget.insert("end", label)
        end = widget.index("end-1c")
        tag = f"link_{start}"
        widget.tag_add(tag, start, end)
        widget.tag_config(tag, foreground="#0A5FFF", underline=True)
        widget.tag_bind(tag, "<Button-1>", lambda _e, u=url: _open_url(u))
        widget.tag_bind(tag, "<Enter>", lambda _e: widget.config(cursor="hand2"))
        widget.tag_bind(tag, "<Leave>", lambda _e: widget.config(cursor=""))

    def _build_step2(self):
        frame, body = self._create_page("title_install_dir", "subtitle_install_dir")
        self.frames[2] = frame

        row = tk.Frame(body, bg=COLOR_PANEL_BG)
        row.pack(fill="x", pady=6)
        tk.Label(row, text=TEXT["install_path_label"], font=(FONT_FAMILY, FONT_BASE_SIZE), bg=COLOR_PANEL_BG).pack(side="left")
        entry = ttk.Entry(row, textvariable=self.path_var)
        entry.pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(row, text=TEXT["browse"], command=self._browse_dir).pack(side="left")

        self.space_label = tk.Label(body, text="", font=(FONT_FAMILY, FONT_BASE_SIZE), fg=COLOR_TEXT_HINT, bg=COLOR_PANEL_BG)
        self.space_label.pack(anchor="w", pady=(12, 0))
        self._update_space_label()

    def _build_step3(self):
        frame, body = self._create_page("title_start_menu", "subtitle_start_menu")
        self.frames[3] = frame

        row = tk.Frame(body, bg=COLOR_PANEL_BG)
        row.pack(fill="x", pady=6)
        tk.Label(row, text=TEXT["start_menu_label"], font=(FONT_FAMILY, FONT_BASE_SIZE), bg=COLOR_PANEL_BG).pack(side="left")
        self.start_menu_entry = ttk.Entry(row, textvariable=self.start_menu_var)
        self.start_menu_entry.pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(row, text=TEXT["browse"], command=self._browse_start_menu).pack(side="left")

        skip = ttk.Checkbutton(
            body,
            text=TEXT["start_menu_skip"],
            variable=self.no_start_menu_var,
            command=self._toggle_start_menu,
        )
        skip.pack(anchor="w", pady=(12, 0))

    def _build_step4(self):
        frame, body = self._create_page("title_tasks", "subtitle_tasks")
        self.frames[4] = frame

        ttk.Checkbutton(body, text=TEXT["task_desktop"], variable=self.desktop_var).pack(anchor="w", pady=4)
        ttk.Checkbutton(body, text=TEXT["task_file_menu"], variable=self.file_menu_var).pack(anchor="w", pady=2)
        ttk.Checkbutton(body, text=TEXT["task_dir_menu"], variable=self.dir_menu_var).pack(anchor="w", pady=2)
        ttk.Checkbutton(body, text=TEXT["task_file_assoc"], variable=self.file_assoc_var).pack(anchor="w", pady=2)

        group = ttk.LabelFrame(body, text=TEXT["task_group_download"])
        group.pack(fill="x", pady=(16, 4))
        desc = tk.Label(group, text=TEXT["task_group_download_hint"], font=(FONT_FAMILY, FONT_BASE_SIZE), fg=COLOR_TEXT_HINT, bg=COLOR_PANEL_BG)
        desc.pack(anchor="w", padx=10, pady=(6, 2))
        row = tk.Frame(group, bg=COLOR_PANEL_BG)
        row.pack(fill="x", padx=10, pady=(0, 8))
        ttk.Combobox(
            row,
            textvariable=self.source_var,
            values=[SOURCE_LABELS_UI["gitee"], SOURCE_LABELS_UI["github"]],
            state="readonly",
            width=16,
        ).pack(side="left")

    def _build_step5(self):
        frame, body = self._create_page("title_ready", "subtitle_ready")
        self.frames[5] = frame

        tk.Label(body, text=TEXT["ready_summary_title"], font=(FONT_FAMILY, FONT_BASE_SIZE), bg=COLOR_PANEL_BG).pack(anchor="w", pady=(0, 6))
        self.summary_text = tk.Text(body, height=10, wrap="word", bg=COLOR_PANEL_BG)
        self.summary_text.pack(fill="both", expand=True)
        self.summary_text.configure(state="disabled")

    def _build_step6(self):
        frame, body = self._create_page("title_progress", "subtitle_progress")
        self.frames[6] = frame

        status = tk.Label(body, textvariable=self.status_var, font=(FONT_FAMILY, FONT_BASE_SIZE), bg=COLOR_PANEL_BG)
        status.pack(anchor="w", pady=(4, 2))
        detail = tk.Label(
            body,
            textvariable=self.detail_var,
            font=(FONT_FAMILY, FONT_SMALL_SIZE),
            fg=COLOR_TEXT_HINT,
            bg=COLOR_PANEL_BG,
            wraplength=760,
            justify="left",
        )
        detail.pack(anchor="w", pady=(0, 8))
        bar = ttk.Progressbar(body, maximum=100, variable=self.progress_var)
        bar.pack(fill="x")

    def _build_step7(self):
        frame, body = self._create_page("title_finish", "subtitle_finish")
        self.frames[7] = frame

        finish_row = tk.Frame(body, bg=COLOR_PANEL_BG)
        finish_row.pack(fill="both", expand=True, padx=6, pady=6)
        if self.brand_icon is not None:
            icon_label = tk.Label(finish_row, image=self.brand_icon, bg=COLOR_PANEL_BG)
            icon_label.grid(row=0, column=0, rowspan=4, sticky="nw", padx=(4, 18), pady=(8, 0))

        title = tk.Label(finish_row, text=TEXT["title_finish"], font=(FONT_FAMILY, FONT_TITLE_SIZE, "bold"), bg=COLOR_PANEL_BG)
        title.grid(row=0, column=1, sticky="w", pady=(4, 6))
        desc = tk.Label(
            finish_row,
            text=TEXT["finish_desc"].format(app=APP_NAME),
            font=(FONT_FAMILY, FONT_BASE_SIZE),
            bg=COLOR_PANEL_BG,
            justify="left",
            wraplength=420,
        )
        desc.grid(row=1, column=1, sticky="w")
        ttk.Checkbutton(finish_row, text=TEXT["finish_run"].format(app=APP_NAME), variable=self.run_after_var).grid(
            row=2, column=1, sticky="w", pady=(8, 0)
        )
        finish_row.columnconfigure(1, weight=1)

    def _build_footer(self):
        sep = tk.Frame(self.container, height=1, bg=COLOR_BORDER)
        sep.grid(row=1, column=0, sticky="ew")
        footer = tk.Frame(self.container, bg=COLOR_PAGE_BG)
        footer.grid(row=2, column=0, sticky="ew", padx=24, pady=(8, 10))
        footer.columnconfigure(0, weight=1)

        self.btn_back = ttk.Button(footer, text=TEXT["btn_back"], command=self._go_back)
        self.btn_next = ttk.Button(footer, text=TEXT["btn_next"], command=self._go_next)
        self.btn_cancel = ttk.Button(footer, text=TEXT["btn_cancel"], command=self._on_cancel)
        self.btn_finish = ttk.Button(footer, text=TEXT["btn_finish"], command=self._finish)

        self.btn_back.grid(row=0, column=1, padx=(0, 6))
        self.btn_next.grid(row=0, column=2, padx=(0, 6))
        self.btn_cancel.grid(row=0, column=3)

    def _toggle_start_menu(self):
        state = "disabled" if self.no_start_menu_var.get() else "normal"
        self.start_menu_entry.configure(state=state)

    def _show_step(self, step: int):
        self.step = step
        for f in self.frames.values():
            f.pack_forget()
        self.frames[step].pack(fill="both", expand=True)
        if step == 5:
            self._refresh_summary()
        if step == 6 and not self.installing and not self.install_done:
            self._start_install()
        self._update_nav()

    def _update_nav(self):
        if self.step == 1:
            self.btn_back.configure(state="disabled")
            self.btn_next.configure(text=TEXT["btn_next"], state="normal" if self.agree_var.get() == "yes" else "disabled")
            self._show_finish_button(False)
        elif self.step in (2, 3, 4):
            self.btn_back.configure(state="normal")
            self.btn_next.configure(text=TEXT["btn_next"], state="normal")
            self._show_finish_button(False)
        elif self.step == 5:
            self.btn_back.configure(state="normal")
            self.btn_next.configure(text=TEXT["btn_install"], state="normal")
            self._show_finish_button(False)
        elif self.step == 6:
            self.btn_back.configure(state="disabled")
            self.btn_next.configure(state="disabled")
            self._show_finish_button(False)
        elif self.step == 7:
            self._show_finish_button(True)

    def _show_finish_button(self, show: bool):
        if show:
            self.btn_back.grid_remove()
            self.btn_next.grid_remove()
            self.btn_cancel.grid_remove()
            self.btn_finish.grid(row=0, column=3)
        else:
            self.btn_finish.grid_remove()
            self.btn_back.grid()
            self.btn_next.grid()
            self.btn_cancel.grid()

    def _go_back(self):
        if self.step > 1 and self.step < 6:
            self._show_step(self.step - 1)

    def _go_next(self):
        if self.step == 1:
            if self.agree_var.get() != "yes":
                return
            self._show_step(2)
            return
        if self.step == 2:
            path = self.path_var.get().strip()
            if not path:
                messagebox.showwarning(TEXT["dialog_title"], TEXT["warning_need_path"])
                return
            if os.path.isdir(path) and os.listdir(path):
                ok = messagebox.askyesno(TEXT["dialog_title"], TEXT["warning_non_empty"])
                if not ok:
                    return
            self._show_step(3)
            return
        if self.step == 3:
            self._show_step(4)
            return
        if self.step == 4:
            self._show_step(5)
            return
        if self.step == 5:
            self._show_step(6)

    def _refresh_summary(self):
        tasks = []
        if self.desktop_var.get():
            tasks.append(TEXT["task_desktop"])
        if self.file_menu_var.get():
            tasks.append(TEXT["task_file_menu"])
        if self.dir_menu_var.get():
            tasks.append(TEXT["task_dir_menu"])
        if self.file_assoc_var.get():
            tasks.append(TEXT["task_file_assoc"])
        if not tasks:
            tasks.append(TEXT["ready_none"])

        start_menu = TEXT["ready_none"] if self.no_start_menu_var.get() else _resolve_start_menu_dir(self.start_menu_var.get())

        summary = (
            f"{TEXT['ready_label_path']}{self.path_var.get().strip()}\n\n"
            f"{TEXT['ready_label_menu']}{start_menu}\n\n"
            f"{TEXT['ready_label_tasks']}\n- " + "\n- ".join(tasks)
        )
        self.summary_text.configure(state="normal")
        self.summary_text.delete("1.0", "end")
        self.summary_text.insert("1.0", summary)
        self.summary_text.configure(state="disabled")

    def _browse_dir(self):
        initial = self.path_var.get().strip() or self._default_install_dir()
        target = filedialog.askdirectory(initialdir=initial, title=TEXT["title_install_dir"])
        if target:
            self.path_var.set(target)
            self._update_space_label()

    def _browse_start_menu(self):
        base = _start_menu_base_dir()
        target = filedialog.askdirectory(initialdir=base, title=TEXT["title_start_menu"])
        if target:
            self.start_menu_var.set(target)

    def _get_prefer_source(self) -> str:
        label = (self.source_var.get() or "").strip()
        for key, text in SOURCE_LABELS_UI.items():
            if text == label:
                return key
        return "gitee"

    def _start_install(self):
        self.installing = True
        self.cancel_requested = False
        self.install_done = False
        self.status_var.set(TEXT["progress_status_default"])
        self.detail_var.set("")
        self.progress_var.set(0)
        thread = threading.Thread(target=self._install_worker, daemon=True)
        thread.start()
        self.after(100, self._poll_queue)

    def _check_cancel(self):
        if self.cancel_requested:
            raise RuntimeError(TEXT["cancelled"])

    def _install_worker(self):
        temp_dir = None
        install_dir = self.path_var.get().strip()
        install_dir_preexists = os.path.exists(install_dir)
        try:
            prefer_source = self._get_prefer_source()
            self._check_cancel()
            info = self.release_info or fetch_latest_release(prefer_source)
            self.queue.put(("step", TEXT["step_download"]))
            self.queue.put(("status", TEXT["progress_download"].format(name=info["name"], source=info["source"])))
            self.queue.put(("detail", TEXT["detail_download"].format(name=info["name"])))
            temp_dir = tempfile.mkdtemp(prefix="niuma_installer_")
            zip_path = os.path.join(temp_dir, info["name"] or "package.zip")

            def _progress(done, total):
                self.queue.put(("progress", done, total))

            _download_file(info["url"], zip_path, _progress, cancel_cb=lambda: self.cancel_requested)
            self._check_cancel()

            self.queue.put(("step", TEXT["step_extract"]))
            self.queue.put(("status", TEXT["progress_extract"]))
            extract_dir = os.path.join(temp_dir, "extracted")
            os.makedirs(extract_dir, exist_ok=True)
            with zipfile.ZipFile(zip_path, "r") as zf:
                zf.extractall(extract_dir)
            _fix_extracted_names(extract_dir)
            root, exe_name = _find_package_root(extract_dir)
            if not root or not exe_name:
                raise RuntimeError(TEXT["error_bad_package"])

            self._check_cancel()
            self.queue.put(("step", TEXT["step_copy"]))
            self.queue.put(("status", TEXT["progress_copy"]))
            os.makedirs(install_dir, exist_ok=True)
            def _copy_progress(rel_path):
                self.queue.put(("detail", TEXT["detail_install"].format(name=rel_path)))

            _copy_tree_with_progress(root, install_dir, progress_cb=_copy_progress, cancel_cb=lambda: self.cancel_requested)
            exe_path = os.path.join(install_dir, exe_name)

            self._check_cancel()
            self.queue.put(("step", TEXT["step_settings"]))
            self.queue.put(("status", TEXT["progress_settings"]))
            _prepare_runtime_files(
                install_dir,
                prefer_source,
                detail_cb=lambda msg: self.queue.put(("detail", msg)),
                cancel_cb=lambda: self.cancel_requested,
            )
            _create_first_run_flag(install_dir)

            self.queue.put(("step", TEXT["step_integration"]))
            self.queue.put(("status", TEXT["progress_context"]))
            start_menu_dir = _resolve_start_menu_dir(self.start_menu_var.get())
            uninstall_script = _write_uninstall_script(install_dir, exe_path, start_menu_dir)
            _register_uninstall_entry(install_dir, exe_path, uninstall_script)

            if not self.no_start_menu_var.get():
                os.makedirs(start_menu_dir, exist_ok=True)
                _create_shortcut(os.path.join(start_menu_dir, f"{APP_NAME}.lnk"), exe_path, install_dir, exe_path)
                _create_shortcut(
                    os.path.join(start_menu_dir, TEXT["uninstall_shortcut_name"].format(app=APP_NAME)),
                    uninstall_script,
                    install_dir,
                    exe_path,
                )

            if self.desktop_var.get():
                self.queue.put(("status", TEXT["progress_shortcut"]))
                _create_desktop_shortcut(exe_path)

            _apply_system_integration(exe_path, self.file_menu_var.get(), self.dir_menu_var.get(), self.file_assoc_var.get())

            self.queue.put(("done", exe_path))
        except RuntimeError as e:
            if TEXT["cancelled"] in str(e):
                if temp_dir and os.path.isdir(temp_dir):
                    shutil.rmtree(temp_dir, ignore_errors=True)
                if not install_dir_preexists and os.path.isdir(install_dir):
                    shutil.rmtree(install_dir, ignore_errors=True)
                self.queue.put(("cancelled",))
            else:
                self.queue.put(("error", str(e)))
        except Exception as e:
            if temp_dir and os.path.isdir(temp_dir):
                shutil.rmtree(temp_dir, ignore_errors=True)
            if not install_dir_preexists and os.path.isdir(install_dir):
                shutil.rmtree(install_dir, ignore_errors=True)
            self.queue.put(("error", str(e)))
        finally:
            if temp_dir and os.path.isdir(temp_dir):
                shutil.rmtree(temp_dir, ignore_errors=True)

    def _poll_queue(self):
        try:
            while True:
                msg = self.queue.get_nowait()
                if not msg:
                    continue
                if msg[0] == "step":
                    self.detail_var.set(TEXT["step_label"].format(step=msg[1]))
                elif msg[0] == "detail":
                    self.detail_var.set(msg[1])
                elif msg[0] == "status":
                    self.status_var.set(msg[1])
                elif msg[0] == "progress":
                    done, total = msg[1], msg[2]
                    if total > 0:
                        percent = int(done * 100 / total)
                        self.progress_var.set(percent)
                        self.status_var.set(
                            TEXT["progress_download_percent"].format(
                                percent=percent,
                                done=done / 1024 / 1024,
                                total=total / 1024 / 1024,
                            )
                        )
                    else:
                        self.status_var.set(TEXT["progress_download_bytes"].format(done=done / 1024 / 1024))
                elif msg[0] == "done":
                    self.install_done = True
                    self.installing = False
                    self.installed_exe_path = msg[1]
                    self.progress_var.set(100)
                    self.status_var.set(TEXT["progress_done"])
                    self._show_step(7)
                elif msg[0] == "error":
                    self.installing = False
                    self.status_var.set(TEXT["progress_failed"])
                    messagebox.showerror(TEXT["dialog_error"], TEXT["error_install"].format(msg=msg[1]))
                    self._show_step(5)
                elif msg[0] == "cancelled":
                    self.installing = False
                    self.status_var.set(TEXT["cancelled"])
                    messagebox.showinfo(TEXT["dialog_title"], TEXT["cancelled"])
                    self._show_step(5)
        except queue.Empty:
            pass
        if self.installing:
            self.after(100, self._poll_queue)

    def _finish(self):
        if self.install_done and self.installed_exe_path and self.run_after_var.get():
            try:
                subprocess.Popen([self.installed_exe_path], cwd=os.path.dirname(self.installed_exe_path))
            except Exception:
                pass
        self.destroy()

    def _on_cancel(self):
        if self.installing:
            ok = messagebox.askyesno(TEXT["dialog_title"], TEXT["warning_cancel"])
            if not ok:
                return
            self.cancel_requested = True
            self.status_var.set(TEXT["cancelled_status"])
            self.btn_cancel.configure(state="disabled")
            return
        self.destroy()

    def _on_close(self):
        if self.installing:
            ok = messagebox.askyesno(TEXT["dialog_title"], TEXT["warning_cancel"])
            if not ok:
                return
        self._on_cancel()


if __name__ == "__main__":
    app = InstallerApp()
    app.mainloop()
