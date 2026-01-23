import os, sys, time, json, socket, shutil, ctypes, subprocess, importlib.util, urllib.request, urllib.parse, urllib.error, base64, re, hashlib, zipfile
from ctypes import wintypes
CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)

# -------------------- Auto install selenium --------------------
def _has_pkg(name: str) -> bool:
    return importlib.util.find_spec(name) is not None

def _pip_install(pkgs):
    cmd = [sys.executable, "-m", "pip", "install", "-U", *pkgs]
    subprocess.check_call(cmd)

def ensure_deps():
    missing = []
    if not _has_pkg("selenium"):
        missing.append("selenium")
    if not _has_pkg("PySide6"):
        missing.append("PySide6")
    if missing:
        _pip_install(missing)
        os.execv(sys.executable, [sys.executable] + sys.argv)

def ensure_pillow():
    if _has_pkg("PIL"):
        return True
    try:
        _pip_install(["pillow"])
    except Exception:
        return False
    return True

ensure_deps()

from PySide6 import QtCore, QtGui, QtWidgets

class ToggleSwitch(QtWidgets.QAbstractButton):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setCheckable(True)
        self.setCursor(QtCore.Qt.PointingHandCursor)
        self.setSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        self._w = 44
        self._h = 22

    def sizeHint(self):
        return QtCore.QSize(self._w, self._h)

    def minimumSizeHint(self):
        return self.sizeHint()

    def paintEvent(self, event):
        painter = QtGui.QPainter(self)
        painter.setRenderHint(QtGui.QPainter.Antialiasing)
        rect = QtCore.QRectF(0, 0, self._w, self._h)
        radius = rect.height() / 2
        if self.isEnabled():
            bg = QtGui.QColor(59, 130, 246) if self.isChecked() else QtGui.QColor(200, 200, 200)
        else:
            bg = QtGui.QColor(180, 180, 180)
        painter.setPen(QtCore.Qt.NoPen)
        painter.setBrush(bg)
        painter.drawRoundedRect(rect.adjusted(1, 1, -1, -1), radius, radius)
        handle_size = self._h - 6
        y = (self._h - handle_size) / 2
        x = self._w - handle_size - 3 if self.isChecked() else 3
        painter.setBrush(QtGui.QColor(255, 255, 255))
        painter.drawEllipse(QtCore.QRectF(x, y, handle_size, handle_size))
        painter.end()

class BrowserStartWorker(QtCore.QObject):
    finished = QtCore.Signal(object, object, int, str)

    def __init__(self, chrome_exe: str, url: str, w: int, h: int, x: int, y: int,
                 profile_dir: str, audio_enabled: bool):
        super().__init__()
        self.chrome_exe = chrome_exe
        self.url = url
        self.w = w
        self.h = h
        self.x = x
        self.y = y
        self.profile_dir = profile_dir
        self.audio_enabled = audio_enabled

    @QtCore.Slot()
    def run(self):
        proc, driver, port, err = launch_browser_session(
            self.chrome_exe, self.url, self.w, self.h, self.x, self.y,
            self.profile_dir, self.audio_enabled
        )
        self.finished.emit(proc, driver, port, err)

class UpdateCheckWorker(QtCore.QObject):
    finished = QtCore.Signal(object, str)

    def __init__(self, current_version: str, update_source: str):
        super().__init__()
        self.current_version = current_version
        self.update_source = update_source

    @QtCore.Slot()
    def run(self):
        info, err = check_update_by_source(self.current_version, self.update_source)
        self.finished.emit(info, err)

class UpdateDownloadWorker(QtCore.QObject):
    progress = QtCore.Signal(int, int)
    finished = QtCore.Signal(object, str)

    def __init__(self, info: dict):
        super().__init__()
        self.info = info

    @QtCore.Slot()
    def run(self):
        def _progress(done, total):
            try:
                self.progress.emit(int(done), int(total))
            except Exception:
                pass
        result, err = download_update_package(self.info, _progress)
        self.finished.emit(result, err)

# -------------------- Chrome find / launch --------------------
def find_chrome_exe():
    p = shutil.which("chrome") or shutil.which("chrome.exe")
    if p:
        return p
    candidates = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
        # Edge fallback
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\Edge\Application\msedge.exe"),
    ]
    for c in candidates:
        if c and os.path.exists(c):
            return c
    return None

def normalize_url(u: str) -> str:
    u = (u or "").strip()
    if not u:
        return "https://www.douyin.com"
    if not (u.startswith("http://") or u.startswith("https://")):
        u = "https://" + u
    return u

def pick_free_port() -> int:
    s = socket.socket()
    s.bind(("127.0.0.1", 0))
    port = s.getsockname()[1]
    s.close()
    return port

def read_devtools_port(profile_dir: str, min_mtime: float = 0.0) -> int:
    try:
        path = os.path.join(profile_dir or "", "DevToolsActivePort")
        if not os.path.exists(path):
            return 0
        if min_mtime:
            try:
                if os.path.getmtime(path) < min_mtime:
                    return 0
            except Exception:
                return 0
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            line = (f.readline() or "").strip()
        return int(line) if line.isdigit() else 0
    except Exception:
        return 0

APP_VERSION = "4.0.13"
APP_TITLE = f"牛马神器V{APP_VERSION}"
GITHUB_REPO_URL = "https://github.com/JerryC0820/Auto-ALL_for-Ai"
GITEE_REPO_URL = "https://gitee.com/chen-bin98/Auto-ALL_for-Ai"
GITHUB_HOME_URL = "https://github.com/JerryC0820"
GITEE_HOME_URL = "https://gitee.com/chen-bin98"
DEFAULT_PANEL_TITLES = {
    "",
    "mini",
    "牛马爱摸鱼V2.0.1",
    "牛马爱摸鱼V2.0.19",
    "牛马爱摸鱼V2.0.20",
    "牛马神器V2.0.3",
    "牛马神器V2.0.45",
    "牛马神器V2.0.46",
    "牛马神器V4.0.10",
    "牛马神器V4.0.11",
    "牛马神器V4.0.12",
    "牛马神器V4.0.13",
}

GITEE_API_BASE = "https://gitee.com/api/v5"
GITEE_REPO_API = f"{GITEE_API_BASE}/repos/chen-bin98/Auto-ALL_for-Ai"
GITHUB_API_BASE = "https://api.github.com"
GITHUB_REPO_API = f"{GITHUB_API_BASE}/repos/JerryC0820/Auto-ALL_for-Ai"
UPDATE_PRODUCT_KEY = "niuma_shenqi"
UPDATE_CHECK_TIMEOUT = 8
UPDATE_DOWNLOAD_TIMEOUT = 20
UPDATE_CHUNK_SIZE = 1024 * 512
UPDATE_SOURCE_AUTO = "auto"
UPDATE_SOURCE_GITEE = "gitee"
UPDATE_SOURCE_GITHUB = "github"
DEFAULT_UPDATE_SOURCE = UPDATE_SOURCE_AUTO
UPDATE_SOURCE_LABELS = {
    UPDATE_SOURCE_AUTO: "自动(优先国内)",
    UPDATE_SOURCE_GITEE: "Gitee(国内)",
    UPDATE_SOURCE_GITHUB: "GitHub(国外)",
}
UPDATE_SOURCE_LABEL_TO_KEY = {v: k for k, v in UPDATE_SOURCE_LABELS.items()}
DEFAULT_SETTINGS_URL_GITEE = "https://gitee.com/chen-bin98/Auto-ALL_for-Ai/raw/main/default_settings.json"
DEFAULT_SETTINGS_URL_GITHUB = "https://raw.githubusercontent.com/JerryC0820/Auto-ALL_for-Ai/main/default_settings.json"

IS_FROZEN = bool(getattr(sys, "frozen", False))
RESOURCE_DIR = os.path.abspath(getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__))))
APP_DIR = os.path.dirname(os.path.abspath(sys.executable)) if IS_FROZEN else os.path.dirname(os.path.abspath(__file__))
PROFILE_DIR_BASE = os.path.join(APP_DIR, "_mini_fish_profile")
CACHE_DIR = os.path.join(APP_DIR, "_mini_fish_cache")
ICON_DIR = os.path.join(APP_DIR, "_mini_fish_icons")
ASSETS_DIR = os.path.join(RESOURCE_DIR, "assets")
SETTINGS_FILENAME = "_mini_fish_settings.json"
SETTINGS_BOOTSTRAP_FLAG = "_mini_fish_settings_bootstrap.json"
FIRST_RUN_FLAG = "_mini_fish_first_run.flag"
SETTINGS_PATH = os.path.join(APP_DIR, SETTINGS_FILENAME)
ALLOWED_ICON_EXTS = {".png", ".gif", ".ico"}
UPDATE_ICON_DIR = os.path.join(ASSETS_DIR, "update_icon_20260122")
APP_MUTEX_NAME = "Global\\MiniFish_" + hashlib.md5(APP_DIR.lower().encode("utf-8")).hexdigest()
INSTANCE_MUTEX_HANDLE = None
LOCAL_ICON_FILES = {}
LOCAL_ICON_CHOICES = []

def _register_local_icon(key: str, path: str):
    if not key or not path or not os.path.exists(path):
        return
    LOCAL_ICON_FILES[key] = path
    if key not in LOCAL_ICON_CHOICES:
        LOCAL_ICON_CHOICES.append(key)

if os.path.isdir(UPDATE_ICON_DIR):
    for entry in sorted(os.listdir(UPDATE_ICON_DIR)):
        ext = os.path.splitext(entry)[1].lower()
        if ext not in ALLOWED_ICON_EXTS:
            continue
        key = os.path.splitext(entry)[0]
        _register_local_icon(key, os.path.join(UPDATE_ICON_DIR, entry))

# Prefer PS.ico for Photoshop; keep PH.ico for Pornhub.
if "PS" in LOCAL_ICON_FILES:
    LOCAL_ICON_FILES["photoshop"] = LOCAL_ICON_FILES["PS"]

def get_instance_id():
    for a in sys.argv[1:]:
        if a.startswith("--instance-id="):
            return a.split("=", 1)[1].strip()
    return ""

INSTANCE_ID = get_instance_id()
PROFILE_DIR = PROFILE_DIR_BASE if not INSTANCE_ID else os.path.join(APP_DIR, f"_mini_fish_profile_{INSTANCE_ID}")
AHK_SCRIPT_PATH = os.path.join(APP_DIR, "_mini_fish_hotkeys.ahk")
AHK_CMD_PATH = os.path.join(APP_DIR, f"_mini_fish_ahk_cmd_{INSTANCE_ID or 'main'}.txt")
AHK_EVT_PATH = os.path.join(APP_DIR, f"_mini_fish_ahk_evt_{INSTANCE_ID or 'main'}.txt")

def find_ahk_exe():
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

def _get_file_version_major(path: str) -> int:
    try:
        size = ctypes.windll.version.GetFileVersionInfoSizeW(path, None)
        if not size:
            return 0
        buf = ctypes.create_string_buffer(size)
        if not ctypes.windll.version.GetFileVersionInfoW(path, 0, size, buf):
            return 0
        value = ctypes.c_void_p()
        length = wintypes.UINT()
        if not ctypes.windll.version.VerQueryValueW(buf, "\\", ctypes.byref(value), ctypes.byref(length)):
            return 0
        class VS_FIXEDFILEINFO(ctypes.Structure):
            _fields_ = [
                ("dwSignature", wintypes.DWORD),
                ("dwStrucVersion", wintypes.DWORD),
                ("dwFileVersionMS", wintypes.DWORD),
                ("dwFileVersionLS", wintypes.DWORD),
                ("dwProductVersionMS", wintypes.DWORD),
                ("dwProductVersionLS", wintypes.DWORD),
                ("dwFileFlagsMask", wintypes.DWORD),
                ("dwFileFlags", wintypes.DWORD),
                ("dwFileOS", wintypes.DWORD),
                ("dwFileType", wintypes.DWORD),
                ("dwFileSubtype", wintypes.DWORD),
                ("dwFileDateMS", wintypes.DWORD),
                ("dwFileDateLS", wintypes.DWORD),
            ]
        info = VS_FIXEDFILEINFO.from_address(value.value)
        if info.dwSignature != 0xFEEF04BD:
            return 0
        return int(info.dwFileVersionMS >> 16)
    except Exception:
        return 0

def is_ahk_v2(path: str) -> bool:
    low = (path or "").lower()
    if "\\v2\\" in low:
        return True
    major = _get_file_version_major(path)
    return major >= 2

def ensure_ahk_script(path: str, use_v2: bool = False):
    script_v1 = r"""#NoEnv
#NoTrayIcon
#SingleInstance Force
#Persistent
SetTitleMatchMode, 2
SetBatchLines, -1
FileEncoding, UTF-8

mode = %1%
panel_title = %2%
browser_title = %3%
hk_toggle = %4%
hk_lock = %5%
hk_close = %6%
cmd_file = %7%
evt_file = %8%

if (mode != "daemon") {
    ExitApp
}

init_toggle := hk_toggle
init_lock := hk_lock
init_close := hk_close
hk_toggle := ""
hk_lock := ""
hk_close := ""
UpdateHotkeys(init_toggle, init_lock, init_close)
SetTimer, CheckCmd, 200
return

GetPanelId() {
    global panel_title
    if (panel_title = "")
        return 0
    WinGet, id, ID, %panel_title%
    return id
}

GetBrowserId() {
    global browser_title
    if (browser_title = "")
        return 0
    WinGet, id, ID, %browser_title% ahk_class Chrome_WidgetWin_1
    if (!id)
        WinGet, id, ID, %browser_title% ahk_class Chrome_WidgetWin_0
    if (!id)
        WinGet, id, ID, %browser_title%
    return id
}

MinimizeBoth() {
    idp := GetPanelId()
    idb := GetBrowserId()
    if (idb)
        WinMinimize, ahk_id %idb%
    if (idp)
        WinMinimize, ahk_id %idp%
    if (!idb && !idp)
        WinMinimize, A
}

HideBoth() {
    idp := GetPanelId()
    idb := GetBrowserId()
    if (idp)
        WinHide, ahk_id %idp%
    if (idb)
        WinHide, ahk_id %idb%
}

ShowBoth() {
    idb := GetBrowserId()
    idp := GetPanelId()
    if (idb) {
        WinShow, ahk_id %idb%
        WinRestore, ahk_id %idb%
    }
    if (idp) {
        WinShow, ahk_id %idp%
        WinRestore, ahk_id %idp%
        WinActivate, ahk_id %idp%
    }
    if (!idb && !idp)
        WinRestore, A
}

CloseBoth() {
    idb := GetBrowserId()
    idp := GetPanelId()
    if (idb)
        WinClose, ahk_id %idb%
    if (idp)
        WinClose, ahk_id %idp%
    if (!idb && !idp)
        WinClose, A
}

DoMinimize:
    MinimizeBoth()
return

DoRestore:
    ShowBoth()
return

DoClose:
    CloseBoth()
return

ApplyTopmost(mode) {
    idb := GetBrowserId()
    if (!idb)
        return 0
    target := "ahk_id " idb
    if (mode = "on")
        WinSet, AlwaysOnTop, On, %target%
    else if (mode = "off")
        WinSet, AlwaysOnTop, Off, %target%
    else {
        WinGet, ex, ExStyle, %target%
        if (ex & 0x8)
            WinSet, AlwaysOnTop, Off, %target%
        else
            WinSet, AlwaysOnTop, On, %target%
    }
    return idb
}

WriteEvent(kind) {
    global evt_file
    if (!evt_file)
        return
    FileDelete, %evt_file%
    FileAppend, %kind%, %evt_file%
}

^#!t::
    ApplyTopmost("on")
    WriteEvent("top_on")
return

^+#t::
    ApplyTopmost("off")
    WriteEvent("top_off")
return

UpdateHotkeys(new_toggle, new_lock, new_close) {
    global hk_toggle, hk_lock, hk_close
    if (hk_toggle != "")
        Hotkey, %hk_toggle%, Off
    if (hk_lock != "")
        Hotkey, %hk_lock%, Off
    if (hk_close != "")
        Hotkey, %hk_close%, Off
    hk_toggle := new_toggle
    hk_lock := new_lock
    hk_close := new_close
    if (hk_toggle != "")
        Hotkey, %hk_toggle%, DoMinimize, On
    if (hk_lock != "")
        Hotkey, %hk_lock%, DoRestore, On
    if (hk_close != "")
        Hotkey, %hk_close%, DoClose, On
}

CheckCmd:
    if (!cmd_file)
        return
    if !FileExist(cmd_file)
        return
    FileRead, cmd, %cmd_file%
    if (cmd = "")
        return
    FileDelete, %cmd_file%
    StringSplit, parts, cmd, |
    if (parts1 = "top_on")
        ApplyTopmost("on")
    else if (parts1 = "top_off")
        ApplyTopmost("off")
    else if (parts1 = "top_toggle")
        ApplyTopmost("toggle")
    else if (parts1 = "update") {
        panel_title := parts2
        browser_title := parts3
        UpdateHotkeys(parts4, parts5, parts6)
    }
return
"""
    script_v2 = r"""#Requires AutoHotkey v2.0
#NoTrayIcon
#SingleInstance Force
SetTitleMatchMode 2
A_FileEncoding := "UTF-8"

if (A_Args.Length < 8)
    ExitApp

mode := A_Args[1]
panel_title := A_Args[2]
browser_title := A_Args[3]
hk_toggle := A_Args[4]
hk_lock := A_Args[5]
hk_close := A_Args[6]
cmd_file := A_Args[7]
evt_file := A_Args[8]

if (mode != "daemon")
    ExitApp

UpdateHotkeys(hk_toggle, hk_lock, hk_close)
SetTimer(CheckCmd, 200)
Hotkey("^#!t", DoTopOn)
Hotkey("^+#t", DoTopOff)

GetPanelId() {
    global panel_title
    if (panel_title = "")
        return 0
    try return WinGetID(panel_title)
    catch
        return 0
}

GetBrowserId() {
    global browser_title
    if (browser_title = "")
        return 0
    id := 0
    try id := WinGetID(browser_title " ahk_class Chrome_WidgetWin_1")
    catch
        id := 0
    if (!id) {
        try id := WinGetID(browser_title " ahk_class Chrome_WidgetWin_0")
        catch
            id := 0
    }
    if (!id) {
        try id := WinGetID(browser_title)
        catch
            id := 0
    }
    return id
}

MinimizeBoth() {
    idp := GetPanelId()
    idb := GetBrowserId()
    if (idb)
        WinMinimize("ahk_id " idb)
    if (idp)
        WinMinimize("ahk_id " idp)
    if (!idb && !idp)
        WinMinimize("A")
}

HideBoth() {
    idp := GetPanelId()
    idb := GetBrowserId()
    if (idp)
        WinHide("ahk_id " idp)
    if (idb)
        WinHide("ahk_id " idb)
}

ShowBoth() {
    idb := GetBrowserId()
    idp := GetPanelId()
    if (idb) {
        WinShow("ahk_id " idb)
        WinRestore("ahk_id " idb)
    }
    if (idp) {
        WinShow("ahk_id " idp)
        WinRestore("ahk_id " idp)
        WinActivate("ahk_id " idp)
    }
    if (!idb && !idp)
        WinRestore("A")
}

CloseBoth() {
    idb := GetBrowserId()
    idp := GetPanelId()
    if (idb)
        WinClose("ahk_id " idb)
    if (idp)
        WinClose("ahk_id " idp)
    if (!idb && !idp)
        WinClose("A")
}

DoMinimize(*) {
    MinimizeBoth()
}

DoRestore(*) {
    ShowBoth()
}

DoClose(*) {
    CloseBoth()
}

ApplyTopmost(mode) {
    idb := GetBrowserId()
    if (!idb)
        return 0
    target := "ahk_id " idb
    if (mode = "on")
        WinSetAlwaysOnTop(1, target)
    else if (mode = "off")
        WinSetAlwaysOnTop(0, target)
    else {
        ex := WinGetExStyle(target)
        if (ex & 0x8)
            WinSetAlwaysOnTop(0, target)
        else
            WinSetAlwaysOnTop(1, target)
    }
    return idb
}

WriteEvent(kind) {
    global evt_file
    if (evt_file = "")
        return
    try FileDelete(evt_file)
    try FileAppend(kind, evt_file)
}

DoTopOn(*) {
    ApplyTopmost("on")
    WriteEvent("top_on")
}

DoTopOff(*) {
    ApplyTopmost("off")
    WriteEvent("top_off")
}

UpdateHotkeys(new_toggle, new_lock, new_close) {
    global hk_toggle, hk_lock, hk_close
    if (hk_toggle != "") {
        try {
            Hotkey(hk_toggle, "Off")
        } catch {
        }
    }
    if (hk_lock != "") {
        try {
            Hotkey(hk_lock, "Off")
        } catch {
        }
    }
    if (hk_close != "") {
        try {
            Hotkey(hk_close, "Off")
        } catch {
        }
    }
    hk_toggle := new_toggle
    hk_lock := new_lock
    hk_close := new_close
    if (hk_toggle != "") {
        try {
            Hotkey(hk_toggle, DoMinimize)
        } catch {
        }
    }
    if (hk_lock != "") {
        try {
            Hotkey(hk_lock, DoRestore)
        } catch {
        }
    }
    if (hk_close != "") {
        try {
            Hotkey(hk_close, DoClose)
        } catch {
        }
    }
}

CheckCmd(*) {
    global cmd_file, panel_title, browser_title
    if (cmd_file = "")
        return
    if !FileExist(cmd_file)
        return
    try cmd := FileRead(cmd_file)
    catch
        return
    if (cmd = "")
        return
    try FileDelete(cmd_file)
    parts := StrSplit(cmd, "|")
    if (parts.Length < 1)
        return
    if (parts[1] = "top_on")
        ApplyTopmost("on")
    else if (parts[1] = "top_off")
        ApplyTopmost("off")
    else if (parts[1] = "top_toggle")
        ApplyTopmost("toggle")
    else if (parts[1] = "update") {
        if (parts.Length >= 2)
            panel_title := parts[2]
        if (parts.Length >= 3)
            browser_title := parts[3]
        if (parts.Length >= 6)
            UpdateHotkeys(parts[4], parts[5], parts[6])
    }
}
"""
    script = script_v2 if use_v2 else script_v1
    try:
        with open(path, "r", encoding="utf-8") as f:
            cur = f.read()
        if cur == script:
            return
    except Exception:
        pass
    with open(path, "w", encoding="utf-8") as f:
        f.write(script)

def make_extra_profile_dir():
    tag = f"{os.getpid()}_{int(time.time() * 1000)}"
    return os.path.join(APP_DIR, f"_mini_fish_profile_extra_{tag}")

ICON_URLS = {
    "chrome": "https://commons.wikimedia.org/wiki/Special:FilePath/Google_Chrome_icon_(February_2022).svg?width=96",
    "photoshop": "https://commons.wikimedia.org/wiki/Special:FilePath/Adobe_Photoshop_CC_icon.svg?width=96",
    "3dsmax": "https://img.icons8.com/color/96/autodesk-3ds-max.png",
    "wps": "https://www.google.com/s2/favicons?domain=wps.cn&sz=128",
    "baidunetdisk": "https://www.google.com/s2/favicons?domain=pan.baidu.com&sz=128",
    "browser360": "https://www.google.com/s2/favicons?domain=browser.360.cn&sz=128",
}

ICON_FILES = {
    name: os.path.join(ICON_DIR, f"{name}.png") for name in ICON_URLS
}
ICON_FILES.update(LOCAL_ICON_FILES)
ICON_META_PATH = os.path.join(ICON_DIR, "_icon_meta.json")

GENERIC_ICON_STYLES = {"globe", "video", "chat", "folder", "star"}
BASE_ICON_CHOICES = ["chrome", "photoshop", "3dsmax", "wps", "baidunetdisk", "browser360"]
HIDDEN_ICON_KEYS = {"PS"}
EXTRA_ICON_CHOICES = [k for k in LOCAL_ICON_CHOICES if k not in BASE_ICON_CHOICES and k not in HIDDEN_ICON_KEYS]
EXTRA_ICON_CHOICES.sort()
PANEL_ICON_CHOICES = BASE_ICON_CHOICES + EXTRA_ICON_CHOICES + ["globe", "video", "chat", "folder", "star", "custom"]
BROWSER_ICON_CHOICES = ["site"] + BASE_ICON_CHOICES + EXTRA_ICON_CHOICES + ["custom"]
ICON_DISPLAY_NAMES = {
    "chrome": "Chrome",
    "photoshop": "Photoshop",
    "3dsmax": "3DsMax",
    "wps": "WPS",
    "baidunetdisk": "百度网盘",
    "browser360": "360浏览器",
    "site": "网站图标",
    "globe": "地球",
    "video": "视频",
    "chat": "聊天",
    "folder": "文件夹",
    "star": "星标",
    "custom": "自定义",
    "Acrobat DC": "Acrobat DC",
    "Ae": "After Effects",
    "Ai": "Illustrator",
    "Ai-chatgpt": "ChatGPT",
    "An": "Animate",
    "Au": "Audition",
    "blender": "Blender",
    "C4DICON": "Cinema 4D",
    "CAD": "AutoCAD",
    "Lr": "Lightroom",
    "Octane": "Octane",
    "Rhinoceros": "Rhinoceros",
    "Topaz Video AI": "Topaz Video AI",
    "Visual Studio": "Visual Studio",
    "vmware": "VMware",
    "VPN": "VPN",
    "Win search": "搜索",
    "XMind": "XMind",
    "剪映": "剪映",
    "浏览器": "浏览器",
    "微软1": "微软",
    "PS": "Photoshop",
    "PH": "Pornhub",
}

def icon_display_name(style: str) -> str:
    return ICON_DISPLAY_NAMES.get(style, style)

SPONSOR_TEXT = "本脚本免费使用，如果对你有帮助，欢迎随意支持作者，谢谢！"
SPONSOR_QR_FILES = [
    ("支付宝", os.path.join(ASSETS_DIR, "支付宝收款10000元.jpg")),
    ("微信", os.path.join(ASSETS_DIR, "微信收款10000元.jpg")),
]
POOR_TEXT = "识相点哈，别给我宝宝一杯蜜雪冰城都点不了！"
POOR_QR_FILES = [
    ("支付宝", os.path.join(ASSETS_DIR, "支付宝收款.jpg")),
    ("微信", os.path.join(ASSETS_DIR, "微信收款.jpg")),
]

RATIO_LABELS = [
    "1:1",
    "16:9 横",
    "9:16 竖",
    "4:5 竖",
    "5:4 横",
    "5:7 竖",
    "7:5 横",
    "3:4 竖",
    "4:3 横",
    "3:5 竖",
    "5:3 横",
    "2:3 竖",
    "3:2 横",
]
RATIO_LABEL_TO_KEY = {
    "1:1": "1:1",
    "16:9 横": "16:9",
    "9:16 竖": "9:16",
    "4:5 竖": "4:5",
    "5:4 横": "5:4",
    "5:7 竖": "5:7",
    "7:5 横": "7:5",
    "3:4 竖": "3:4",
    "4:3 横": "4:3",
    "3:5 竖": "3:5",
    "5:3 横": "5:3",
    "2:3 竖": "2:3",
    "3:2 横": "3:2",
}
RATIO_KEY_TO_LABEL = {v: k for k, v in RATIO_LABEL_TO_KEY.items()}
SIZE_LEVEL_LABEL = {"S": "小", "M": "中", "L": "大"}
WINDOW_SIZE_PRESETS = {
    "1:1": {"S": (480, 480), "M": (640, 640), "L": (800, 800)},
    "16:9": {"S": (640, 360), "M": (960, 540), "L": (1280, 720)},
    "9:16": {"S": (360, 640), "M": (540, 960), "L": (720, 1280)},
    "4:5": {"S": (480, 600), "M": (640, 800), "L": (800, 1000)},
    "5:4": {"S": (600, 480), "M": (800, 640), "L": (1000, 800)},
    "5:7": {"S": (500, 700), "M": (600, 840), "L": (700, 980)},
    "7:5": {"S": (700, 500), "M": (840, 600), "L": (980, 700)},
    "3:4": {"S": (480, 640), "M": (600, 800), "L": (720, 960)},
    "4:3": {"S": (640, 480), "M": (800, 600), "L": (960, 720)},
    "3:5": {"S": (360, 600), "M": (480, 800), "L": (600, 1000)},
    "5:3": {"S": (600, 360), "M": (800, 480), "L": (1000, 600)},
    "2:3": {"S": (400, 600), "M": (480, 720), "L": (640, 960)},
    "3:2": {"S": (600, 400), "M": (720, 480), "L": (960, 640)},
}
BROWSER_SCALE_MIN = 0.2
BROWSER_SCALE_MAX = 1.6
BROWSER_POS_MARGIN = 20
ATTACH_MARGIN = 8
BROWSER_POS_LABELS = [
    "右下角",
    "左下角",
    "右上角",
    "左上角",
    "中心",
    "右中",
    "左中",
    "上中",
    "下中",
]
BROWSER_POS_LABEL_TO_KEY = {
    "右下角": "bottom_right",
    "左下角": "bottom_left",
    "右上角": "top_right",
    "左上角": "top_left",
    "中心": "center",
    "右中": "right_center",
    "左中": "left_center",
    "上中": "top_center",
    "下中": "bottom_center",
}
BROWSER_POS_KEY_TO_LABEL = {v: k for k, v in BROWSER_POS_LABEL_TO_KEY.items()}

HOTKEY_DEFAULT_TOGGLE_OLD = "Ctrl+Shift+Alt+0"
HOTKEY_DEFAULT_LOCK_OLD = "Ctrl+Shift+Alt+."
HOTKEY_DEFAULT_TOGGLE = "Ctrl+Win+Alt+0"
HOTKEY_DEFAULT_LOCK = "Ctrl+Win+Alt+."
HOTKEY_DEFAULT_CLOSE = "Ctrl+Shift+Win+0"
HOTKEY_BROWSER_REFRESH = "Ctrl+R"
HOTKEY_BROWSER_BACK = "Alt+Left"
HOTKEY_PAGE_ZOOM_IN = "Ctrl+="
HOTKEY_PAGE_ZOOM_OUT = "Ctrl+-"
HOTKEY_TOGGLE_MUTE = "Ctrl+Shift+M"
HOTKEY_WINDOW_SCALE_UP = "Ctrl+Shift+="
HOTKEY_WINDOW_SCALE_DOWN = "Ctrl+Shift+-"

def launch_chrome_app(chrome_exe: str, url: str, w: int, h: int, x: int, y: int, port: int,
                      profile_dir: str = "", audio_enabled: bool = True):
    profile_dir = profile_dir or PROFILE_DIR
    os.makedirs(profile_dir, exist_ok=True)
    args = [
        chrome_exe,
        f"--app={url}",
        f"--window-size={w},{h}",
        f"--window-position={x},{y}",
        f"--remote-debugging-port={port}",
        f"--user-data-dir={profile_dir}",
        "--no-first-run",
        "--no-default-browser-check",
        "--disable-infobars",
    ]
    if not audio_enabled:
        args.append("--mute-audio")
    return subprocess.Popen(args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

def is_debug_port_ready(port: int) -> bool:
    if not port:
        return False
    url = f"http://127.0.0.1:{port}/json/version"
    try:
        with urllib.request.urlopen(url, timeout=0.5) as resp:
            _ = resp.read(1)
        return True
    except Exception:
        return False

def attach_selenium(port: int):
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    opts = Options()
    opts.add_experimental_option("debuggerAddress", f"127.0.0.1:{port}")
    return webdriver.Chrome(options=opts)

def launch_browser_session(chrome_exe: str, url: str, w: int, h: int, x: int, y: int,
                           profile_dir: str, audio_enabled: bool, attempts: int = 2):
    driver = None
    proc = None
    port = None
    kill_browsers_by_profile(profile_dir)
    for i in range(attempts):
        attempt_start = time.time()
        port = pick_free_port()
        proc = launch_chrome_app(chrome_exe, url, w, h, x, y, port, profile_dir, audio_enabled)
        t0 = time.time()
        while time.time() - t0 < 12:
            try:
                if not is_debug_port_ready(port):
                    time.sleep(0.25)
                    continue
                driver = attach_selenium(port)
                break
            except Exception:
                time.sleep(0.25)
        if not driver:
            alt_port = read_devtools_port(profile_dir, min_mtime=attempt_start - 1)
            if alt_port and alt_port != port and is_debug_port_ready(alt_port):
                try:
                    driver = attach_selenium(alt_port)
                    port = alt_port
                except Exception:
                    driver = None
        if driver:
            return proc, driver, port, ""
        if i < attempts - 1:
            try:
                kill_pid_tree(getattr(proc, "pid", 0))
                kill_browsers_by_profile(profile_dir)
            except Exception:
                pass
            proc = None
            port = None
            time.sleep(0.4)
    try:
        if proc:
            kill_pid_tree(getattr(proc, "pid", 0))
    except Exception:
        pass
    kill_browsers_by_profile(profile_dir)
    return None, None, None, "浏览器未就绪，请稍后重试"

def kill_pid_tree(pid: int):
    if not pid:
        return
    try:
        subprocess.run(
            ["taskkill", "/PID", str(int(pid)), "/T", "/F"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=CREATE_NO_WINDOW,
        )
    except Exception:
        pass

def kill_browsers_by_profile(profile_dir: str):
    if not profile_dir:
        return
    try:
        pattern = profile_dir.replace("'", "''")
        cmd = (
            "Get-CimInstance Win32_Process | "
            "Where-Object { $_.Name -in @('chrome.exe','msedge.exe') -and $_.CommandLine "
            "-and $_.CommandLine -like '*--user-data-dir=*"
            + pattern
            + "*' } | "
            "ForEach-Object { Stop-Process -Id $_.ProcessId -Force }"
        )
        subprocess.run(
            ["powershell", "-NoProfile", "-Command", cmd],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=CREATE_NO_WINDOW,
        )
    except Exception:
        pass

# -------------------- Windows HWND helpers --------------------
user32 = ctypes.windll.user32
EnumWindows = user32.EnumWindows
EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
GetWindowThreadProcessId = user32.GetWindowThreadProcessId
IsWindowVisible = user32.IsWindowVisible
GetClassNameW = user32.GetClassNameW
GetWindowRect = user32.GetWindowRect
GetWindowLongW = user32.GetWindowLongW
SetWindowLongW = user32.SetWindowLongW
SetWindowLongPtrW = getattr(user32, "SetWindowLongPtrW", None)
GetWindowLongPtrW = getattr(user32, "GetWindowLongPtrW", None)
SetLayeredWindowAttributes = user32.SetLayeredWindowAttributes
SetWindowPos = user32.SetWindowPos
ShowWindow = user32.ShowWindow
SetWindowTextW = user32.SetWindowTextW
GetWindowTextW = user32.GetWindowTextW
GetWindowTextLengthW = user32.GetWindowTextLengthW
GetForegroundWindow = user32.GetForegroundWindow
RegisterHotKey = user32.RegisterHotKey
UnregisterHotKey = user32.UnregisterHotKey
PeekMessageW = user32.PeekMessageW
MessageBoxW = user32.MessageBoxW
kernel32 = ctypes.windll.kernel32
CreateToolhelp32Snapshot = kernel32.CreateToolhelp32Snapshot
Process32FirstW = kernel32.Process32FirstW
Process32NextW = kernel32.Process32NextW
CloseHandle = kernel32.CloseHandle
GetConsoleWindow = kernel32.GetConsoleWindow
CreateMutexW = kernel32.CreateMutexW
GetLastError = kernel32.GetLastError
iphlpapi = ctypes.WinDLL("Iphlpapi.dll")
GetExtendedTcpTable = iphlpapi.GetExtendedTcpTable

GWL_EXSTYLE = -20
GWL_HWNDPARENT = -8
WS_EX_LAYERED = 0x00080000
LWA_ALPHA = 0x00000002

HWND_TOPMOST = -1
HWND_NOTOPMOST = -2
HWND_TOP = 0
SWP_NOSIZE = 0x0001
SWP_NOMOVE = 0x0002
SWP_NOZORDER = 0x0004
SWP_NOACTIVATE = 0x0010
SWP_SHOWWINDOW = 0x0040
ERROR_ALREADY_EXISTS = 183

SW_HIDE = 0
SW_MINIMIZE = 6
SW_RESTORE = 9

WM_HOTKEY = 0x0312
PM_REMOVE = 0x0001

MOD_ALT = 0x0001
MOD_CONTROL = 0x0002
MOD_SHIFT = 0x0004
MOD_WIN = 0x0008
MOD_NOREPEAT = 0x4000

TH32CS_SNAPPROCESS = 0x00000002
INVALID_HANDLE_VALUE = ctypes.c_void_p(-1).value

CHROME_WINDOW_CLASSES = {"Chrome_WidgetWin_0", "Chrome_WidgetWin_1", "Chrome_WidgetWin_2"}
CHROME_WINDOW_CLASS_PREFIX = "Chrome_WidgetWin_"

def is_chrome_window_class(name: str) -> bool:
    if not name:
        return False
    return name in CHROME_WINDOW_CLASSES or name.startswith(CHROME_WINDOW_CLASS_PREFIX)

class RECT(ctypes.Structure):
    _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long), ("right", ctypes.c_long), ("bottom", ctypes.c_long)]

class PROCESSENTRY32(ctypes.Structure):
    _fields_ = [
        ("dwSize", ctypes.c_uint32),
        ("cntUsage", ctypes.c_uint32),
        ("th32ProcessID", ctypes.c_uint32),
        ("th32DefaultHeapID", ctypes.c_void_p),
        ("th32ModuleID", ctypes.c_uint32),
        ("cntThreads", ctypes.c_uint32),
        ("th32ParentProcessID", ctypes.c_uint32),
        ("pcPriClassBase", ctypes.c_long),
        ("dwFlags", ctypes.c_uint32),
        ("szExeFile", ctypes.c_wchar * 260),
    ]

GetWindowTextW.argtypes = [ctypes.c_void_p, ctypes.c_wchar_p, ctypes.c_int]
GetWindowTextW.restype = ctypes.c_int
GetWindowTextLengthW.argtypes = [ctypes.c_void_p]
GetWindowTextLengthW.restype = ctypes.c_int
SetWindowTextW.argtypes = [ctypes.c_void_p, ctypes.c_wchar_p]
SetWindowTextW.restype = ctypes.c_bool
GetForegroundWindow.argtypes = []
GetForegroundWindow.restype = ctypes.c_void_p
RegisterHotKey.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_uint, ctypes.c_uint]
RegisterHotKey.restype = ctypes.c_bool
UnregisterHotKey.argtypes = [ctypes.c_void_p, ctypes.c_int]
UnregisterHotKey.restype = ctypes.c_bool
PeekMessageW.argtypes = [ctypes.POINTER(wintypes.MSG), ctypes.c_void_p, ctypes.c_uint, ctypes.c_uint, ctypes.c_uint]
PeekMessageW.restype = ctypes.c_bool
CreateToolhelp32Snapshot.argtypes = [ctypes.c_uint32, ctypes.c_uint32]
CreateToolhelp32Snapshot.restype = ctypes.c_void_p
Process32FirstW.argtypes = [ctypes.c_void_p, ctypes.POINTER(PROCESSENTRY32)]
Process32FirstW.restype = ctypes.c_bool
Process32NextW.argtypes = [ctypes.c_void_p, ctypes.POINTER(PROCESSENTRY32)]
Process32NextW.restype = ctypes.c_bool
CloseHandle.argtypes = [ctypes.c_void_p]
CloseHandle.restype = ctypes.c_bool
GetConsoleWindow.argtypes = []
GetConsoleWindow.restype = ctypes.c_void_p
GetExtendedTcpTable.argtypes = [ctypes.c_void_p, ctypes.POINTER(wintypes.DWORD), wintypes.BOOL, wintypes.ULONG, wintypes.ULONG, wintypes.ULONG]
GetExtendedTcpTable.restype = wintypes.DWORD

AF_INET = 2
TCP_TABLE_OWNER_PID_ALL = 5
MIB_TCP_STATE_LISTEN = 2

class MIB_TCPROW_OWNER_PID(ctypes.Structure):
    _fields_ = [
        ("state", wintypes.DWORD),
        ("localAddr", wintypes.DWORD),
        ("localPort", wintypes.DWORD),
        ("remoteAddr", wintypes.DWORD),
        ("remotePort", wintypes.DWORD),
        ("owningPid", wintypes.DWORD),
    ]

class MIB_TCPTABLE_OWNER_PID(ctypes.Structure):
    _fields_ = [("dwNumEntries", wintypes.DWORD), ("table", MIB_TCPROW_OWNER_PID * 1)]

def get_window_text(hwnd):
    if not hwnd:
        return ""
    try:
        length = GetWindowTextLengthW(hwnd)
        if length <= 0:
            return ""
        buf = ctypes.create_unicode_buffer(length + 1)
        GetWindowTextW(hwnd, buf, length + 1)
        return buf.value
    except Exception:
        return ""

def _parse_local_port(addr: str) -> int:
    if not addr:
        return 0
    try:
        if addr.startswith("[") and "]" in addr:
            return int(addr.rsplit("]:", 1)[-1])
        if ":" in addr:
            return int(addr.rsplit(":", 1)[-1])
    except Exception:
        return 0
    return 0

def _get_pid_by_port_api(port: int) -> int:
    size = wintypes.DWORD(0)
    res = GetExtendedTcpTable(None, ctypes.byref(size), False, AF_INET, TCP_TABLE_OWNER_PID_ALL, 0)
    if res not in (0, 122):
        return 0
    buf = ctypes.create_string_buffer(size.value)
    res = GetExtendedTcpTable(buf, ctypes.byref(size), False, AF_INET, TCP_TABLE_OWNER_PID_ALL, 0)
    if res != 0:
        return 0
    table = ctypes.cast(buf, ctypes.POINTER(MIB_TCPTABLE_OWNER_PID)).contents
    count = int(table.dwNumEntries)
    if count <= 0:
        return 0
    rows_type = MIB_TCPROW_OWNER_PID * count
    rows = ctypes.cast(ctypes.byref(table.table), ctypes.POINTER(rows_type)).contents
    for row in rows:
        if int(row.state) != MIB_TCP_STATE_LISTEN:
            continue
        try:
            p = socket.ntohs(int(row.localPort) & 0xFFFF)
        except Exception:
            continue
        if p == int(port):
            return int(row.owningPid)
    return 0

def get_pid_by_port(port: int) -> int:
    if not port:
        return 0
    try:
        pid = _get_pid_by_port_api(port)
        if pid:
            return pid
    except Exception:
        pass
    if getattr(sys, "frozen", False):
        return 0
    try:
        out = subprocess.check_output(
            ["netstat", "-ano", "-p", "tcp"],
            stderr=subprocess.DEVNULL,
            creationflags=CREATE_NO_WINDOW,
            text=True,
            encoding="utf-8",
            errors="ignore",
        )
    except Exception:
        return 0
    for line in out.splitlines():
        line = line.strip()
        if not line or not line.lower().startswith("tcp"):
            continue
        parts = line.split()
        if len(parts) < 5:
            continue
        local = parts[1]
        state = parts[3]
        pid_str = parts[4]
        if state.upper() != "LISTENING":
            continue
        if _parse_local_port(local) != int(port):
            continue
        try:
            return int(pid_str)
        except Exception:
            return 0
    return 0

def get_related_pids(root_pid: int):
    # Include descendants in case the launcher PID doesn't own the window.
    if not root_pid:
        return set()
    snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    if snapshot == INVALID_HANDLE_VALUE:
        return {int(root_pid)}
    entry = PROCESSENTRY32()
    entry.dwSize = ctypes.sizeof(PROCESSENTRY32)
    ok = Process32FirstW(snapshot, ctypes.byref(entry))
    if not ok:
        CloseHandle(snapshot)
        return {int(root_pid)}

    ppid_map = {}
    while ok:
        ppid_map[int(entry.th32ProcessID)] = int(entry.th32ParentProcessID)
        ok = Process32NextW(snapshot, ctypes.byref(entry))
    CloseHandle(snapshot)

    related = {int(root_pid)}
    stack = [int(root_pid)]
    while stack:
        cur = stack.pop()
        for pid, ppid in ppid_map.items():
            if ppid == cur and pid not in related:
                related.add(pid)
                stack.append(pid)
    return related

def get_pid_hwnds(pid_or_pids):
    if not pid_or_pids:
        return []
    if isinstance(pid_or_pids, (set, list, tuple)):
        pids = {int(p) for p in pid_or_pids if p}
    else:
        pids = {int(pid_or_pids)}
    hwnds = []
    def callback(hwnd, lParam):
        if not IsWindowVisible(hwnd):
            return True
        _pid = ctypes.c_ulong()
        GetWindowThreadProcessId(hwnd, ctypes.byref(_pid))
        if _pid.value not in pids:
            return True
        buf = ctypes.create_unicode_buffer(256)
        GetClassNameW(hwnd, buf, 256)
        if is_chrome_window_class(buf.value):
            hwnds.append(hwnd)
        return True
    EnumWindows(EnumWindowsProc(callback), 0)
    return hwnds

def get_chrome_hwnds():
    hwnds = []
    def callback(hwnd, lParam):
        if not IsWindowVisible(hwnd):
            return True
        buf = ctypes.create_unicode_buffer(256)
        GetClassNameW(hwnd, buf, 256)
        if is_chrome_window_class(buf.value):
            hwnds.append(hwnd)
        return True
    EnumWindows(EnumWindowsProc(callback), 0)
    return hwnds

def find_chrome_hwnds_by_title(title: str, exact: bool = True, include_hidden: bool = True):
    title_l = (title or "").strip().lower()
    if not title_l:
        return []
    hwnds = []
    def callback(hwnd, lParam):
        try:
            if (not include_hidden) and (not IsWindowVisible(hwnd)):
                return True
            buf = ctypes.create_unicode_buffer(256)
            GetClassNameW(hwnd, buf, 256)
            if not is_chrome_window_class(buf.value):
                return True
            t = get_window_text(hwnd).strip().lower()
            if not t:
                return True
            if exact:
                if t == title_l:
                    hwnds.append(hwnd)
            else:
                if title_l in t:
                    hwnds.append(hwnd)
        except Exception:
            pass
        return True
    EnumWindows(EnumWindowsProc(callback), 0)
    return hwnds

def acquire_single_instance() -> bool:
    global INSTANCE_MUTEX_HANDLE
    if INSTANCE_ID:
        return True
    handle = CreateMutexW(None, True, APP_MUTEX_NAME)
    if not handle:
        return True
    if GetLastError() == ERROR_ALREADY_EXISTS:
        try:
            CloseHandle(handle)
        except Exception:
            pass
        return False
    INSTANCE_MUTEX_HANDLE = handle
    return True

def notify_already_running():
    try:
        MessageBoxW(0, "程序已在运行，请从任务栏托盘恢复。", "提示", 0x00000040)
    except Exception:
        pass

def pick_main_hwnd(pid: int, title_hint: str = "", host_hint: str = "", size_hint=None, include_all: bool = True):
    hwnds = get_pid_hwnds(pid)
    if not hwnds and pid:
        try:
            hwnds = get_pid_hwnds(get_related_pids(pid))
        except Exception:
            hwnds = []
    if not hwnds and include_all:
        hwnds = get_chrome_hwnds()
    if not hwnds:
        return None

    title_l = (title_hint or "").lower()
    host_l = (host_hint or "").lower()
    target_area = None
    if size_hint and size_hint[0] and size_hint[1]:
        target_area = int(size_hint[0] * size_hint[1])

    best = None
    best_score = (-1, -1.0, -1)
    for hwnd in hwnds:
        r = RECT()
        if not GetWindowRect(hwnd, ctypes.byref(r)):
            continue
        w = max(0, r.right - r.left)
        h = max(0, r.bottom - r.top)
        area = w * h

        title_score = 0
        text_l = ""
        if title_l or host_l:
            text_l = get_window_text(hwnd).lower()
            if title_l and title_l in text_l:
                title_score += 2
            if host_l and host_l in text_l:
                title_score += 1

        size_score = 0.0
        if target_area and area:
            size_score = min(area, target_area) / max(area, target_area)

        score = (title_score, size_score, area)
        if score > best_score:
            best_score = score
            best = hwnd
    return best

def set_window_alpha(hwnd, alpha_0_1: float):
    if not hwnd:
        return
    a = int(max(0.15, min(1.0, alpha_0_1)) * 255)
    ex = GetWindowLongW(hwnd, GWL_EXSTYLE)
    if not (ex & WS_EX_LAYERED):
        SetWindowLongW(hwnd, GWL_EXSTYLE, ex | WS_EX_LAYERED)
    SetLayeredWindowAttributes(hwnd, 0, a, LWA_ALPHA)

def set_window_topmost(hwnd, topmost: bool, force=False):
    if not hwnd:
        return
    insert_after = HWND_TOPMOST if topmost else HWND_NOTOPMOST
    flags = SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW
    if not force:
        flags |= SWP_NOACTIVATE
    if force and topmost:
        SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, flags)
    SetWindowPos(hwnd, insert_after, 0, 0, 0, 0, flags)

def get_window_rect(hwnd):
    if not hwnd:
        return None
    r = RECT()
    if not GetWindowRect(hwnd, ctypes.byref(r)):
        return None
    return r.left, r.top, r.right, r.bottom

def set_window_owner(hwnd, owner_hwnd):
    if not hwnd:
        return
    try:
        if SetWindowLongPtrW and GetWindowLongPtrW:
            SetWindowLongPtrW(hwnd, GWL_HWNDPARENT, owner_hwnd)
        else:
            SetWindowLongW(hwnd, GWL_HWNDPARENT, owner_hwnd)
    except Exception:
        pass

def minimize_window(hwnd):
    if hwnd:
        ShowWindow(hwnd, SW_MINIMIZE)

def restore_window(hwnd):
    if hwnd:
        ShowWindow(hwnd, SW_RESTORE)

def set_window_title(hwnd, title: str):
    if hwnd and title:
        try:
            SetWindowTextW(hwnd, title)
        except Exception:
            pass

def hide_window(hwnd):
    if hwnd:
        ShowWindow(hwnd, SW_HIDE)

def hide_console_window():
    try:
        hwnd = GetConsoleWindow()
        if hwnd:
            ShowWindow(hwnd, SW_HIDE)
    except Exception:
        pass

# -------------------- Settings --------------------
def load_settings():
    default = {
        "presets": ["https://www.douyin.com", "https://www.bilibili.com"],
        "last_url": "https://www.douyin.com",
        "recent": [],
        "panel_topmost": True,
        "browser_topmost": False,
        "panel_icon_style": "globe",   # built-in style name
        "remember_zoom": 0.85,
        "remember_alpha": 1.0,
        "remember_panel_alpha": 1.0,
        "panel_title": APP_TITLE,
        "browser_title": "mini-browser",
        "panel_custom_icon_path": "",
        "browser_icon_style": "",
        "browser_custom_icon_path": "",
        "custom_icon_path": "",
        "browser_ratio": "4:3",
        "browser_size_level": "S",
        "browser_scale": 1.0,
        "browser_position": "bottom_right",
        "hotkey_toggle": HOTKEY_DEFAULT_TOGGLE,
        "hotkey_lock": HOTKEY_DEFAULT_LOCK,
        "hotkey_close": HOTKEY_DEFAULT_CLOSE,
        "custom_status": "",
        "attach_enabled": False,
        "attach_side": "left",
        "merge_taskbar": False,
        "audio_enabled": True,
        "sync_taskbar_icon": True,
        "panel_tray_enabled": False,
        "browser_tray_enabled": False,
        "update_source": DEFAULT_UPDATE_SOURCE,
    }
    try:
        with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict):
            default.update(data)
    except Exception:
        pass

    if default.get("panel_icon_style") in ("PS", "PH"):
        default["panel_icon_style"] = "photoshop"
    if default.get("browser_icon_style") in ("PS", "PH"):
        default["browser_icon_style"] = "photoshop"

    if not default.get("browser_icon_style"):
        panel_style = default.get("panel_icon_style", "globe")
        if panel_style in ICON_FILES or panel_style == "custom":
            default["browser_icon_style"] = panel_style
        else:
            default["browser_icon_style"] = "site"

    legacy_custom = (default.get("custom_icon_path") or "").strip()
    if legacy_custom:
        if not default.get("panel_custom_icon_path"):
            default["panel_custom_icon_path"] = legacy_custom
        if not default.get("browser_custom_icon_path"):
            default["browser_custom_icon_path"] = legacy_custom
        if default.get("browser_icon_style") in ("", None):
            default["browser_icon_style"] = "custom"
        if default.get("panel_icon_style") in ("", None):
            default["panel_icon_style"] = "custom"

    if default.get("browser_icon_style") not in BROWSER_ICON_CHOICES:
        default["browser_icon_style"] = "site"

    if default.get("browser_ratio") not in WINDOW_SIZE_PRESETS:
        default["browser_ratio"] = "4:3"
    if default.get("browser_size_level") not in ("S", "M", "L"):
        default["browser_size_level"] = "S"
    pos = default.get("browser_position", "bottom_right")
    if pos in BROWSER_POS_LABEL_TO_KEY:
        pos = BROWSER_POS_LABEL_TO_KEY[pos]
    if pos not in BROWSER_POS_KEY_TO_LABEL:
        pos = "bottom_right"
    default["browser_position"] = pos
    try:
        scale = float(default.get("browser_scale", 1.0))
    except Exception:
        scale = 1.0
    default["browser_scale"] = max(BROWSER_SCALE_MIN, min(BROWSER_SCALE_MAX, scale))
    if default.get("hotkey_toggle") == HOTKEY_DEFAULT_TOGGLE_OLD:
        default["hotkey_toggle"] = HOTKEY_DEFAULT_TOGGLE
    if default.get("hotkey_lock") == HOTKEY_DEFAULT_LOCK_OLD:
        default["hotkey_lock"] = HOTKEY_DEFAULT_LOCK
    if default.get("hotkey_toggle") == "Ctrl+Win+Alt+T" and default.get("hotkey_lock") == "Ctrl+Shift+Win+T":
        default["hotkey_toggle"] = HOTKEY_DEFAULT_TOGGLE
        default["hotkey_lock"] = HOTKEY_DEFAULT_LOCK
    if not default.get("hotkey_toggle"):
        default["hotkey_toggle"] = HOTKEY_DEFAULT_TOGGLE
    if not default.get("hotkey_lock"):
        default["hotkey_lock"] = HOTKEY_DEFAULT_LOCK
    if not default.get("hotkey_close"):
        default["hotkey_close"] = HOTKEY_DEFAULT_CLOSE
    if default.get("attach_side") not in ("left", "right"):
        default["attach_side"] = "left"
    if default.get("update_source") not in UPDATE_SOURCE_LABELS:
        default["update_source"] = DEFAULT_UPDATE_SOURCE

    # normalize & dedupe
    def dedupe(urls):
        seen = set()
        out = []
        for u in urls:
            u = normalize_url(u)
            if u not in seen:
                out.append(u)
                seen.add(u)
        return out

    default["presets"] = dedupe(default.get("presets", []))
    default["recent"] = dedupe(default.get("recent", []))[:30]
    default["last_url"] = normalize_url(default.get("last_url", "https://www.douyin.com"))
    return default

def is_frozen_app() -> bool:
    return IS_FROZEN

def get_app_dir() -> str:
    return APP_DIR

def get_settings_path(base_dir: str = None) -> str:
    base_dir = base_dir or APP_DIR
    return os.path.join(base_dir, SETTINGS_FILENAME)

def get_settings_bootstrap_path(base_dir: str = None) -> str:
    base_dir = base_dir or APP_DIR
    return os.path.join(base_dir, SETTINGS_BOOTSTRAP_FLAG)

def get_first_run_flag_path(base_dir: str = None) -> str:
    base_dir = base_dir or APP_DIR
    return os.path.join(base_dir, FIRST_RUN_FLAG)

def get_update_cache_dir() -> str:
    return os.path.join(get_app_dir(), "_mini_fish_update")

def _parse_version(text: str):
    parts = [int(x) for x in re.findall(r"\d+", text or "")]
    return tuple(parts)

def _normalize_version(parts, length: int = 4):
    parts = list(parts[:length])
    while len(parts) < length:
        parts.append(0)
    return tuple(parts)

def _version_gt(a, b) -> bool:
    return _normalize_version(a) > _normalize_version(b)

def _get_env_token(*names) -> str:
    for name in names:
        value = os.environ.get(name, "").strip()
        if value:
            return value
    return ""

def _append_query_param(url: str, key: str, value: str) -> str:
    if not value:
        return url
    try:
        parsed = urllib.parse.urlparse(url)
        query = urllib.parse.parse_qs(parsed.query, keep_blank_values=True)
        if key in query:
            return url
        query[key] = [value]
        new_query = urllib.parse.urlencode(query, doseq=True)
        return urllib.parse.urlunparse(parsed._replace(query=new_query))
    except Exception:
        return url

def _http_get_json(url: str, timeout: int = UPDATE_CHECK_TIMEOUT, headers: dict = None):
    req_headers = {"User-Agent": "Mozilla/5.0", "Accept": "application/json"}
    if headers:
        req_headers.update(headers)
    req = urllib.request.Request(url, headers=req_headers)
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        data = resp.read()
    return json.loads(data.decode("utf-8"))

def _pick_update_asset(assets):
    if not assets:
        return None
    candidates = []
    for asset in assets:
        name = (asset.get("name") or "").lower()
        if not name.endswith("_package.zip"):
            continue
        score = 0
        if UPDATE_PRODUCT_KEY in name:
            score += 2
        candidates.append((score, asset))
    if not candidates:
        return None
    candidates.sort(key=lambda x: x[0], reverse=True)
    return candidates[0][1]

def check_gitee_update(current_version: str):
    token = _get_env_token("GITEE_TOKEN", "GITEE_ACCESS_TOKEN")
    try:
        cur_ver = _parse_version(current_version)
        if not cur_ver:
            return None, "当前版本号无效"
        releases_url = f"{GITEE_REPO_API}/releases?per_page=20&page=1"
        releases_url = _append_query_param(releases_url, "access_token", token)
        releases = _http_get_json(releases_url, timeout=UPDATE_CHECK_TIMEOUT)
        latest = None
        for rel in releases:
            tag = rel.get("tag_name") or ""
            ver = _parse_version(tag)
            if not ver:
                continue
            if latest is None or _version_gt(ver, latest["version_tuple"]):
                latest = {
                    "version_tuple": ver,
                    "tag": tag,
                    "name": rel.get("name") or tag,
                    "body": rel.get("body") or "",
                    "release_id": rel.get("id"),
                }
        if not latest or not _version_gt(latest["version_tuple"], cur_ver):
            return None, ""
        assets_url = f"{GITEE_REPO_API}/releases/{latest['release_id']}/attach_files"
        assets_url = _append_query_param(assets_url, "access_token", token)
        assets = _http_get_json(assets_url, timeout=UPDATE_CHECK_TIMEOUT)
        asset = _pick_update_asset(assets)
        if not asset:
            return None, "未找到更新包附件"
        version_str = ".".join(str(x) for x in latest["version_tuple"])
        info = {
            "version": version_str,
            "tag": latest["tag"],
            "name": latest["name"],
            "body": latest["body"],
            "release_id": latest["release_id"],
            "asset_name": asset.get("name") or "",
            "asset_url": asset.get("browser_download_url") or "",
            "asset_size": int(asset.get("size") or 0),
            "source": UPDATE_SOURCE_GITEE,
        }
        if not info["asset_url"] or not info["asset_name"]:
            return None, "更新包链接无效"
        return info, ""
    except urllib.error.HTTPError as e:
        if e.code == 403:
            if token:
                return None, "Gitee API 403（可能触发限流或权限不足）"
            return None, "Gitee API 403（可能触发限流），请设置环境变量 GITEE_TOKEN 后重试"
        return None, f"Gitee API HTTP {e.code}: {e.reason}"
    except Exception as e:
        return None, str(e)

def check_github_update(current_version: str):
    token = _get_env_token("GITHUB_TOKEN", "GH_TOKEN")
    try:
        cur_ver = _parse_version(current_version)
        if not cur_ver:
            return None, "当前版本号无效"
        releases_url = f"{GITHUB_REPO_API}/releases?per_page=20&page=1"
        headers = {"Accept": "application/vnd.github+json"}
        if token:
            headers["Authorization"] = f"Bearer {token}"
            headers["X-GitHub-Api-Version"] = "2022-11-28"
        releases = _http_get_json(releases_url, timeout=UPDATE_CHECK_TIMEOUT, headers=headers)
        latest = None
        for rel in releases:
            if rel.get("draft") or rel.get("prerelease"):
                continue
            tag = rel.get("tag_name") or ""
            ver = _parse_version(tag)
            if not ver:
                continue
            if latest is None or _version_gt(ver, latest["version_tuple"]):
                latest = {
                    "version_tuple": ver,
                    "tag": tag,
                    "name": rel.get("name") or tag,
                    "body": rel.get("body") or "",
                    "assets": rel.get("assets") or [],
                }
        if not latest or not _version_gt(latest["version_tuple"], cur_ver):
            return None, ""
        asset = _pick_update_asset(latest["assets"])
        if not asset:
            return None, "未找到更新包附件"
        version_str = ".".join(str(x) for x in latest["version_tuple"])
        info = {
            "version": version_str,
            "tag": latest["tag"],
            "name": latest["name"],
            "body": latest["body"],
            "release_id": 0,
            "asset_name": asset.get("name") or "",
            "asset_url": asset.get("browser_download_url") or "",
            "asset_size": int(asset.get("size") or 0),
            "source": UPDATE_SOURCE_GITHUB,
        }
        if not info["asset_url"] or not info["asset_name"]:
            return None, "更新包链接无效"
        return info, ""
    except urllib.error.HTTPError as e:
        if e.code == 403:
            if token:
                return None, "GitHub API 403（可能触发限流或权限不足）"
            return None, "GitHub API 403（可能触发限流），请设置环境变量 GITHUB_TOKEN 后重试"
        return None, f"GitHub API HTTP {e.code}: {e.reason}"
    except Exception as e:
        return None, str(e)

def check_update_by_source(current_version: str, update_source: str):
    source = update_source or DEFAULT_UPDATE_SOURCE
    if source == UPDATE_SOURCE_GITEE:
        return check_gitee_update(current_version)
    if source == UPDATE_SOURCE_GITHUB:
        return check_github_update(current_version)
    info, err = check_gitee_update(current_version)
    if info:
        return info, ""
    info2, err2 = check_github_update(current_version)
    if info2:
        return info2, ""
    if err and err2:
        return None, f"{err}; {err2}"
    return None, ""

def _download_file(url: str, dest: str, expected_size: int = 0, progress_cb=None):
    os.makedirs(os.path.dirname(dest), exist_ok=True)
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=UPDATE_DOWNLOAD_TIMEOUT) as resp:
        total = int(resp.headers.get("Content-Length") or 0) or int(expected_size or 0)
        downloaded = 0
        with open(dest, "wb") as f:
            while True:
                chunk = resp.read(UPDATE_CHUNK_SIZE)
                if not chunk:
                    break
                f.write(chunk)
                downloaded += len(chunk)
                if progress_cb:
                    progress_cb(downloaded, total)
    size = os.path.getsize(dest)
    if expected_size and size != expected_size:
        raise RuntimeError("更新包大小校验失败")
    if total and size != total:
        raise RuntimeError("更新包下载不完整")

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
            if "牛马神器" in exe_name or "niuma" in exe_name.lower():
                return root, exe_name
        return root, exe_candidates[0]
    return "", ""

def download_update_package(info: dict, progress_cb=None):
    try:
        tag = info.get("tag") or "update"
        asset_name = info.get("asset_name") or "update_package.zip"
        url = info.get("asset_url") or ""
        if info.get("source") == UPDATE_SOURCE_GITEE:
            token = _get_env_token("GITEE_TOKEN", "GITEE_ACCESS_TOKEN")
            url = _append_query_param(url, "access_token", token)
        if not url:
            return None, "更新包链接无效"
        base_dir = os.path.join(get_update_cache_dir(), tag)
        zip_path = os.path.join(base_dir, asset_name)
        _download_file(
            url,
            zip_path,
            expected_size=int(info.get("asset_size") or 0),
            progress_cb=progress_cb,
        )
        extract_dir = os.path.join(base_dir, "extracted")
        if os.path.isdir(extract_dir):
            shutil.rmtree(extract_dir, ignore_errors=True)
        os.makedirs(extract_dir, exist_ok=True)
        with zipfile.ZipFile(zip_path, "r") as zf:
            zf.extractall(extract_dir)
        pkg_root, exe_name = _find_package_root(extract_dir)
        if not pkg_root or not exe_name:
            return None, "更新包解压失败"
        result = {
            "package_root": pkg_root,
            "exe_name": exe_name,
            "version": info.get("version") or "",
        }
        return result, ""
    except Exception as e:
        return None, str(e)

def download_default_settings(update_source: str, target_dir: str = None):
    urls = []
    source = update_source or DEFAULT_UPDATE_SOURCE
    if source == UPDATE_SOURCE_GITEE:
        urls = [DEFAULT_SETTINGS_URL_GITEE]
    elif source == UPDATE_SOURCE_GITHUB:
        urls = [DEFAULT_SETTINGS_URL_GITHUB]
    else:
        urls = [DEFAULT_SETTINGS_URL_GITEE, DEFAULT_SETTINGS_URL_GITHUB]
    settings_path = get_settings_path(target_dir) if target_dir else SETTINGS_PATH
    if target_dir:
        os.makedirs(target_dir, exist_ok=True)
    gitee_token = _get_env_token("GITEE_TOKEN", "GITEE_ACCESS_TOKEN")
    for url in urls:
        try:
            download_url = url
            if "gitee.com" in url:
                download_url = _append_query_param(url, "access_token", gitee_token)
            _download_file(download_url, settings_path, expected_size=0)
            return True
        except Exception:
            continue
    return False

def ensure_default_settings(update_source: str, target_dir: str = None):
    settings_path = get_settings_path(target_dir) if target_dir else SETTINGS_PATH
    if os.path.exists(settings_path):
        return True, None
    info, err = check_update_by_source(APP_VERSION, update_source)
    if info:
        return False, info
    ok = download_default_settings(update_source, target_dir)
    if not ok and err:
        return False, None
    return ok, None

def save_settings(s):
    try:
        with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
            json.dump(s, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

# -------------------- Favicon (best-effort) --------------------
def safe_filename(s: str) -> str:
    return "".join(ch for ch in s if ch.isalnum() or ch in ("-", "_", "."))[:80] or "site"

def download_to(path: str, url: str):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0", "Accept": "*/*"})
    with urllib.request.urlopen(req, timeout=8) as resp:
        data = resp.read()
    with open(path, "wb") as f:
        f.write(data)

def load_icon_meta():
    try:
        with open(ICON_META_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def save_icon_meta(meta):
    try:
        with open(ICON_META_PATH, "w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def ensure_icon_assets():
    os.makedirs(ICON_DIR, exist_ok=True)
    meta = load_icon_meta()
    changed = False
    for name, url in ICON_URLS.items():
        path = ICON_FILES.get(name)
        if not path:
            continue
        if os.path.exists(path) and meta.get(name) == url:
            continue
        try:
            download_to(path, url)
            meta[name] = url
            changed = True
        except Exception:
            pass
    if changed:
        save_icon_meta(meta)

def load_icon_image(path: str, master=None, max_size=32):
    if not path or not os.path.exists(path):
        return None
    pix = QtGui.QPixmap(path)
    if pix.isNull():
        return None
    try:
        w, h = pix.width(), pix.height()
        if max_size and max(w, h) > max_size:
            pix = pix.scaled(max_size, max_size, QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation)
    except Exception:
        pass
    return pix

def load_any_image(path: str, master=None, max_size=None):
    if not path or not os.path.exists(path):
        return None
    pix = QtGui.QPixmap(path)
    if pix.isNull():
        return None
    if max_size:
        try:
            pix = pix.scaled(max_size[0], max_size[1], QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation)
        except Exception:
            pass
    return pix

def is_icon_file(path: str):
    ext = os.path.splitext(path)[1].lower()
    return ext in ALLOWED_ICON_EXTS

def icon_data_url_from_path(path: str):
    if not path or not os.path.exists(path):
        return "", ""
    ext = os.path.splitext(path)[1].lower()
    mime = {
        ".png": "image/png",
        ".gif": "image/gif",
        ".ico": "image/x-icon",
    }.get(ext, "")
    if not mime:
        return "", ""
    try:
        with open(path, "rb") as f:
            data = base64.b64encode(f.read()).decode("ascii")
        return f"data:{mime};base64,{data}", mime
    except Exception:
        return "", ""

def get_best_icon_url(driver):
    js = r"""
    (function(){
      function pick(sel){
        var el = document.querySelector(sel);
        return el && el.href ? el.href : "";
      }
      var a = pick("link[rel='apple-touch-icon']");
      if (a) return a;
      var p = pick("link[rel~='icon'][href$='.png']");
      if (p) return p;
      var i = pick("link[rel~='icon']");
      if (i) return i;
      return "";
    })();
    """
    try:
        return driver.execute_script(js) or ""
    except Exception:
        return ""

def fallback_favicon_url(page_url: str):
    try:
        u = urllib.parse.urlparse(page_url)
        if not u.scheme or not u.netloc:
            return ""
        return f"{u.scheme}://{u.netloc}/favicon.ico"
    except Exception:
        return ""

# -------------------- Early-inject script to prevent new windows --------------------
PREVENT_NEW_WINDOWS_JS = r"""
  (() => {
    // Force window.open to navigate same window
    try {
      const _open = window.open;
      window.open = function(url){
        if (url) {
          try { location.href = url; } catch(e) {}
        }
        return window;
      };
    } catch(e) {}

    // Make all links open in same window
    function fixAnchors(root){
    try {
      const as = (root || document).querySelectorAll ? (root || document).querySelectorAll("a[target]") : [];
      for (const a of as) {
        if (a.target && a.target.toLowerCase() !== "_self") a.target = "_self";
      }
    } catch(e) {}
  }

  fixAnchors(document);

  // Watch DOM for new anchors
  try {
    const obs = new MutationObserver((muts) => {
      for (const m of muts) {
        if (m.addedNodes) {
          for (const n of m.addedNodes) {
            if (n && n.querySelectorAll) fixAnchors(n);
          }
        }
      }
    });
    obs.observe(document.documentElement, {childList:true, subtree:true});
    } catch(e) {}

    // Capture clicks early and force _self
    try {
      document.addEventListener("click", (e) => {
        const a = e.target && e.target.closest ? e.target.closest("a") : null;
        if (!a) return;
        const rawHref = (a.getAttribute && a.getAttribute("href")) ? a.getAttribute("href") : "";
        if (!rawHref) return;
        const lowHref = rawHref.toLowerCase();
        if (lowHref.startsWith("javascript:") || lowHref.startsWith("#")) return;
        const target = (a.getAttribute && a.getAttribute("target")) ? a.getAttribute("target").toLowerCase() : "";
        if (target && target != "_self") {
          try { e.preventDefault(); } catch(e) {}
          try { location.href = a.href; } catch(e) {}
        }
        if (a.target && a.target.toLowerCase() !== "_self") a.target = "_self";
      }, true);
    } catch(e) {}
  })();
  """

# -------------------- Built-in generic icons (NOT branded) --------------------
def make_icon(style: str):
    # Very simple 16x16 pixel icons made with QImage pixels.
    img = QtGui.QImage(16, 16, QtGui.QImage.Format_ARGB32)
    bg = QtGui.QColor("#f2f2f2")
    img.fill(bg)

    def dot(x, y, c):
        if 0 <= x < 16 and 0 <= y < 16:
            img.setPixelColor(x, y, QtGui.QColor(c))

    if style == "globe":
        # circle-ish + meridian
        c1 = "#2f6fff"
        c2 = "#1d3f99"
        for y in range(16):
            for x in range(16):
                dx, dy = x-8, y-8
                if dx*dx + dy*dy <= 36:
                    dot(x, y, c1)
        for y in range(4, 13):
            dot(8, y, c2)
        for x in range(4, 13):
            dot(x, 8, c2)

    elif style == "video":
        c1 = "#111111"
        c2 = "#ffffff"
        for y in range(4, 12):
            for x in range(3, 13):
                dot(x, y, c1)
        # play triangle
        for i in range(0, 5):
            for j in range(i+1):
                dot(6+i, 6+j, c2)
                dot(6+i, 10-j, c2)

    elif style == "chat":
        c1 = "#111111"
        c2 = "#ffffff"
        for y in range(4, 11):
            for x in range(3, 13):
                dot(x, y, c1)
        # tail
        dot(5, 11, c1); dot(4, 12, c1); dot(5, 12, c1)
        # dots
        dot(6, 7, c2); dot(8, 7, c2); dot(10, 7, c2)

    elif style == "folder":
        c1 = "#c9a300"
        c2 = "#8b6f00"
        for y in range(6, 13):
            for x in range(3, 13):
                dot(x, y, c1)
        for x in range(4, 9):
            dot(x, 5, c1)
        for x in range(3, 13):
            dot(x, 6, c2)

    else:  # "star"
        c1 = "#111111"
        pts = [(8,3),(9,6),(12,6),(10,8),(11,12),(8,10),(5,12),(6,8),(4,6),(7,6)]
        for x,y in pts:
            dot(x,y,c1)
        dot(8,6,c1); dot(8,7,c1); dot(8,8,c1); dot(7,8,c1); dot(9,8,c1)

    return QtGui.QPixmap.fromImage(img)

# -------------------- App --------------------
class QtVar:
    def __init__(self, getter, setter=None, signal=None):
        self._getter = getter
        self._setter = setter
        self._signal = signal

    def get(self):
        return self._getter()

    def set(self, value):
        if self._setter:
            self._setter(value)

    def trace_add(self, _mode, callback):
        if self._signal:
            self._signal.connect(lambda *_: callback())

class FloatVar:
    def __init__(self, slider, scale=100):
        self.slider = slider
        self.scale = scale

    def get(self):
        return self.slider.value() / self.scale

    def set(self, value):
        self.slider.setValue(int(round(value * self.scale)))

class ButtonGroupVar:
    def __init__(self, group, mapping):
        self.group = group
        self.mapping = mapping
        self.reverse = {btn: val for val, btn in mapping.items()}

    def get(self):
        btn = self.group.checkedButton()
        return self.reverse.get(btn, "")

    def set(self, value):
        btn = self.mapping.get(value)
        if btn:
            btn.setChecked(True)

class LabelVar:
    def __init__(self, label):
        self.label = label

    def get(self):
        return self.label.text()

    def set(self, value):
        self.label.setText(value)

class MiniFish(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        hide_console_window()
        os.makedirs(CACHE_DIR, exist_ok=True)
        ensure_icon_assets()

        self.chrome_exe = find_chrome_exe()
        if not self.chrome_exe:
            raise RuntimeError("找不到 Chrome/Edge。请先安装 Chrome。")

        self._settings_missing = not os.path.exists(SETTINGS_PATH)
        self._forced_update_info = None
        self._settings_downloaded = False
        self._settings_target_dir = ""
        self._pending_relocate_dir = ""
        self._update_target_dir = ""
        self._skip_update_check = False
        self._first_run_hint_needed = bool(self._settings_missing)
        self._first_run_hint = None
        self._first_run_blink_timer = None
        self._first_run_blink_state = True
        if self._settings_missing:
            self._handle_missing_settings()

        self.settings = load_settings()
        self._maybe_show_first_run_hint()
        if self.settings.get("panel_title") in DEFAULT_PANEL_TITLES:
            self.settings["panel_title"] = APP_TITLE
            save_settings(self.settings)
        self.ahk_exe = find_ahk_exe()
        self._ahk_proc = None
        self._ahk_use_v2 = False
        self.proc = None
        self.driver = None
        self.port = None
        self.chrome_hwnd = None
        self.site_icon_pixmap = None
        self.panel_icon_pixmap = None
        self.browser_icon_data_url = ""
        self.browser_icon_mime = ""
        self.hidden_toggle = False
        self.lock_hidden = False
        self._global_hotkeys = {}
        self._local_shortcuts = []
        self.hotkey_toggle = self.settings.get("hotkey_toggle", HOTKEY_DEFAULT_TOGGLE)
        self.hotkey_lock = self.settings.get("hotkey_lock", HOTKEY_DEFAULT_LOCK)
        self.hotkey_close = self.settings.get("hotkey_close", HOTKEY_DEFAULT_CLOSE)
        self.extra_sessions = []
        self.profile_dir = PROFILE_DIR
        self.attach_enabled = bool(self.settings.get("attach_enabled", False))
        self.attach_side = self.settings.get("attach_side", "left")
        self.merge_taskbar = bool(self.settings.get("merge_taskbar", False))
        self.audio_enabled = bool(self.settings.get("audio_enabled", True))
        self.sync_taskbar_icon = bool(self.settings.get("sync_taskbar_icon", True))
        self.panel_tray_enabled = bool(self.settings.get("panel_tray_enabled", False))
        self.browser_tray_enabled = bool(self.settings.get("browser_tray_enabled", False))
        self._force_close = False
        self._last_panel_pos = None
        self._last_browser_rect = None
        self._last_audio_url = ""
        self._syncing_attach = False
        self._attach_timer = None
        self._attach_retry = 0
        self._browser_gone_ticks = 0
        self.main_window_handle = None
        self._initializing = True
        self._starting_browser = False
        self._browser_start_thread = None
        self._browser_start_worker = None
        self._multi_start_thread = None
        self._multi_start_worker = None
        self.panel_tray = None
        self.browser_tray = None
        self.update_available = False
        self.update_info = None
        self._update_checked_once = False
        self._update_prompted_version = ""
        self._update_check_thread = None
        self._update_check_worker = None
        self._update_check_force = False
        self._update_download_thread = None
        self._update_download_worker = None
        self._update_progress_dialog = None
        self._update_progress_label = None
        self._update_progress_bar = None

        self.win_w = 460
        self.win_h = 340

        self.setAttribute(QtCore.Qt.WA_NativeWindow, True)
        self.setWindowTitle(APP_TITLE)
        self.setSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)

        main_layout = QtWidgets.QVBoxLayout(self)
        main_layout.setContentsMargins(8, 8, 8, 8)
        main_layout.setSpacing(6)

        # presets + add/remove
        rowp = QtWidgets.QHBoxLayout()
        rowp.setContentsMargins(0, 0, 0, 0)
        self.preset_combo = QtWidgets.QComboBox()
        self.preset_combo.setEditable(True)
        self.preset_combo.addItems(self.settings["presets"])
        self.preset_combo.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        self.preset_combo.activated.connect(lambda _: self.pick_preset())
        rowp.addWidget(self.preset_combo, 1)
        btn_add = QtWidgets.QPushButton("+")
        btn_add.setFixedWidth(28)
        btn_add.clicked.connect(self.add_preset)
        rowp.addWidget(btn_add)
        btn_remove = QtWidgets.QPushButton("-")
        btn_remove.setFixedWidth(28)
        btn_remove.clicked.connect(self.remove_preset)
        rowp.addWidget(btn_remove)
        btn_multi = QtWidgets.QPushButton("多开")
        btn_multi.setCheckable(True)
        btn_multi.setToolTip("多开模式：蓝色为开启，按 Go 新开窗口")
        btn_multi.setStyleSheet("QPushButton:checked { background-color: #2f6fff; color: white; }")
        rowp.addWidget(btn_multi)
        self.multi_open_button = btn_multi
        btn_sponsor = QtWidgets.QPushButton("赞助作者")
        btn_sponsor.clicked.connect(self.open_sponsor_dialog)
        rowp.addWidget(btn_sponsor)
        btn_github = QtWidgets.QPushButton("Github源码")
        btn_github.clicked.connect(self.open_github)
        rowp.addWidget(btn_github)
        btn_about = QtWidgets.QPushButton("关于")
        btn_about.clicked.connect(self.open_about_dialog)
        rowp.addWidget(btn_about)
        self.about_button = btn_about
        self.about_badge = QtWidgets.QLabel(btn_about)
        self.about_badge.setFixedSize(8, 8)
        self.about_badge.setStyleSheet("background-color: #ef4444; border-radius: 4px;")
        self.about_badge.setToolTip("发现新版本")
        self.about_badge.hide()
        self.about_badge.raise_()
        self._position_about_badge()
        btn_about.installEventFilter(self)
        main_layout.addLayout(rowp)

        # URL row
        row0 = QtWidgets.QHBoxLayout()
        row0.setContentsMargins(0, 0, 0, 0)
        self.url_edit = QtWidgets.QLineEdit()
        self.url_edit.setText(self.settings.get("last_url") or "https://www.douyin.com")
        self.url_edit.returnPressed.connect(self.go)
        row0.addWidget(self.url_edit, 1)
        btn_go = QtWidgets.QPushButton("Go")
        btn_go.setFixedWidth(40)
        btn_go.clicked.connect(self.go)
        row0.addWidget(btn_go)
        main_layout.addLayout(row0)

        # icon rows
        rowi = QtWidgets.QHBoxLayout()
        rowi.setContentsMargins(0, 0, 0, 0)
        rowi.addWidget(QtWidgets.QLabel("面板图标"))
        self.panel_icon_style_combo = QtWidgets.QComboBox()
        self._populate_icon_combo(self.panel_icon_style_combo, PANEL_ICON_CHOICES)
        self._set_combo_by_data(self.panel_icon_style_combo, self.settings.get("panel_icon_style", "globe"))
        self.panel_icon_style_combo.currentTextChanged.connect(lambda _: self.apply_panel_icon_style(auto_rename=True))
        rowi.addWidget(self.panel_icon_style_combo)
        btn_panel_icon = QtWidgets.QPushButton("选择")
        btn_panel_icon.setFixedWidth(40)
        btn_panel_icon.clicked.connect(self.choose_panel_icon)
        rowi.addWidget(btn_panel_icon)

        rowi.addSpacing(8)
        rowi.addWidget(QtWidgets.QLabel("浏览器图标"))
        self.browser_icon_style_combo = QtWidgets.QComboBox()
        self._populate_icon_combo(self.browser_icon_style_combo, BROWSER_ICON_CHOICES)
        self._set_combo_by_data(self.browser_icon_style_combo, self.settings.get("browser_icon_style", "site"))
        self.browser_icon_style_combo.currentTextChanged.connect(lambda _: self.apply_browser_icon_style(auto_rename=True))
        rowi.addWidget(self.browser_icon_style_combo)
        btn_browser_icon = QtWidgets.QPushButton("选择")
        btn_browser_icon.setFixedWidth(40)
        btn_browser_icon.clicked.connect(self.choose_browser_icon)
        rowi.addWidget(btn_browser_icon)
        main_layout.addLayout(rowi)

        rowi2 = QtWidgets.QHBoxLayout()
        rowi2.setContentsMargins(0, 0, 0, 0)
        rowi2.addWidget(QtWidgets.QLabel("任务栏图标"))
        self.taskbar_sync_checkbox = ToggleSwitch()
        self.taskbar_sync_checkbox.setChecked(self.sync_taskbar_icon)
        self.taskbar_sync_checkbox.toggled.connect(self.on_taskbar_sync_toggle)
        rowi2.addWidget(self.taskbar_sync_checkbox)
        btn_taskbar_sync = QtWidgets.QPushButton("同步一次")
        btn_taskbar_sync.clicked.connect(self.update_taskbar_icon)
        rowi2.addWidget(btn_taskbar_sync)
        rowi2.addStretch(1)
        main_layout.addLayout(rowi2)

        rowi3 = QtWidgets.QHBoxLayout()
        rowi3.setContentsMargins(0, 0, 0, 0)
        rowi3.addWidget(QtWidgets.QLabel("任务栏驻留"))
        rowi3.addWidget(QtWidgets.QLabel("面板"))
        self.panel_tray_checkbox = ToggleSwitch()
        self.panel_tray_checkbox.setChecked(self.panel_tray_enabled)
        self.panel_tray_checkbox.toggled.connect(self.on_panel_tray_toggle)
        rowi3.addWidget(self.panel_tray_checkbox)
        rowi3.addSpacing(6)
        rowi3.addWidget(QtWidgets.QLabel("浏览器"))
        self.browser_tray_checkbox = ToggleSwitch()
        self.browser_tray_checkbox.setChecked(self.browser_tray_enabled)
        self.browser_tray_checkbox.toggled.connect(self.on_browser_tray_toggle)
        rowi3.addWidget(self.browser_tray_checkbox)
        rowi3.addStretch(1)
        main_layout.addLayout(rowi3)

        # titles
        rowt = QtWidgets.QHBoxLayout()
        rowt.setContentsMargins(0, 0, 0, 0)
        rowt.addWidget(QtWidgets.QLabel("面板名"))
        self.panel_title_edit = QtWidgets.QLineEdit()
        self.panel_title_edit.setText(self.settings.get("panel_title", "mini"))
        self.panel_title_edit.setFixedWidth(120)
        rowt.addWidget(self.panel_title_edit)
        rowt.addSpacing(8)
        rowt.addWidget(QtWidgets.QLabel("浏览器名"))
        self.browser_title_edit = QtWidgets.QLineEdit()
        self.browser_title_edit.setText(self.settings.get("browser_title", "mini-browser"))
        self.browser_title_edit.setFixedWidth(160)
        rowt.addWidget(self.browser_title_edit)
        btn_save = QtWidgets.QPushButton("保存")
        btn_save.clicked.connect(self.apply_titles_and_refresh)
        rowt.addWidget(btn_save)
        main_layout.addLayout(rowt)

        # custom status text
        row0b = QtWidgets.QHBoxLayout()
        row0b.setContentsMargins(0, 0, 0, 0)
        row0b.addWidget(QtWidgets.QLabel("状态栏名"))
        self.custom_status_edit = QtWidgets.QLineEdit()
        self.custom_status_edit.setText(self.settings.get("custom_status", ""))
        row0b.addWidget(self.custom_status_edit, 1)
        btn_clear = QtWidgets.QPushButton("清除")
        btn_clear.setFixedWidth(40)
        btn_clear.clicked.connect(self.clear_custom_status)
        row0b.addWidget(btn_clear)
        main_layout.addLayout(row0b)

        # window size presets
        row_size = QtWidgets.QHBoxLayout()
        row_size.setContentsMargins(0, 0, 0, 0)
        row_size.addWidget(QtWidgets.QLabel("窗口比例"))
        ratio_label = RATIO_KEY_TO_LABEL.get(self.settings.get("browser_ratio", "4:3"), "4:3 横")
        self.browser_ratio_combo = QtWidgets.QComboBox()
        self.browser_ratio_combo.addItems(RATIO_LABELS)
        self.browser_ratio_combo.setCurrentText(ratio_label)
        self.browser_ratio_combo.currentTextChanged.connect(lambda _: self.on_browser_ratio_change())
        row_size.addWidget(self.browser_ratio_combo)
        row_size.addSpacing(8)
        row_size.addWidget(QtWidgets.QLabel("尺寸"))
        self.size_group = QtWidgets.QButtonGroup(self)
        self.size_s = QtWidgets.QRadioButton("小")
        self.size_m = QtWidgets.QRadioButton("中")
        self.size_l = QtWidgets.QRadioButton("大")
        self.size_group.addButton(self.size_s)
        self.size_group.addButton(self.size_m)
        self.size_group.addButton(self.size_l)
        row_size.addWidget(self.size_s)
        row_size.addWidget(self.size_m)
        row_size.addWidget(self.size_l)
        self.size_group.buttonClicked.connect(lambda _: self.on_browser_size_level_change())
        main_layout.addLayout(row_size)

        # window position
        row_pos = QtWidgets.QHBoxLayout()
        row_pos.setContentsMargins(0, 0, 0, 0)
        row_pos.addWidget(QtWidgets.QLabel("窗口位置"))
        pos_label = BROWSER_POS_KEY_TO_LABEL.get(self.settings.get("browser_position", "bottom_right"), "右下角")
        self.browser_pos_combo = QtWidgets.QComboBox()
        self.browser_pos_combo.addItems(BROWSER_POS_LABELS)
        self.browser_pos_combo.setCurrentText(pos_label)
        self.browser_pos_combo.currentTextChanged.connect(lambda _: self.on_browser_position_change())
        row_pos.addWidget(self.browser_pos_combo)
        main_layout.addLayout(row_pos)

        # window size scale
        row_scale = QtWidgets.QHBoxLayout()
        row_scale.setContentsMargins(0, 0, 0, 0)
        row_scale.addWidget(QtWidgets.QLabel("窗口缩放"))
        self.window_scale_slider = QtWidgets.QSlider(QtCore.Qt.Horizontal)
        self.window_scale_slider.setRange(int(BROWSER_SCALE_MIN * 100), int(BROWSER_SCALE_MAX * 100))
        self.window_scale_slider.setValue(int(float(self.settings.get("browser_scale", 1.0)) * 100))
        self.window_scale_slider.valueChanged.connect(self.on_window_scale)
        row_scale.addWidget(self.window_scale_slider, 1)
        self.window_scale_label = QtWidgets.QLabel("")
        row_scale.addWidget(self.window_scale_label)
        self.window_size_label = QtWidgets.QLabel("")
        row_scale.addWidget(self.window_size_label)
        main_layout.addLayout(row_scale)

        # zoom
        row1 = QtWidgets.QHBoxLayout()
        row1.setContentsMargins(0, 0, 0, 0)
        row1.addWidget(QtWidgets.QLabel("页面缩放"))
        self.zoom_slider = QtWidgets.QSlider(QtCore.Qt.Horizontal)
        self.zoom_slider.setRange(20, 300)
        self.zoom_slider.setValue(int(float(self.settings.get("remember_zoom", 0.85)) * 100))
        self.zoom_slider.valueChanged.connect(self.on_zoom)
        row1.addWidget(self.zoom_slider, 1)
        self.zoom_label = QtWidgets.QLabel("")
        row1.addWidget(self.zoom_label)
        main_layout.addLayout(row1)

        # browser alpha
        row2 = QtWidgets.QHBoxLayout()
        row2.setContentsMargins(0, 0, 0, 0)
        row2.addWidget(QtWidgets.QLabel("浏览器透明"))
        self.alpha_slider = QtWidgets.QSlider(QtCore.Qt.Horizontal)
        self.alpha_slider.setRange(20, 100)
        self.alpha_slider.setValue(int(float(self.settings.get("remember_alpha", 1.0)) * 100))
        self.alpha_slider.valueChanged.connect(self.on_alpha)
        row2.addWidget(self.alpha_slider, 1)
        self.alpha_label = QtWidgets.QLabel("")
        row2.addWidget(self.alpha_label)
        main_layout.addLayout(row2)

        # audio
        row2b = QtWidgets.QHBoxLayout()
        row2b.setContentsMargins(0, 0, 0, 0)
        row2b.addWidget(QtWidgets.QLabel("声音启用"))
        self.audio_checkbox = ToggleSwitch()
        self.audio_checkbox.setChecked(self.audio_enabled)
        self.audio_checkbox.toggled.connect(self.on_audio_toggle)
        row2b.addWidget(self.audio_checkbox)
        row2b.addStretch(1)
        main_layout.addLayout(row2b)

        # panel alpha
        row3 = QtWidgets.QHBoxLayout()
        row3.setContentsMargins(0, 0, 0, 0)
        row3.addWidget(QtWidgets.QLabel("面板透明"))
        self.panel_alpha_slider = QtWidgets.QSlider(QtCore.Qt.Horizontal)
        self.panel_alpha_slider.setRange(30, 100)
        self.panel_alpha_slider.setValue(int(float(self.settings.get("remember_panel_alpha", 1.0)) * 100))
        self.panel_alpha_slider.valueChanged.connect(self.on_panel_alpha)
        row3.addWidget(self.panel_alpha_slider, 1)
        self.panel_alpha_label = QtWidgets.QLabel("")
        row3.addWidget(self.panel_alpha_label)
        main_layout.addLayout(row3)

        # topmost
        row4 = QtWidgets.QHBoxLayout()
        row4.setContentsMargins(0, 0, 0, 0)
        row4.addWidget(QtWidgets.QLabel("置顶面板"))
        self.panel_top_checkbox = ToggleSwitch()
        self.panel_top_checkbox.setChecked(bool(self.settings.get("panel_topmost", True)))
        self.panel_top_checkbox.toggled.connect(self.apply_panel_topmost)
        row4.addWidget(self.panel_top_checkbox)
        row4.addSpacing(8)
        row4.addWidget(QtWidgets.QLabel("置顶浏览器窗"))
        self.browser_top_checkbox = ToggleSwitch()
        self.browser_top_checkbox.setChecked(bool(self.settings.get("browser_topmost", False)))
        self.browser_top_checkbox.toggled.connect(self.on_browser_top_toggle)
        row4.addWidget(self.browser_top_checkbox)
        row4.addStretch(1)
        btn_find = QtWidgets.QPushButton("找回浏览器")
        btn_find.clicked.connect(self.recover_browser)
        row4.addWidget(btn_find)
        btn_restart = QtWidgets.QPushButton("重启浏览器")
        btn_restart.clicked.connect(self.restart_browser)
        row4.addWidget(btn_restart)
        btn_hotkey = QtWidgets.QPushButton("快捷键")
        btn_hotkey.clicked.connect(self.open_hotkey_dialog)
        row4.addWidget(btn_hotkey)
        main_layout.addLayout(row4)

        # attach + restack
        row4b = QtWidgets.QHBoxLayout()
        row4b.setContentsMargins(0, 0, 0, 0)
        row4b.addWidget(QtWidgets.QLabel("吸附面板"))
        self.attach_checkbox = ToggleSwitch()
        self.attach_checkbox.setChecked(self.attach_enabled)
        self.attach_checkbox.toggled.connect(self.on_attach_toggle)
        row4b.addWidget(self.attach_checkbox)
        row4b.addWidget(QtWidgets.QLabel("位置"))
        self.attach_side_combo = QtWidgets.QComboBox()
        self.attach_side_combo.addItems(["左侧", "右侧"])
        self.attach_side_combo.setCurrentText("左侧" if self.attach_side == "left" else "右侧")
        self.attach_side_combo.currentTextChanged.connect(self.on_attach_side_change)
        row4b.addWidget(self.attach_side_combo)
        row4b.addWidget(QtWidgets.QLabel("任务栏合并(实验)"))
        self.merge_checkbox = ToggleSwitch()
        self.merge_checkbox.setChecked(self.merge_taskbar)
        self.merge_checkbox.toggled.connect(self.apply_taskbar_merge)
        row4b.addWidget(self.merge_checkbox)
        row4b.addStretch(1)
        btn_restack = QtWidgets.QPushButton("强制重排")
        btn_restack.clicked.connect(self.force_restack)
        row4b.addWidget(btn_restack)
        main_layout.addLayout(row4b)

        # status bar
        bar = QtWidgets.QHBoxLayout()
        bar.setContentsMargins(0, 0, 0, 0)
        self.icon_label = QtWidgets.QLabel("◎")
        self.icon_label.setFixedWidth(20)
        self.icon_label.setAlignment(QtCore.Qt.AlignCenter)
        bar.addWidget(self.icon_label)
        self.status_label = QtWidgets.QLabel("ready")
        self.status_label.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        bar.addWidget(self.status_label, 1)
        main_layout.addLayout(bar)

        # vars
        self.preset_var = QtVar(self.preset_combo.currentText, self.preset_combo.setCurrentText)
        self.url_var = QtVar(self.url_edit.text, self.url_edit.setText)
        self.panel_icon_style_var = QtVar(
            lambda: self._combo_data(self.panel_icon_style_combo),
            lambda v: self._set_combo_by_data(self.panel_icon_style_combo, v),
        )
        self.browser_icon_style_var = QtVar(
            lambda: self._combo_data(self.browser_icon_style_combo),
            lambda v: self._set_combo_by_data(self.browser_icon_style_combo, v),
        )
        self.panel_title_var = QtVar(self.panel_title_edit.text, self.panel_title_edit.setText)
        self.browser_title_var = QtVar(self.browser_title_edit.text, self.browser_title_edit.setText)
        self.custom_status_var = QtVar(self.custom_status_edit.text, self.custom_status_edit.setText, self.custom_status_edit.textChanged)
        self.browser_ratio_var = QtVar(self.browser_ratio_combo.currentText, self.browser_ratio_combo.setCurrentText)
        self.browser_pos_var = QtVar(self.browser_pos_combo.currentText, self.browser_pos_combo.setCurrentText)
        self.window_scale_var = FloatVar(self.window_scale_slider)
        self.zoom_var = FloatVar(self.zoom_slider)
        self.alpha_var = FloatVar(self.alpha_slider)
        self.panel_alpha_var = FloatVar(self.panel_alpha_slider)
        self.panel_top_var = QtVar(self.panel_top_checkbox.isChecked, self.panel_top_checkbox.setChecked)
        self.browser_top_var = QtVar(self.browser_top_checkbox.isChecked, self.browser_top_checkbox.setChecked)
        self.audio_var = QtVar(self.audio_checkbox.isChecked, self.audio_checkbox.setChecked)
        self.taskbar_sync_var = QtVar(self.taskbar_sync_checkbox.isChecked, self.taskbar_sync_checkbox.setChecked)
        self.status_var = LabelVar(self.status_label)
        self.browser_size_level_var = ButtonGroupVar(self.size_group, {"S": self.size_s, "M": self.size_m, "L": self.size_l})
        self.browser_size_level_var.set(self.settings.get("browser_size_level", "S"))

        self.custom_status_edit.textChanged.connect(self.on_custom_status_change)

        # apply panel icon + alpha + topmost
        self.apply_titles(save=False)
        self.apply_panel_icon_style()
        self.apply_browser_icon_style()
        self.apply_browser_window_size(resize_now=False)
        self.zoom_label.setText(f"{self.zoom_var.get():.2f}x")
        self.alpha_label.setText(f"{self.alpha_var.get():.2f}")
        self.panel_alpha_label.setText(f"{self.panel_alpha_var.get():.2f}")
        self.on_panel_alpha(None)
        self.apply_panel_topmost()
        self.setup_icon_drop()
        self.adjustSize()
        self.setFixedSize(self.sizeHint())
        self.center_panel()
        self._sync_tray_icons()

        QtCore.QTimer.singleShot(0, lambda: self.apply_hotkeys(self.hotkey_toggle, self.hotkey_lock, self.hotkey_close, save=False))
        if self.browser_top_checkbox.isChecked():
            QtCore.QTimer.singleShot(200, self._show_topmost_invalid_hint)

        self.status_var.set("等待Go启动浏览器")
        self.state_timer = QtCore.QTimer(self)
        self.state_timer.timeout.connect(self.poll_state)
        self.state_timer.start(1200)
        self._initializing = False
        if self._first_run_hint:
            QtCore.QTimer.singleShot(1500, self._finish_first_run_hint)
        if self._pending_relocate_dir:
            QtCore.QTimer.singleShot(200, self._start_relocation)
            return
        if self._forced_update_info:
            self.update_available = True
            self.update_info = self._forced_update_info
            self._set_update_badge(True)
            QtCore.QTimer.singleShot(300, lambda: self._prompt_update_required(self.update_info))
        elif not self._skip_update_check:
            QtCore.QTimer.singleShot(1200, lambda: self.check_update_async(force=False))

    def _show_warning(self, text: str, title: str = "提示"):
        try:
            QtWidgets.QMessageBox.warning(self, title, text)
        except Exception:
            pass

    def _show_error(self, text: str, title: str = "错误"):
        try:
            QtWidgets.QMessageBox.critical(self, title, text)
        except Exception:
            pass

    def _show_update_dialog(self, info: dict, required: bool) -> bool:
        version = info.get("version") or ""
        notes = (info.get("body") or "").strip() or "暂无更新说明"
        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle("需要更新" if required else "发现新版本")
        dlg.setModal(True)
        layout = QtWidgets.QVBoxLayout(dlg)
        title = QtWidgets.QLabel(
            f"检测到新版本 {version}，需要更新后才能继续使用。"
            if required
            else f"发现新版本 {version}，是否立即更新？"
        )
        title.setWordWrap(True)
        layout.addWidget(title)
        text = QtWidgets.QTextEdit()
        text.setReadOnly(True)
        text.setPlainText(notes)
        text.setMinimumHeight(180)
        layout.addWidget(text)
        btn_row = QtWidgets.QHBoxLayout()
        btn_row.addStretch(1)
        btn_update = QtWidgets.QPushButton("立即更新" if required else "确认更新")
        btn_update.clicked.connect(dlg.accept)
        btn_row.addWidget(btn_update)
        if not required:
            btn_cancel = QtWidgets.QPushButton("我再想想")
            btn_cancel.clicked.connect(dlg.reject)
            btn_row.addWidget(btn_cancel)
        layout.addLayout(btn_row)
        return dlg.exec() == QtWidgets.QDialog.Accepted

    def _should_show_first_run_hint(self) -> bool:
        if self._first_run_hint_needed:
            return True
        try:
            flag_path = get_first_run_flag_path(get_app_dir())
            if os.path.exists(flag_path):
                return True
        except Exception:
            return False
        return False

    def _maybe_show_first_run_hint(self):
        if self._first_run_hint:
            return
        if not self._should_show_first_run_hint():
            return
        self._show_first_run_hint()

    def _show_first_run_hint(self):
        text = "首次启动部署配置加载中，请稍等..."
        label = QtWidgets.QLabel(text)
        label.setWindowFlags(
            QtCore.Qt.Tool | QtCore.Qt.FramelessWindowHint | QtCore.Qt.WindowStaysOnTopHint
        )
        label.setAttribute(QtCore.Qt.WA_TranslucentBackground, True)
        font = label.font()
        font.setPointSize(16)
        font.setBold(True)
        label.setFont(font)
        label.setAlignment(QtCore.Qt.AlignCenter)
        label.setStyleSheet(
            "color: #ffffff; background-color: rgba(0, 0, 0, 180);"
            "padding: 16px 28px; border-radius: 10px;"
        )
        label.adjustSize()
        screen = QtGui.QGuiApplication.primaryScreen()
        if screen:
            geo = screen.availableGeometry()
            x = geo.x() + (geo.width() - label.width()) // 2
            y = geo.y() + (geo.height() - label.height()) // 2
            label.move(max(x, 0), max(y, 0))
        label.show()
        label.raise_()
        self._first_run_hint = label
        self._first_run_blink_timer = QtCore.QTimer(self)
        self._first_run_blink_timer.timeout.connect(self._blink_first_run_hint)
        self._first_run_blink_timer.start(500)

    def _blink_first_run_hint(self):
        if not self._first_run_hint:
            return
        self._first_run_blink_state = not self._first_run_blink_state
        color = "#ffffff" if self._first_run_blink_state else "#facc15"
        self._first_run_hint.setStyleSheet(
            f"color: {color}; background-color: rgba(0, 0, 0, 180);"
            "padding: 16px 28px; border-radius: 10px;"
        )

    def _finish_first_run_hint(self):
        if self._first_run_blink_timer:
            self._first_run_blink_timer.stop()
        if self._first_run_hint:
            self._first_run_hint.close()
            self._first_run_hint.deleteLater()
        self._first_run_hint = None
        self._first_run_blink_timer = None
        try:
            flag_path = get_first_run_flag_path(get_app_dir())
            if os.path.exists(flag_path):
                os.remove(flag_path)
        except Exception:
            pass
        try:
            self.settings["first_run_done"] = True
            save_settings(self.settings)
        except Exception:
            pass

    def _prompt_settings_deploy_dir(self) -> str:
        app_dir = os.path.abspath(get_app_dir())
        box = QtWidgets.QMessageBox(self)
        box.setWindowTitle("配置文件缺失")
        box.setIcon(QtWidgets.QMessageBox.Warning)
        box.setText("检测到配置文件缺失，是否在当前目录部署？")
        btn_current = box.addButton("当前目录", QtWidgets.QMessageBox.AcceptRole)
        btn_new = box.addButton("选择新目录", QtWidgets.QMessageBox.ActionRole)
        btn_cancel = box.addButton("取消", QtWidgets.QMessageBox.RejectRole)
        box.setDefaultButton(btn_current)
        box.exec()
        clicked = box.clickedButton()
        if clicked == btn_new:
            target = QtWidgets.QFileDialog.getExistingDirectory(self, "选择部署目录", app_dir)
            return target or ""
        if clicked == btn_current:
            return app_dir
        return ""

    def _handle_missing_settings(self):
        try:
            bootstrap_path = get_settings_bootstrap_path(get_app_dir())
            if os.path.exists(bootstrap_path):
                try:
                    os.remove(bootstrap_path)
                except Exception:
                    pass
                ok, forced = ensure_default_settings(DEFAULT_UPDATE_SOURCE)
                self._settings_downloaded = bool(ok)
                self._forced_update_info = forced
                self._settings_missing = not bool(ok) and not os.path.exists(SETTINGS_PATH)
                return
            target_dir = self._prompt_settings_deploy_dir()
            if not target_dir:
                return
            app_dir = os.path.abspath(get_app_dir())
            target_dir = os.path.abspath(target_dir)
            if target_dir != app_dir and not is_frozen_app():
                self._show_warning("源码模式不支持自动迁移，将在当前目录创建配置。")
                target_dir = app_dir
            if target_dir == app_dir:
                ok, forced = ensure_default_settings(DEFAULT_UPDATE_SOURCE)
                self._settings_downloaded = bool(ok)
                self._forced_update_info = forced
                self._settings_missing = not bool(ok) and not os.path.exists(SETTINGS_PATH)
                if forced:
                    try:
                        with open(get_settings_bootstrap_path(app_dir), "w", encoding="utf-8") as f:
                            f.write("1")
                    except Exception:
                        pass
                return
            ok, forced = ensure_default_settings(DEFAULT_UPDATE_SOURCE, target_dir)
            self._settings_downloaded = bool(ok)
            self._forced_update_info = forced
            self._settings_target_dir = target_dir
            if forced:
                self._update_target_dir = target_dir
                try:
                    os.makedirs(target_dir, exist_ok=True)
                    with open(get_settings_bootstrap_path(target_dir), "w", encoding="utf-8") as f:
                        f.write("1")
                except Exception:
                    pass
                return
            if ok:
                self._pending_relocate_dir = target_dir
                self._skip_update_check = True
        except Exception:
            pass

    def _start_relocation(self):
        target_dir = self._pending_relocate_dir or ""
        if not target_dir:
            return
        if not is_frozen_app():
            return
        try:
            self._run_relocate_script(target_dir)
            self._force_close = True
            self.close()
        except Exception as e:
            self._show_error(f"迁移失败: {e}")

    def _run_relocate_script(self, target_dir: str):
        exe_path = os.path.abspath(sys.executable)
        app_dir = os.path.dirname(exe_path)
        exe_name = os.path.basename(exe_path)
        ps_path = os.path.join(get_update_cache_dir(), f"_mini_fish_relocate_{int(time.time())}.ps1")
        script = (
            "param([int]$Pid,[string]$SrcDir,[string]$DstDir,[string]$ExeName)\n"
            "while (Get-Process -Id $Pid -ErrorAction SilentlyContinue) { Start-Sleep -Milliseconds 500 }\n"
            "if (!(Test-Path $DstDir)) { New-Item -ItemType Directory -Force $DstDir | Out-Null }\n"
            "$internal = Join-Path $SrcDir \"_internal\"\n"
            "if (Test-Path $internal) { Copy-Item -Recurse -Force $internal (Join-Path $DstDir \"_internal\") }\n"
            "$assets = Join-Path $SrcDir \"assets\"\n"
            "if (Test-Path $assets) { Copy-Item -Recurse -Force $assets (Join-Path $DstDir \"assets\") }\n"
            "Copy-Item -Force (Join-Path $SrcDir $ExeName) (Join-Path $DstDir $ExeName)\n"
            "$settings = Join-Path $SrcDir \"_mini_fish_settings.json\"\n"
            "if (Test-Path $settings) { Copy-Item -Force $settings (Join-Path $DstDir \"_mini_fish_settings.json\") }\n"
            "$newPath = Join-Path $DstDir $ExeName\n"
            "$oldPath = Join-Path $SrcDir $ExeName\n"
            "if (Test-Path $oldPath -and ($oldPath -ne $newPath)) { Remove-Item -Force $oldPath }\n"
            "Start-Process $newPath\n"
            "Remove-Item -Force $MyInvocation.MyCommand.Path\n"
        )
        os.makedirs(os.path.dirname(ps_path), exist_ok=True)
        with open(ps_path, "w", encoding="utf-8-sig") as f:
            f.write(script)
        args = [
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            ps_path,
            "-Pid",
            str(os.getpid()),
            "-SrcDir",
            app_dir,
            "-DstDir",
            target_dir,
            "-ExeName",
            exe_name,
        ]
        subprocess.Popen(args, creationflags=CREATE_NO_WINDOW)

    def _position_about_badge(self):
        try:
            if not self.about_button or not self.about_badge:
                return
            x = self.about_button.width() - self.about_badge.width() - 4
            y = 4
            self.about_badge.move(max(x, 0), y)
        except Exception:
            pass

    def _set_update_badge(self, show: bool):
        try:
            if self.about_badge:
                self.about_badge.setVisible(bool(show))
                if show:
                    self._position_about_badge()
        except Exception:
            pass

    def eventFilter(self, obj, event):
        try:
            if obj == self.about_button and event.type() == QtCore.QEvent.Resize:
                self._position_about_badge()
        except Exception:
            pass
        return super().eventFilter(obj, event)

    def check_update_async(self, force: bool = False):
        if self._update_check_thread:
            return
        self._update_check_force = force
        update_source = self.settings.get("update_source", DEFAULT_UPDATE_SOURCE)
        worker = UpdateCheckWorker(APP_VERSION, update_source)
        thread = QtCore.QThread(self)
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.finished.connect(self._on_update_checked)
        worker.finished.connect(thread.quit)
        worker.finished.connect(worker.deleteLater)
        thread.finished.connect(thread.deleteLater)
        self._update_check_thread = thread
        self._update_check_worker = worker
        try:
            if force:
                self.status_var.set("正在检查更新...")
        except Exception:
            pass
        thread.start()

    def _on_update_checked(self, info, err: str):
        self._update_checked_once = True
        self._update_check_thread = None
        self._update_check_worker = None
        force = bool(self._update_check_force)
        self._update_check_force = False
        if err:
            if force:
                self._show_warning(f"检查更新失败: {err}")
            return
        if not info:
            self.update_available = False
            self.update_info = None
            self._set_update_badge(False)
            if force:
                self._show_warning("已是最新版本")
            return
        self.update_available = True
        self.update_info = info
        self._set_update_badge(True)
        try:
            version = info.get("version") or ""
            if version:
                self.status_var.set(f"发现新版本 {version}")
        except Exception:
            pass
        new_version = info.get("version") or ""
        if force:
            if new_version:
                self._update_prompted_version = new_version
                self._prompt_update(info)
            return
        if new_version and new_version != self._update_prompted_version:
            self._update_prompted_version = new_version
            self._prompt_update(info)

    def _prompt_update(self, info: dict):
        try:
            if self._show_update_dialog(info, required=False):
                self.start_update_download()
        except Exception as e:
            self._show_error(f"更新提示失败: {e}")

    def _prompt_update_required(self, info: dict):
        try:
            self._show_update_dialog(info, required=True)
            self.start_update_download()
        except Exception as e:
            self._show_error(f"更新提示失败: {e}")

    def _show_update_progress_dialog(self):
        if self._update_progress_dialog:
            return
        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle("正在更新")
        dlg.setModal(True)
        layout = QtWidgets.QVBoxLayout(dlg)
        label = QtWidgets.QLabel("正在下载更新...")
        bar = QtWidgets.QProgressBar()
        bar.setRange(0, 100)
        bar.setValue(0)
        layout.addWidget(label)
        layout.addWidget(bar)
        dlg.setLayout(layout)
        dlg.setFixedSize(320, 110)
        dlg.show()
        self._update_progress_dialog = dlg
        self._update_progress_label = label
        self._update_progress_bar = bar

    def _close_update_progress_dialog(self):
        try:
            if self._update_progress_dialog:
                self._update_progress_dialog.close()
                self._update_progress_dialog.deleteLater()
        finally:
            self._update_progress_dialog = None
            self._update_progress_label = None
            self._update_progress_bar = None

    def _on_update_progress(self, done: int, total: int):
        try:
            if not self._update_progress_dialog:
                return
            if total > 0:
                percent = int(done * 100 / total)
                if self._update_progress_bar:
                    self._update_progress_bar.setValue(percent)
                if self._update_progress_label:
                    self._update_progress_label.setText(
                        f"正在下载更新... {percent}% ({done/1024/1024:.1f}MB/{total/1024/1024:.1f}MB)"
                    )
            else:
                if self._update_progress_label:
                    self._update_progress_label.setText(
                        f"正在下载更新... {done/1024/1024:.1f}MB"
                    )
        except Exception:
            pass

    def on_update_source_change(self, label: str):
        source = UPDATE_SOURCE_LABEL_TO_KEY.get(label, DEFAULT_UPDATE_SOURCE)
        self.settings["update_source"] = source
        save_settings(self.settings)

    def start_update_download(self):
        if self._update_download_thread:
            return
        info = self.update_info
        if not info:
            self._show_warning("未检测到可用更新")
            return
        if not is_frozen_app():
            self._show_warning("当前为源码运行，自动更新仅支持 EXE 版。")
            try:
                QtGui.QDesktopServices.openUrl(QtCore.QUrl(GITEE_REPO_URL))
            except Exception:
                pass
            return
        try:
            self.status_var.set("正在下载更新...")
        except Exception:
            pass
        self._show_update_progress_dialog()
        worker = UpdateDownloadWorker(info)
        thread = QtCore.QThread(self)
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.progress.connect(self._on_update_progress)
        worker.finished.connect(self._on_update_downloaded)
        worker.finished.connect(thread.quit)
        worker.finished.connect(worker.deleteLater)
        thread.finished.connect(thread.deleteLater)
        self._update_download_thread = thread
        self._update_download_worker = worker
        thread.start()

    def _on_update_downloaded(self, result, err: str):
        self._update_download_thread = None
        self._update_download_worker = None
        self._close_update_progress_dialog()
        if err or not result:
            self._show_error(f"更新下载失败: {err or '未知错误'}")
            return
        try:
            self.status_var.set("正在应用更新...")
        except Exception:
            pass
        self._run_update_script(result)

    def _run_update_script(self, result: dict):
        try:
            package_root = result.get("package_root") or ""
            exe_name = result.get("exe_name") or ""
            if not package_root or not exe_name:
                self._show_error("更新包信息无效")
                return
            exe_path = os.path.abspath(sys.executable)
            app_dir = os.path.dirname(exe_path)
            dst_dir = os.path.abspath(self._update_target_dir or app_dir)
            ps_path = os.path.join(get_update_cache_dir(), f"_mini_fish_update_{int(time.time())}.ps1")
            script = (
                "param([int]$Pid,[string]$SrcDir,[string]$DstDir,[string]$NewExe,[string]$OldExe,[string]$CacheDir)\n"
                "$wait = 0\n"
                "while (Get-Process -Id $Pid -ErrorAction SilentlyContinue) { Start-Sleep -Milliseconds 300; $wait++; if ($wait -ge 20) { break } }\n"
                "$proc = Get-Process -Id $Pid -ErrorAction SilentlyContinue\n"
                "if ($proc) { Stop-Process -Id $Pid -Force -ErrorAction SilentlyContinue }\n"
                "if ($OldExe) { $oldName = [System.IO.Path]::GetFileNameWithoutExtension($OldExe); if ($oldName) { Get-Process -Name $oldName -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue } }\n"
                "if (!(Test-Path $DstDir)) { New-Item -ItemType Directory -Force $DstDir | Out-Null }\n"
                "$oldDir = $DstDir\n"
                "if ($OldExe) { $oldDir = Split-Path $OldExe -Parent }\n"
                "$cacheDirs = @(\"_mini_fish_cache\",\"_mini_fish_icons\",\"_mini_fish_profile\")\n"
                "$cacheBases = @($DstDir, $oldDir) | Where-Object { $_ } | Select-Object -Unique\n"
                "foreach ($base in $cacheBases) { foreach ($d in $cacheDirs) { $p = Join-Path $base $d; if (Test-Path $p) { Remove-Item -Recurse -Force $p -ErrorAction SilentlyContinue } } }\n"
                "$internal = Join-Path $DstDir \"_internal\"\n"
                "if (Test-Path $internal) { Remove-Item -Recurse -Force $internal -ErrorAction SilentlyContinue }\n"
                "New-Item -ItemType Directory -Force $internal | Out-Null\n"
                "Copy-Item -Recurse -Force (Join-Path $SrcDir \"_internal\\*\") $internal\n"
                "$assetsSrc = Join-Path $SrcDir \"assets\"\n"
                "$assetsDst = Join-Path $DstDir \"assets\"\n"
                "if (Test-Path $assetsDst) { Remove-Item -Recurse -Force $assetsDst -ErrorAction SilentlyContinue }\n"
                "if (Test-Path $assetsSrc) { Copy-Item -Recurse -Force $assetsSrc $assetsDst }\n"
                "$srcExe = Join-Path $SrcDir $NewExe\n"
                "$newPath = Join-Path $DstDir $NewExe\n"
                "Copy-Item -Force $srcExe $newPath\n"
                "if ($OldExe -and (Test-Path $OldExe) -and ($OldExe -ne $newPath)) { Remove-Item -Force $OldExe -ErrorAction SilentlyContinue }\n"
                "$desktop = [Environment]::GetFolderPath('Desktop')\n"
                "if ($desktop) { $w = New-Object -ComObject WScript.Shell; Get-ChildItem -LiteralPath $desktop -Filter *.lnk | ForEach-Object { $lnk = $w.CreateShortcut($_.FullName); if ($lnk.TargetPath -ieq $OldExe -or $_.BaseName -like '*牛马神器*') { $lnk.TargetPath = $newPath; $lnk.WorkingDirectory = (Split-Path $newPath); $lnk.IconLocation = \"$newPath,0\"; $lnk.Save() } } }\n"
                "Start-Process $newPath\n"
                "$cmd = \"/c rmdir /s /q \\\"$CacheDir\\\"\"\n"
                "if ($CacheDir -and (Test-Path $CacheDir)) { Start-Process -FilePath cmd.exe -ArgumentList $cmd -WindowStyle Hidden }\n"
                "Remove-Item -Force $MyInvocation.MyCommand.Path -ErrorAction SilentlyContinue\n"
            )
            os.makedirs(os.path.dirname(ps_path), exist_ok=True)
            with open(ps_path, "w", encoding="utf-8-sig") as f:
                f.write(script)
            args = [
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                ps_path,
                "-Pid",
                str(os.getpid()),
                "-SrcDir",
                package_root,
                "-DstDir",
                dst_dir,
                "-NewExe",
                exe_name,
                "-OldExe",
                exe_path,
                "-CacheDir",
                get_update_cache_dir(),
            ]
            subprocess.Popen(args, creationflags=CREATE_NO_WINDOW)
            self._force_close = True
            self.close()
        except Exception as e:
            self._show_error(f"启动更新失败: {e}")

    def _get_hwnd(self):
        try:
            return int(self.winId())
        except Exception:
            return 0

    def _ahk_sanitize(self, text: str):
        return (text or "").replace("|", " ").replace("\r", " ").replace("\n", " ")

    def _hotkey_to_ahk(self, info):
        if not info or "error" in info:
            return ""
        mods = ""
        if info.get("mods", 0) & MOD_CONTROL:
            mods += "^"
        if info.get("mods", 0) & MOD_ALT:
            mods += "!"
        if info.get("mods", 0) & MOD_SHIFT:
            mods += "+"
        if info.get("mods", 0) & MOD_WIN:
            mods += "#"
        vk = info.get("vk")
        if isinstance(vk, int) and 0 <= vk <= 0xFF:
            key = f"vk{vk:02X}"
        else:
            key = str(info.get("keysym", "") or "").strip()
        return (mods + key).strip()

    def _ahk_payload(self):
        pt = self._ahk_sanitize(self.settings.get("panel_title") or "")
        bt = self._ahk_sanitize(self.settings.get("browser_title") or "")
        toggle_info = self._parse_hotkey(self.hotkey_toggle)
        lock_info = self._parse_hotkey(self.hotkey_lock)
        close_info = self._parse_hotkey(self.hotkey_close)
        hk_toggle = self._hotkey_to_ahk(toggle_info)
        hk_lock = self._hotkey_to_ahk(lock_info)
        hk_close = self._hotkey_to_ahk(close_info)
        return pt, bt, hk_toggle, hk_lock, hk_close

    def _ensure_ahk_running(self):
        if self._ahk_proc and self._ahk_proc.poll() is None:
            return True
        self._ahk_proc = None
        if not self.ahk_exe:
            self.ahk_exe = find_ahk_exe()
        if not self.ahk_exe:
            return False
        self._ahk_use_v2 = is_ahk_v2(self.ahk_exe)
        try:
            ensure_ahk_script(AHK_SCRIPT_PATH, use_v2=self._ahk_use_v2)
        except Exception:
            return False
        pt, bt, hk_toggle, hk_lock, hk_close = self._ahk_payload()
        try:
            cmd = [self.ahk_exe, AHK_SCRIPT_PATH, "daemon", pt, bt, hk_toggle, hk_lock, hk_close, AHK_CMD_PATH, AHK_EVT_PATH]
            self._ahk_proc = subprocess.Popen(cmd, creationflags=CREATE_NO_WINDOW)
            time.sleep(0.2)
            if self._ahk_proc.poll() is not None:
                self._ahk_proc = None
                return False
            return True
        except Exception:
            self._ahk_proc = None
            return False

    def _write_ahk_cmd(self, cmd: str):
        if not cmd:
            return False
        try:
            tmp_path = AHK_CMD_PATH + ".tmp"
            with open(tmp_path, "w", encoding="utf-8") as f:
                f.write(cmd)
            os.replace(tmp_path, AHK_CMD_PATH)
            return True
        except Exception:
            return False

    def _send_ahk_cmd(self, cmd: str):
        if not self._ensure_ahk_running():
            return False
        return self._write_ahk_cmd(cmd)

    def _sync_ahk_config(self):
        pt, bt, hk_toggle, hk_lock, hk_close = self._ahk_payload()
        cmd = f"update|{pt}|{bt}|{hk_toggle}|{hk_lock}|{hk_close}"
        return self._send_ahk_cmd(cmd)

    def _check_ahk_events(self):
        try:
            if not os.path.exists(AHK_EVT_PATH):
                return
            with open(AHK_EVT_PATH, "r", encoding="utf-8") as f:
                data = (f.read() or "").strip()
            try:
                os.remove(AHK_EVT_PATH)
            except Exception:
                pass
        except Exception:
            return
        if not data:
            return
        evt = data.split("|", 1)[0].strip().lower()
        if evt in ("top_on", "top_off"):
            new_state = evt == "top_on"
            try:
                if self.browser_top_checkbox.isChecked() != new_state:
                    self.browser_top_checkbox.blockSignals(True)
                    self.browser_top_checkbox.setChecked(new_state)
                    self.browser_top_checkbox.blockSignals(False)
                    self.settings["browser_topmost"] = new_state
                    save_settings(self.settings)
            except Exception:
                pass

    def _stop_ahk(self):
        if self._ahk_proc and self._ahk_proc.poll() is None:
            try:
                self._ahk_proc.terminate()
            except Exception:
                pass
        self._ahk_proc = None

    def _clear_local_shortcuts(self):
        for sc in self._local_shortcuts:
            try:
                sc.setEnabled(False)
            except Exception:
                pass
        self._local_shortcuts = []

    def _set_local_shortcut(self, seq: str, handler):
        try:
            shortcut = QtGui.QShortcut(QtGui.QKeySequence(seq), self)
            shortcut.setAutoRepeat(False)
            shortcut.activated.connect(handler)
            self._local_shortcuts.append(shortcut)
        except Exception:
            pass

    def _combo_data(self, combo: QtWidgets.QComboBox) -> str:
        data = combo.currentData()
        return data if data is not None else combo.currentText()

    def _set_combo_by_data(self, combo: QtWidgets.QComboBox, value: str):
        value = (value or "").strip()
        for i in range(combo.count()):
            if combo.itemData(i) == value:
                combo.setCurrentIndex(i)
                return
        idx = combo.findText(value)
        if idx >= 0:
            combo.setCurrentIndex(idx)

    def _populate_icon_combo(self, combo: QtWidgets.QComboBox, choices):
        combo.clear()
        for key in choices:
            combo.addItem(icon_display_name(key), key)

    def _show_topmost_invalid_hint(self):
        self._show_warning("该功能对您的系统失效，请检查设置")

    def nativeEvent(self, eventType, message):
        if eventType == "windows_generic_MSG":
            try:
                msg = wintypes.MSG.from_address(int(message))
                if msg.message == WM_HOTKEY:
                    if msg.wParam == 1:
                        self.minimize_all()
                    elif msg.wParam == 2:
                        self.restore_all()
                    elif msg.wParam == 3:
                        self.close_all()
                    return True, 0
            except Exception:
                pass
        return super().nativeEvent(eventType, message)

    def closeEvent(self, event):
        if self._force_close:
            self.on_close()
            event.accept()
            return
        if self._should_close_to_tray():
            self._close_to_tray()
            event.ignore()
            return
        self.on_close()
        event.accept()

    def dragEnterEvent(self, event):
        try:
            if event.mimeData().hasUrls():
                url = event.mimeData().urls()[0]
                path = url.toLocalFile()
                if is_icon_file(path):
                    event.acceptProposedAction()
                    return
        except Exception:
            pass
        event.ignore()

    def dropEvent(self, event):
        try:
            urls = event.mimeData().urls()
            if not urls:
                return
            files = [u.toLocalFile() for u in urls if u.isLocalFile()]
            if files:
                self.on_drop_files(files)
                event.acceptProposedAction()
        except Exception:
            pass

# ---------- presets ----------
    def pick_preset(self):
        u = normalize_url(self.preset_var.get())
        self.url_var.set(u)
        self.go()

    def add_preset(self):
        u = normalize_url(self.url_var.get())
        if u not in self.settings["presets"]:
            self.settings["presets"].insert(0, u)
            self.preset_combo.blockSignals(True)
            self.preset_combo.clear()
            self.preset_combo.addItems(self.settings["presets"])
            self.preset_combo.blockSignals(False)
            save_settings(self.settings)
        self.preset_var.set(u)

    def remove_preset(self):
        u = normalize_url(self.preset_var.get() or self.url_var.get())
        if u in self.settings["presets"]:
            self.settings["presets"].remove(u)
            self.preset_combo.blockSignals(True)
            self.preset_combo.clear()
            self.preset_combo.addItems(self.settings["presets"])
            self.preset_combo.blockSignals(False)
            save_settings(self.settings)
        self.preset_var.set("")

    def open_new_instance(self):
        try:
            inst = f"{os.getpid()}_{int(time.time() * 1000)}"
            subprocess.Popen([sys.executable, os.path.abspath(__file__), f"--instance-id={inst}"])
        except Exception as e:
            self._show_error(f"无法打开新窗口: {e}")

    def open_github(self):
        self.open_source_dialog()

    def open_source_dialog(self):
        try:
            win = QtWidgets.QDialog(self)
            win.setWindowTitle("源码地址")
            win.setModal(False)
            layout = QtWidgets.QVBoxLayout(win)

            tip = QtWidgets.QLabel("请选择访问来源：")
            layout.addWidget(tip)

            row = QtWidgets.QHBoxLayout()
            btn_github = QtWidgets.QPushButton("GitHub（国外）")
            btn_gitee = QtWidgets.QPushButton("Gitee（国内）")
            row.addWidget(btn_github)
            row.addWidget(btn_gitee)
            layout.addLayout(row)

            btns = QtWidgets.QHBoxLayout()
            btns.addStretch(1)
            close_btn = QtWidgets.QPushButton("关闭")
            btns.addWidget(close_btn)
            layout.addLayout(btns)

            btn_github.clicked.connect(lambda: QtGui.QDesktopServices.openUrl(QtCore.QUrl(GITHUB_REPO_URL)))
            btn_gitee.clicked.connect(lambda: QtGui.QDesktopServices.openUrl(QtCore.QUrl(GITEE_REPO_URL)))
            close_btn.clicked.connect(win.close)

            win.adjustSize()
            self._place_dialog_next_to_panel(win)
            win.show()
        except Exception:
            pass

    def open_about_dialog(self):
        try:
            win = QtWidgets.QDialog(self)
            win.setWindowTitle("关于")
            win.setModal(False)
            layout = QtWidgets.QVBoxLayout(win)

            title = QtWidgets.QLabel(f"{APP_TITLE}")
            title_font = title.font()
            title_font.setPointSize(16)
            title_font.setBold(True)
            title.setFont(title_font)
            title.setAlignment(QtCore.Qt.AlignCenter)
            layout.addWidget(title)

            info = QtWidgets.QLabel()
            info.setWordWrap(True)
            info.setTextFormat(QtCore.Qt.RichText)
            info.setOpenExternalLinks(True)
            info.setTextInteractionFlags(QtCore.Qt.TextBrowserInteraction)
            info.setText(
                f"<b>版本</b>: {APP_VERSION}<br>"
                f"<b>运行环境</b>: Windows 10/11 · Python 3.8+ · Chrome/Edge<br>"
                f"<b>配置文件</b>: {os.path.basename(SETTINGS_PATH)}<br>"
                f"<b>数据目录</b>: {os.path.basename(PROFILE_DIR_BASE)}<br>"
                f"<b>许可</b>: MIT<br>"
                f"<b>在线更新</b>: 已支持自动更新（Gitee 源）<br>"
                f"<b>GitHub 主页</b>: <a href=\"{GITHUB_HOME_URL}\">{GITHUB_HOME_URL}</a><br>"
                f"<b>Gitee 主页</b>: <a href=\"{GITEE_HOME_URL}\">{GITEE_HOME_URL}</a><br>"
                f"<b>脚本仓库（GitHub）</b>: <a href=\"{GITHUB_REPO_URL}\">{GITHUB_REPO_URL}</a><br>"
                f"<b>脚本仓库（Gitee）</b>: <a href=\"{GITEE_REPO_URL}\">{GITEE_REPO_URL}</a><br>"
                f"<b>联系作者</b>: 1870511741@qq.com / jerrychencnfirst@gmail.com<br>"
                f"<b>微信</b>: 见下方二维码"
            )
            layout.addWidget(info)

            update_row = QtWidgets.QHBoxLayout()
            update_row.setContentsMargins(0, 0, 0, 0)
            update_label = QtWidgets.QLabel()
            update_label.setWordWrap(True)
            if self._update_check_thread:
                update_text = "正在检查更新..."
            elif self.update_available and self.update_info:
                update_text = f"发现新版本 {self.update_info.get('version', '')}"
                update_label.setStyleSheet("color: #ef4444;")
                dot = QtWidgets.QLabel()
                dot.setFixedSize(8, 8)
                dot.setStyleSheet("background-color: #ef4444; border-radius: 4px;")
                update_row.addWidget(dot)
            elif self._update_checked_once:
                update_text = "已是最新版本"
            else:
                update_text = "未检查更新"
            update_label.setText(update_text)
            update_row.addWidget(update_label)
            update_row.addStretch(1)
            btn_check_update = QtWidgets.QPushButton("检查更新")
            btn_check_update.clicked.connect(lambda: self.check_update_async(force=True))
            update_row.addWidget(btn_check_update)
            if self.update_available and self.update_info:
                btn_update_now = QtWidgets.QPushButton("立即更新")
                def _do_update_now():
                    win.close()
                    self.start_update_download()
                btn_update_now.clicked.connect(_do_update_now)
                update_row.addWidget(btn_update_now)
            layout.addLayout(update_row)

            source_row = QtWidgets.QHBoxLayout()
            source_row.setContentsMargins(0, 0, 0, 0)
            source_row.addWidget(QtWidgets.QLabel("更新源"))
            source_combo = QtWidgets.QComboBox()
            source_combo.addItems(list(UPDATE_SOURCE_LABELS.values()))
            current_source = self.settings.get("update_source", DEFAULT_UPDATE_SOURCE)
            source_combo.setCurrentText(UPDATE_SOURCE_LABELS.get(current_source, UPDATE_SOURCE_LABELS[DEFAULT_UPDATE_SOURCE]))
            source_combo.currentTextChanged.connect(self.on_update_source_change)
            source_row.addWidget(source_combo)
            source_row.addStretch(1)
            layout.addLayout(source_row)

            wx_title = QtWidgets.QLabel("微信联系（扫码添加）")
            wx_title.setAlignment(QtCore.Qt.AlignCenter)
            layout.addWidget(wx_title)
            wx_img = QtWidgets.QLabel()
            wx_img.setAlignment(QtCore.Qt.AlignCenter)
            pix = load_any_image(os.path.join(ASSETS_DIR, "weixin.png"), max_size=(240, 240))
            if pix:
                wx_img.setPixmap(pix)
            else:
                wx_img.setText("weixin.png 未找到")
            layout.addWidget(wx_img)

            btns = QtWidgets.QHBoxLayout()
            btns.addStretch(1)
            close_btn = QtWidgets.QPushButton("关闭")
            btns.addWidget(close_btn)
            layout.addLayout(btns)
            close_btn.clicked.connect(win.close)

            win.adjustSize()
            self._place_dialog_next_to_panel(win)
            win.show()
        except Exception as e:
            self._show_error(f"关于窗口打开失败: {e}")

    def open_sponsor_dialog(self):
        try:
            win = QtWidgets.QDialog(self)
            win.setWindowTitle("赞助作者")
            win.setModal(False)
            layout = QtWidgets.QVBoxLayout(win)

            text_label = QtWidgets.QLabel("")
            text_label.setWordWrap(True)
            text_label.setAlignment(QtCore.Qt.AlignCenter)
            text_label.setTextFormat(QtCore.Qt.RichText)
            text_font = text_label.font()
            text_font.setPointSize(14)
            text_label.setFont(text_font)
            layout.addWidget(text_label)

            img_row = QtWidgets.QHBoxLayout()
            cells = []
            for _ in range(2):
                cell = QtWidgets.QVBoxLayout()
                img_label = QtWidgets.QLabel()
                img_label.setAlignment(QtCore.Qt.AlignCenter)
                name_label = QtWidgets.QLabel("")
                name_label.setAlignment(QtCore.Qt.AlignCenter)
                cell.addWidget(img_label)
                cell.addWidget(name_label)
                img_row.addLayout(cell)
                cells.append((img_label, name_label))
            layout.addLayout(img_row)

            btns = QtWidgets.QHBoxLayout()
            btns.addStretch(1)
            close_hint = QtWidgets.QLabel("")
            close_hint.setVisible(False)
            close_hint.setWordWrap(False)
            close_hint.setAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
            close_hint.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
            hint_font = close_hint.font()
            hint_font.setPointSize(18)
            close_hint.setFont(hint_font)
            close_hint.setStyleSheet("color: #d00000;")
            toggle_btn = QtWidgets.QPushButton("穷鬼入口")
            close_btn = QtWidgets.QPushButton("关闭")
            btns.addWidget(close_hint)
            btns.addWidget(toggle_btn)
            btns.addWidget(close_btn)
            layout.addLayout(btns)

            state = {"mode": "sponsor"}
            def apply_mode(mode: str):
                if mode == "poor":
                    text_label.setText(POOR_TEXT)
                    items = POOR_QR_FILES
                    toggle_btn.setText("返回赞助")
                    win.setWindowTitle("穷鬼入口")
                else:
                    sponsor_html = SPONSOR_TEXT.replace(
                        "随意",
                        "<span style='color:#d00000;font-size:20pt;'>随意</span>",
                        1,
                    )
                    text_label.setText(sponsor_html)
                    items = SPONSOR_QR_FILES
                    toggle_btn.setText("穷鬼入口")
                    win.setWindowTitle("赞助作者")
                for idx, (label, path) in enumerate(items):
                    img_label, name_label = cells[idx]
                    pix = load_any_image(path, max_size=(360, 360))
                    if pix:
                        img_label.setPixmap(pix)
                        img_label.setText("")
                    else:
                        img_label.setPixmap(QtGui.QPixmap())
                        img_label.setText("二维码加载失败")
                    name_label.setText(label)
                win.adjustSize()
                self._place_dialog_next_to_panel(win)

            def toggle_mode():
                state["mode"] = "poor" if state["mode"] == "sponsor" else "sponsor"
                apply_mode(state["mode"])

            toggle_btn.clicked.connect(toggle_mode)
            close_msgs = [
                "唉别走呀，这么不要面子的吗你→",
                "不是我说你这家伙，你今年想不想发财了！！！",
                "客官行行好吧，三天天没开张了快饿死了！！！",
                "恭喜你通过考验，刚才我是装的，我怎么可能求你呢，放你走吧",
            ]
            close_state = {"count": 0, "allow_close": False}
            def advance_close_message():
                if close_state.get("allow_close"):
                    return True
                close_state["count"] += 1
                idx = min(close_state["count"], len(close_msgs)) - 1
                close_hint.setText(close_msgs[idx])
                close_hint.setVisible(True)
                win.adjustSize()
                self._place_dialog_next_to_panel(win)
                if close_state["count"] >= len(close_msgs):
                    close_state["allow_close"] = True
                return False

            def on_close_click():
                if advance_close_message():
                    win.close()
            close_btn.clicked.connect(on_close_click)

            def on_close_event(event):
                if advance_close_message():
                    event.accept()
                else:
                    event.ignore()
            win.closeEvent = on_close_event
            apply_mode("sponsor")
            win.show()
        except Exception as e:
            self._show_error(f"赞助窗口打开失败: {e}")

    # ---------- panel icon ----------
    def _icon_title_from_style(self, style: str, custom_path: str = "") -> str:
        style = (style or "").strip()
        if style == "custom":
            base = os.path.splitext(os.path.basename(custom_path or ""))[0].strip()
            if base:
                return base
        return icon_display_name(style)

    def _auto_set_panel_title_from_icon(self, style: str):
        title = self._icon_title_from_style(style, self.settings.get("panel_custom_icon_path", ""))
        if title:
            self.panel_title_var.set(title)
            self.apply_titles(save=True)

    def _auto_set_browser_title_from_icon(self, style: str):
        title = self._icon_title_from_style(style, self.settings.get("browser_custom_icon_path", ""))
        if title:
            self.browser_title_var.set(title)
            self.apply_titles(save=True)

    def apply_panel_icon_style(self, auto_rename: bool = False):
        style = self.panel_icon_style_var.get().strip() or "globe"
        self.settings["panel_icon_style"] = style
        save_settings(self.settings)
        try:
            applied = False
            if style == "custom":
                path = (self.settings.get("panel_custom_icon_path") or "").strip()
                pix = load_icon_image(path, max_size=32)
                if pix:
                    self.panel_icon_pixmap = pix
                    self.setWindowIcon(QtGui.QIcon(pix))
                    applied = True

            if not applied and style in ICON_FILES:
                pix = load_icon_image(ICON_FILES[style], max_size=32)
                if pix:
                    self.panel_icon_pixmap = pix
                    self.setWindowIcon(QtGui.QIcon(pix))
                    applied = True

            if not applied:
                base_style = style if style in GENERIC_ICON_STYLES else "globe"
                pix = make_icon(base_style)
                self.panel_icon_pixmap = pix
                self.setWindowIcon(QtGui.QIcon(pix))
        except Exception:
            pass
        if auto_rename:
            self._auto_set_panel_title_from_icon(style)
        if not self._initializing and self.sync_taskbar_icon:
            self.update_taskbar_icon()
        self._update_tray_icons()

    def apply_browser_icon_style(self, auto_rename: bool = False):
        style = self.browser_icon_style_var.get().strip() or "site"
        self.settings["browser_icon_style"] = style
        save_settings(self.settings)
        self.refresh_browser_icon_data()
        self.apply_browser_icon()
        if auto_rename:
            self._auto_set_browser_title_from_icon(style)
        self._update_tray_icons()

    def refresh_browser_icon_data(self):
        style = self.browser_icon_style_var.get().strip() or "site"
        if style == "site":
            self.browser_icon_data_url = ""
            self.browser_icon_mime = ""
            return
        path = ""
        if style == "custom":
            path = (self.settings.get("browser_custom_icon_path") or "").strip()
        elif style in ICON_FILES:
            path = ICON_FILES.get(style, "")
        if path and os.path.exists(path) and is_icon_file(path):
            data_url, mime = icon_data_url_from_path(path)
            self.browser_icon_data_url = data_url
            self.browser_icon_mime = mime
        else:
            self.browser_icon_data_url = ""
            self.browser_icon_mime = ""

    def apply_browser_icon(self):
        if not self.driver:
            return
        if not self.browser_icon_data_url:
            return
        self._ensure_driver_window()
        js = r"""
        (function(href, mime){
          try {
            var link = document.querySelector("link[rel~='icon']");
            if (!link) {
              link = document.createElement("link");
              link.rel = "icon";
              var head = document.head || document.getElementsByTagName("head")[0] || document.documentElement;
              if (head) { head.appendChild(link); }
            }
            if (mime) link.type = mime;
            link.href = href;
          } catch(e) {}
        })(arguments[0], arguments[1]);
        """
        try:
            self.driver.execute_script(js, self.browser_icon_data_url, self.browser_icon_mime)
        except Exception:
            pass

    def on_taskbar_sync_toggle(self, checked: bool):
        self.sync_taskbar_icon = bool(checked)
        self.settings["sync_taskbar_icon"] = self.sync_taskbar_icon
        save_settings(self.settings)
        if self.sync_taskbar_icon:
            self.update_taskbar_icon()

    def on_panel_tray_toggle(self, checked: bool):
        self.panel_tray_enabled = bool(checked)
        self.settings["panel_tray_enabled"] = self.panel_tray_enabled
        save_settings(self.settings)
        self._sync_tray_icons()

    def on_browser_tray_toggle(self, checked: bool):
        self.browser_tray_enabled = bool(checked)
        self.settings["browser_tray_enabled"] = self.browser_tray_enabled
        save_settings(self.settings)
        self._sync_tray_icons()

    def _tray_available(self) -> bool:
        try:
            return QtWidgets.QSystemTrayIcon.isSystemTrayAvailable()
        except Exception:
            return False

    def _panel_tray_icon(self) -> QtGui.QIcon:
        if self.panel_icon_pixmap:
            return QtGui.QIcon(self.panel_icon_pixmap)
        style = (self.panel_icon_style_var.get() or "globe").strip()
        if style in ICON_FILES:
            return QtGui.QIcon(ICON_FILES[style])
        return QtGui.QIcon(make_icon(style if style in GENERIC_ICON_STYLES else "globe"))

    def _browser_tray_icon(self) -> QtGui.QIcon:
        style = (self.browser_icon_style_var.get() or "site").strip()
        if style == "site" and self.site_icon_pixmap:
            return QtGui.QIcon(self.site_icon_pixmap)
        path = ""
        if style == "custom":
            path = (self.settings.get("browser_custom_icon_path") or "").strip()
        elif style in ICON_FILES:
            path = ICON_FILES.get(style, "")
        if path and os.path.exists(path):
            return QtGui.QIcon(path)
        if self.panel_icon_pixmap:
            return QtGui.QIcon(self.panel_icon_pixmap)
        return QtGui.QIcon(make_icon("globe"))

    def _sync_tray_icons(self):
        if not self._tray_available():
            if self.panel_tray_enabled or self.browser_tray_enabled:
                self._show_warning("系统托盘不可用，无法显示驻留图标")
            self._destroy_tray_icons()
            return

        if self.panel_tray_enabled and not self.panel_tray:
            self.panel_tray = QtWidgets.QSystemTrayIcon(self._panel_tray_icon(), self)
            self.panel_tray.setToolTip(self.settings.get("panel_title") or APP_TITLE)
            self.panel_tray.activated.connect(self._on_panel_tray_activated)
            menu = QtWidgets.QMenu()
            act_show = menu.addAction("显示面板")
            act_hide = menu.addAction("隐藏面板")
            act_exit = menu.addAction("退出")
            act_show.triggered.connect(self._show_panel_from_tray)
            act_hide.triggered.connect(self._hide_panel_from_tray)
            act_exit.triggered.connect(self.request_exit)
            self.panel_tray.setContextMenu(menu)
            self.panel_tray.show()
        elif not self.panel_tray_enabled and self.panel_tray:
            self.panel_tray.hide()
            self.panel_tray.deleteLater()
            self.panel_tray = None

        if self.browser_tray_enabled and not self.browser_tray:
            self.browser_tray = QtWidgets.QSystemTrayIcon(self._browser_tray_icon(), self)
            self.browser_tray.setToolTip(self.settings.get("browser_title") or "mini-browser")
            self.browser_tray.activated.connect(self._on_browser_tray_activated)
            menu = QtWidgets.QMenu()
            act_show = menu.addAction("找回浏览器")
            act_restart = menu.addAction("重启浏览器")
            act_close = menu.addAction("关闭浏览器")
            act_show.triggered.connect(self.recover_browser)
            act_restart.triggered.connect(self.restart_browser)
            act_close.triggered.connect(self.close_browser_only)
            self.browser_tray.setContextMenu(menu)
            self.browser_tray.show()
        elif not self.browser_tray_enabled and self.browser_tray:
            self.browser_tray.hide()
            self.browser_tray.deleteLater()
            self.browser_tray = None

        self._update_tray_icons()

    def _update_tray_icons(self):
        if self.panel_tray:
            self.panel_tray.setIcon(self._panel_tray_icon())
            self.panel_tray.setToolTip(self.settings.get("panel_title") or APP_TITLE)
        if self.browser_tray:
            self.browser_tray.setIcon(self._browser_tray_icon())
            self.browser_tray.setToolTip(self.settings.get("browser_title") or "mini-browser")

    def _destroy_tray_icons(self):
        if self.panel_tray:
            self.panel_tray.hide()
            self.panel_tray.deleteLater()
            self.panel_tray = None
        if self.browser_tray:
            self.browser_tray.hide()
            self.browser_tray.deleteLater()
            self.browser_tray = None

    def _should_close_to_tray(self) -> bool:
        return self._tray_available() and (self.panel_tray_enabled or self.browser_tray_enabled)

    def _hide_browser_window(self):
        try:
            self.ensure_chrome_hwnd()
            hwnds = self.get_browser_hwnds(include_all=False)
            if self.chrome_hwnd:
                hwnds.insert(0, self.chrome_hwnd)
            seen = set()
            for hwnd in hwnds:
                if hwnd and hwnd not in seen:
                    hide_window(hwnd)
                    seen.add(hwnd)
        except Exception:
            pass

    def _close_to_tray(self):
        try:
            self.hide()
        except Exception:
            pass
        if self.browser_tray_enabled:
            self._hide_browser_window()

    def _show_panel_from_tray(self):
        try:
            self.showNormal()
            self.raise_()
            self.activateWindow()
        except Exception:
            pass
        if self.browser_tray_enabled:
            self._restore_browser_window(silent=True)

    def _hide_panel_from_tray(self):
        try:
            self.showMinimized()
        except Exception:
            pass

    def _on_panel_tray_activated(self, reason):
        if reason == QtWidgets.QSystemTrayIcon.Trigger:
            self._show_panel_from_tray()

    def _on_browser_tray_activated(self, reason):
        if reason == QtWidgets.QSystemTrayIcon.Trigger:
            self._restore_browser_window(silent=True)

    def close_browser_only(self):
        try:
            self.safe_quit_driver()
            self.safe_kill_proc()
            self.port = None
            self.chrome_hwnd = None
            self.main_window_handle = None
            self._stop_attach_timer()
            self.status_var.set("浏览器已关闭")
        except Exception:
            pass

    def _panel_icon_source_path(self) -> str:
        style = (self.panel_icon_style_var.get() or "").strip() or "globe"
        if style == "custom":
            return (self.settings.get("panel_custom_icon_path") or "").strip()
        if style in ICON_FILES:
            return ICON_FILES.get(style, "") or ""
        base_style = style if style in GENERIC_ICON_STYLES else "globe"
        try:
            os.makedirs(ICON_DIR, exist_ok=True)
            pix = make_icon(base_style)
            out_path = os.path.join(ICON_DIR, f"_generic_{base_style}.png")
            pix.save(out_path, "PNG")
            return out_path
        except Exception:
            return ""

    def _ensure_taskbar_ico(self, source_path: str) -> str:
        if not source_path or not os.path.exists(source_path):
            return ""
        ext = os.path.splitext(source_path)[1].lower()
        if ext == ".ico":
            return source_path
        if not ensure_pillow():
            return ""
        try:
            from PIL import Image
            os.makedirs(ICON_DIR, exist_ok=True)
            base = os.path.splitext(os.path.basename(source_path))[0]
            safe = re.sub(r"[^A-Za-z0-9_.-]+", "_", base) or "icon"
            out_path = os.path.join(ICON_DIR, f"{safe}_taskbar.ico")
            img = Image.open(source_path)
            img.save(out_path, format="ICO", sizes=[(256, 256)])
            return out_path if os.path.exists(out_path) else ""
        except Exception:
            return ""

    def _update_taskbar_shortcut(self, exe_path: str, icon_path: str) -> bool:
        if not exe_path or not icon_path:
            return False
        ps_exe = shutil.which("pwsh") or shutil.which("powershell")
        if not ps_exe:
            return False
        def ps_quote(text: str) -> str:
            return "'" + text.replace("'", "''") + "'"
        script = "\n".join([
            "$ErrorActionPreference = 'Stop'",
            f"$exe = {ps_quote(exe_path)}",
            f"$icon = {ps_quote(icon_path)}",
            "$dir = Join-Path $env:APPDATA 'Microsoft\\Internet Explorer\\Quick Launch\\User Pinned\\TaskBar'",
            "if (!(Test-Path -LiteralPath $dir)) { exit 3 }",
            "$w = New-Object -ComObject WScript.Shell",
            "$found = $false",
            "Get-ChildItem -LiteralPath $dir -Filter *.lnk | ForEach-Object {",
            "  $lnk = $w.CreateShortcut($_.FullName)",
            "  if ($lnk.TargetPath -ieq $exe) {",
            "    $lnk.IconLocation = \"$icon,0\"",
            "    $lnk.Save()",
            "    $found = $true",
            "  }",
            "}",
            "if (-not $found) { exit 2 }",
        ])
        try:
            proc = subprocess.run(
                [ps_exe, "-NoProfile", "-Command", script],
                capture_output=True,
                text=True,
                creationflags=CREATE_NO_WINDOW,
            )
        except Exception:
            return False
        if proc.returncode != 0:
            return False
        try:
            ctypes.windll.shell32.SHChangeNotify(0x8000000, 0, None, None)
        except Exception:
            pass
        return True

    def update_taskbar_icon(self):
        if not getattr(sys, "frozen", False):
            self._show_warning("仅 EXE 版可同步任务栏固定图标")
            return
        exe_path = os.path.abspath(sys.executable)
        src_path = self._panel_icon_source_path()
        icon_path = self._ensure_taskbar_ico(src_path)
        if not icon_path:
            self._show_warning("未找到可用 .ico 图标，请选择 .ico 图标或切换图标库")
            return
        ok = self._update_taskbar_shortcut(exe_path, icon_path)
        if not ok:
            self._show_warning("未找到任务栏固定快捷方式，请先固定到任务栏")
            return
        try:
            self.status_var.set("任务栏图标已同步")
        except Exception:
            pass

    def set_panel_icon_path(self, path: str):
        if not path:
            return
        path = path.strip()
        if not os.path.isfile(path):
            return
        if not is_icon_file(path):
            self._show_warning("仅支持 png/gif/ico 图标文件")
            return
        self.settings["panel_custom_icon_path"] = path
        self.settings["panel_icon_style"] = "custom"
        save_settings(self.settings)
        try:
            self.panel_icon_style_combo.blockSignals(True)
            self.panel_icon_style_var.set("custom")
        finally:
            self.panel_icon_style_combo.blockSignals(False)
        self.apply_panel_icon_style(auto_rename=True)

    def set_browser_icon_path(self, path: str):
        if not path:
            return
        path = path.strip()
        if not os.path.isfile(path):
            return
        if not is_icon_file(path):
            self._show_warning("仅支持 png/gif/ico 图标文件")
            return
        self.settings["browser_custom_icon_path"] = path
        self.settings["browser_icon_style"] = "custom"
        save_settings(self.settings)
        try:
            self.browser_icon_style_combo.blockSignals(True)
            self.browser_icon_style_var.set("custom")
        finally:
            self.browser_icon_style_combo.blockSignals(False)
        self.apply_browser_icon_style(auto_rename=True)

    def choose_panel_icon(self):
        types = "Image (*.png *.gif *.ico);;PNG (*.png);;GIF (*.gif);;ICO (*.ico);;All (*.*)"
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "选择图标", "", types)
        if not path:
            return
        self.set_panel_icon_path(path)

    def choose_browser_icon(self):
        types = "Image (*.png *.gif *.ico);;PNG (*.png);;GIF (*.gif);;ICO (*.ico);;All (*.*)"
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "选择图标", "", types)
        if not path:
            return
        self.set_browser_icon_path(path)

    def setup_icon_drop(self):
        self.setAcceptDrops(True)

    def on_drop_files(self, files):
        if not files:
            return
        path = files[0]
        if isinstance(path, bytes):
            path = path.decode(sys.getfilesystemencoding(), errors="ignore")
        browser_is_custom = (self.browser_icon_style_var.get() == "custom")
        panel_is_custom = (self.panel_icon_style_var.get() == "custom")
        if browser_is_custom and not panel_is_custom:
            self.set_browser_icon_path(path)
        else:
            self.set_panel_icon_path(path)

    # ---------- titles ----------
    def apply_titles(self, save: bool = True):
        pt = (self.panel_title_var.get() or "").strip() or "mini"
        bt = (self.browser_title_var.get() or "").strip() or "mini-browser"

        if save:
            self.settings["panel_title"] = pt
            self.settings["browser_title"] = bt
            save_settings(self.settings)

        try:
            self.setWindowTitle(pt)
        except Exception:
            pass

        self.apply_browser_title()
        try:
            self._sync_ahk_config()
        except Exception:
            pass
        self._update_tray_icons()

    def apply_browser_title(self):
        bt = (self.settings.get("browser_title") or "").strip()
        if not bt:
            return

        if self.driver:
            self._ensure_driver_window()
            js = r"""
            (function(name){
              try { document.title = name; } catch(e){}
              try {
                if (!window.__miniFishTitleLock) {
                  window.__miniFishTitleLock = setInterval(function(){
                    try { document.title = name; } catch(e){}
                  }, 1200);
                }
              } catch(e){}
            })(arguments[0]);
            """
            try:
                self.driver.execute_script(js, bt)
            except Exception:
                pass

        try:
            self.ensure_chrome_hwnd()
            if self.chrome_hwnd:
                set_window_title(self.chrome_hwnd, bt)
        except Exception:
            pass

    def apply_titles_and_refresh(self):
        self.apply_titles(save=True)
        self.manual_refresh()

    # ---------- browser lifecycle ----------
    def _stop_attach_timer(self):
        if self._attach_timer:
            try:
                self._attach_timer.stop()
            except Exception:
                pass
            try:
                self._attach_timer.deleteLater()
            except Exception:
                pass
            self._attach_timer = None

    def _start_attach_timer(self):
        if self.driver:
            return
        if self._attach_timer:
            return
        if not self.port:
            alt_port = read_devtools_port(self.profile_dir)
            if alt_port:
                self.port = alt_port
        self._attach_retry = 0
        self._attach_timer = QtCore.QTimer(self)
        self._attach_timer.timeout.connect(self._try_attach_driver)
        self._attach_timer.start(1000)

    def _try_attach_driver(self):
        if self.driver:
            self._stop_attach_timer()
            return
        if not self._is_browser_alive(check_driver=False):
            self._stop_attach_timer()
            try:
                self.status_var.set("浏览器已关闭，点击Go重新启动")
            except Exception:
                pass
            return
        if not self.port:
            alt_port = read_devtools_port(self.profile_dir)
            if alt_port:
                self.port = alt_port
        if not self.port:
            self._attach_retry += 1
            if self._attach_retry >= 20:
                self._stop_attach_timer()
                try:
                    self.status_var.set("连接失败，请点击“重启浏览器”")
                except Exception:
                    pass
            return
        if not is_debug_port_ready(self.port):
            alt_port = read_devtools_port(self.profile_dir)
            if alt_port and alt_port != self.port:
                self.port = alt_port
            if not is_debug_port_ready(self.port):
                self._attach_retry += 1
                if self._attach_retry >= 20:
                    self._stop_attach_timer()
                    try:
                        self.status_var.set("连接失败，请点击“重启浏览器”")
                    except Exception:
                        pass
                return
        try:
            self.driver = attach_selenium(self.port)
            try:
                self.main_window_handle = self.driver.current_window_handle
            except Exception:
                self.main_window_handle = None
            try:
                self.driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {"source": PREVENT_NEW_WINDOWS_JS})
            except Exception:
                pass
            try:
                self.driver.execute_cdp_cmd("Page.setWindowOpenHandler", {"handler": "deny"})
            except Exception:
                pass
            self._stop_attach_timer()
            try:
                self.status_var.set("浏览器已连接")
            except Exception:
                pass
            self._finalize_browser_start()
        except Exception:
            self._attach_retry += 1
    def _has_browser_window(self) -> bool:
        try:
            return bool(self.get_browser_hwnds(include_all=False))
        except Exception:
            return False

    def _is_browser_alive(self, check_driver: bool = True) -> bool:
        if check_driver and self.driver:
            try:
                _ = self.driver.title
                return True
            except Exception:
                pass
        if self.port and is_debug_port_ready(self.port):
            return True
        alt_port = read_devtools_port(self.profile_dir)
        if alt_port and is_debug_port_ready(alt_port):
            self.port = alt_port
            return True
        if self._has_browser_window():
            return True
        return False

    def _ensure_driver_window(self, expected_url: str = ""):
        if not self.driver:
            return
        try:
            if self.main_window_handle:
                self.driver.switch_to.window(self.main_window_handle)
            else:
                self._select_main_window_handle(expected_url=expected_url)
        except Exception:
            pass

    def start_browser_safe(self, profile_dir: str = "", url: str = ""):
        self._stop_attach_timer()
        self.start_browser_async(profile_dir=profile_dir, url=url)

    def _start_browser_blocking(self, profile_dir: str = "", url: str = ""):
        url = normalize_url(url or self.url_var.get())
        self.url_var.set(url)
        self.profile_dir = profile_dir or self.profile_dir or PROFILE_DIR
        x, y = self.calc_browser_position(self.win_w, self.win_h)
        proc, driver, port, err = launch_browser_session(
            self.chrome_exe, url, self.win_w, self.win_h, x, y, self.profile_dir,
            self.audio_enabled
        )
        if err and not driver:
            raise RuntimeError(err)
        self.proc, self.driver, self.port = proc, driver, port

        try:
            self.main_window_handle = self.driver.current_window_handle
        except Exception:
            self.main_window_handle = None

        # inject "prevent new windows" for all future navigations
        try:
            self.driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {"source": PREVENT_NEW_WINDOWS_JS})
        except Exception:
            pass
        try:
            self.driver.execute_cdp_cmd("Page.setWindowOpenHandler", {"handler": "deny"})
        except Exception:
            pass

        time.sleep(0.6)
        self._select_main_window_handle()
        self.ensure_chrome_hwnd(force=True)
        self.apply_taskbar_merge()
        self.sync_attach_positions(force=True)

        self.apply_browser_window_size(resize_now=True)
        self.apply_zoom()
        self.apply_alpha()
        self.apply_browser_topmost()
        self.apply_browser_title()
        self.apply_browser_icon()
        self.refresh_status(force_icon=True)
        self._apply_audio_state(silent=True)
        self._apply_audio_state(silent=True)

    def start_browser_async(self, profile_dir: str = "", url: str = ""):
        if self._starting_browser:
            return
        url = normalize_url(url or self.url_var.get())
        self.url_var.set(url)
        self.profile_dir = profile_dir or self.profile_dir or PROFILE_DIR
        x, y = self.calc_browser_position(self.win_w, self.win_h)
        self._starting_browser = True
        try:
            self.status_var.set("正在启动浏览器...")
        except Exception:
            pass

        worker = BrowserStartWorker(
            self.chrome_exe, url, self.win_w, self.win_h, x, y, self.profile_dir, self.audio_enabled
        )
        thread = QtCore.QThread(self)
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.finished.connect(lambda proc, driver, port, err: self._on_browser_started(proc, driver, port, err))
        worker.finished.connect(thread.quit)
        worker.finished.connect(worker.deleteLater)
        thread.finished.connect(thread.deleteLater)
        self._browser_start_thread = thread
        self._browser_start_worker = worker
        thread.start()

    def _on_browser_started(self, proc, driver, port, err: str):
        self._starting_browser = False
        self._browser_start_thread = None
        self._browser_start_worker = None
        if err or not driver:
            try:
                self._stop_attach_timer()
                self.safe_quit_driver()
                self._kill_proc_and_profile(proc, self.profile_dir)
            except Exception:
                pass
            self.driver = None
            self.proc = None
            self.port = None
            self.chrome_hwnd = None
            self.main_window_handle = None
            try:
                self.status_var.set(f"启动失败: {err or '浏览器未就绪'}")
            except Exception:
                pass
            self._show_error(err or "浏览器未就绪")
            return

        self.proc = proc
        self.driver = driver
        self.port = port
        try:
            self.main_window_handle = self.driver.current_window_handle
        except Exception:
            self.main_window_handle = None

        try:
            self.driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {"source": PREVENT_NEW_WINDOWS_JS})
        except Exception:
            pass
        try:
            self.driver.execute_cdp_cmd("Page.setWindowOpenHandler", {"handler": "deny"})
        except Exception:
            pass
        QtCore.QTimer.singleShot(600, self._finalize_browser_start)

    def _finalize_browser_start(self):
        if not self.driver:
            return
        self._select_main_window_handle()
        self.ensure_chrome_hwnd(force=True)
        self.apply_taskbar_merge()
        self.sync_attach_positions(force=True)

        self.apply_browser_window_size(resize_now=True)
        self.apply_zoom()
        self.apply_alpha()
        self.apply_browser_topmost()
        self.apply_browser_title()
        self.apply_browser_icon()
        self.refresh_status(force_icon=True)

    def _kill_proc_obj(self, proc):
        if not proc:
            return
        try:
            kill_pid_tree(getattr(proc, "pid", 0))
        except Exception:
            pass
        try:
            proc.terminate()
            proc.wait(timeout=2)
        except Exception:
            try:
                proc.kill()
            except Exception:
                pass

    def _kill_proc_and_profile(self, proc, profile_dir: str):
        self._kill_proc_obj(proc)
        try:
            kill_browsers_by_profile(profile_dir)
        except Exception:
            pass

    def _quit_driver_obj(self, driver):
        try:
            if driver:
                driver.quit()
        except Exception:
            pass

    def safe_kill_proc(self):
        self._kill_proc_and_profile(self.proc, self.profile_dir)
        self.proc = None

    def safe_quit_driver(self):
        self._quit_driver_obj(self.driver)
        self.driver = None
        self.main_window_handle = None

    def restart_browser(self):
        cur = self.get_current_url()
        if cur:
            self.url_var.set(cur)

        self.safe_quit_driver()
        self.safe_kill_proc()
        self.chrome_hwnd = None
        self._stop_attach_timer()
        self.start_browser_async()

    def _restore_browser_window(self, silent: bool = False) -> bool:
        if not self.proc and not self.driver:
            if not silent:
                self._show_warning("浏览器未启动")
            return False
        try:
            hwnd = self.ensure_chrome_hwnd(force=True)
            if not hwnd:
                hwnds = self.get_browser_hwnds(include_all=False)
                if hwnds:
                    hwnd = hwnds[0]
            if not hwnd:
                if not silent:
                    self._show_warning("找不到浏览器窗口")
                return False
            self.chrome_hwnd = hwnd
            restore_window(hwnd)
            self.resize_browser_window(self.win_w, self.win_h)
            self.apply_taskbar_merge()
            self.sync_attach_positions(force=True)
            if self.browser_top_var.get():
                self.apply_browser_topmost(force=True, save=False)
            else:
                self.raise_browser_above_panel()
            self.arrange_zorder()
            return True
        except Exception:
            return False

    def recover_browser(self, silent: bool = False):
        self._restore_browser_window(silent=silent)

    def open_additional_window(self, url: str):
        prev = None
        if self.proc or self.driver:
            prev = {
                "proc": self.proc,
                "driver": self.driver,
                "port": self.port,
                "hwnd": self.chrome_hwnd,
                "profile_dir": self.profile_dir,
            }
        profile_dir = make_extra_profile_dir()
        old_state = (self.proc, self.driver, self.port, self.chrome_hwnd, self.profile_dir)
        self.proc = None
        self.driver = None
        self.port = None
        self.chrome_hwnd = None
        self.profile_dir = profile_dir
        try:
            self._start_browser_blocking(profile_dir=profile_dir, url=url)
        except Exception as e:
            self.proc, self.driver, self.port, self.chrome_hwnd, self.profile_dir = old_state
            self._show_error(str(e))
            return False
        if prev:
            self.extra_sessions.append(prev)
        return True

    def start_extra_window_async(self, url: str):
        if self._multi_start_thread:
            try:
                self.status_var.set("正在打开新窗口...")
            except Exception:
                pass
            return
        url = normalize_url(url)
        self.url_var.set(url)
        self.remember_url(url)

        prev = None
        if self.proc or self.driver:
            prev = {
                "proc": self.proc,
                "driver": self.driver,
                "port": self.port,
                "hwnd": self.chrome_hwnd,
                "profile_dir": self.profile_dir,
            }

        profile_dir = make_extra_profile_dir()
        x, y = self.calc_browser_position(self.win_w, self.win_h)
        try:
            self.status_var.set("正在打开新窗口...")
        except Exception:
            pass

        worker = BrowserStartWorker(
            self.chrome_exe, url, self.win_w, self.win_h, x, y, profile_dir, self.audio_enabled
        )
        thread = QtCore.QThread(self)
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.finished.connect(
            lambda proc, driver, port, err: self._on_extra_window_started(
                proc, driver, port, err, prev, profile_dir
            )
        )
        worker.finished.connect(thread.quit)
        worker.finished.connect(worker.deleteLater)
        thread.finished.connect(thread.deleteLater)
        self._multi_start_thread = thread
        self._multi_start_worker = worker
        thread.start()

    def _on_extra_window_started(self, proc, driver, port, err: str, prev, profile_dir: str):
        self._multi_start_thread = None
        self._multi_start_worker = None
        if err or not driver:
            msg = err or "浏览器未就绪"
            try:
                self.status_var.set(f"启动失败: {msg}")
            except Exception:
                pass
            try:
                self._kill_proc_and_profile(proc, profile_dir)
            except Exception:
                pass
            self._show_error(msg)
            return

        if prev:
            self.extra_sessions.append(prev)
        self.proc = proc
        self.driver = driver
        self.port = port
        self.profile_dir = profile_dir
        self.chrome_hwnd = None
        self.main_window_handle = None

        try:
            self.driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {"source": PREVENT_NEW_WINDOWS_JS})
        except Exception:
            pass
        try:
            self.driver.execute_cdp_cmd("Page.setWindowOpenHandler", {"handler": "deny"})
        except Exception:
            pass
        QtCore.QTimer.singleShot(600, self._finalize_browser_start)

    def go(self):
        url = normalize_url(self.url_var.get())
        self.url_var.set(url)
        self.remember_url(url)

        if self.multi_open_button.isChecked() and self.driver:
            self.open_additional_window(url)
            return
        if not self.driver:
            self.start_browser_safe()
            return
        try:
            self._ensure_driver_window(expected_url=url)
            self.driver.get(url)
            time.sleep(0.25)
            self.apply_zoom()
            self.apply_browser_title()
            self.apply_browser_icon()
            self.refresh_status(force_icon=True)
            self._apply_audio_state(silent=True)
        except Exception:
            self.restart_browser()

    def _step_slider(self, slider: QtWidgets.QSlider, delta: int):
        try:
            value = slider.value() + delta
            value = max(slider.minimum(), min(slider.maximum(), value))
            slider.setValue(value)
        except Exception:
            pass

    def browser_refresh(self):
        if not self.driver:
            return
        try:
            self._ensure_driver_window()
            self.driver.refresh()
            time.sleep(0.1)
            self.apply_zoom()
            self.apply_browser_title()
            self.apply_browser_icon()
            self.refresh_status(force_icon=True)
            self._apply_audio_state(silent=True)
        except Exception:
            self.restart_browser()

    def browser_back(self):
        if not self.driver:
            return
        try:
            self._ensure_driver_window()
            self.driver.back()
            time.sleep(0.1)
            self.apply_zoom()
            self.refresh_status(force_icon=True)
            self._apply_audio_state(silent=True)
        except Exception:
            pass

    def page_zoom_in(self):
        self._step_slider(self.zoom_slider, 5)

    def page_zoom_out(self):
        self._step_slider(self.zoom_slider, -5)

    def window_scale_up(self):
        self._step_slider(self.window_scale_slider, 5)

    def window_scale_down(self):
        self._step_slider(self.window_scale_slider, -5)

    def toggle_mute_shortcut(self):
        try:
            self.audio_checkbox.setChecked(not self.audio_checkbox.isChecked())
        except Exception:
            pass

    # ---------- memory ----------
    def remember_url(self, url: str):
        url = normalize_url(url)
        self.settings["last_url"] = url
        rec = self.settings.get("recent", [])
        if url in rec:
            rec.remove(url)
        rec.insert(0, url)
        self.settings["recent"] = rec[:30]
        save_settings(self.settings)

    def get_current_url(self):
        try:
            return self.driver.current_url if self.driver else ""
        except Exception:
            return ""

    def enforce_single_window(self):
        if not self.driver:
            return
        try:
            handles = list(self.driver.window_handles)
        except Exception:
            return
        if not handles:
            return
        if len(handles) > 1:
            self._select_main_window_handle()
        elif not self.main_window_handle or self.main_window_handle not in handles:
            self.main_window_handle = handles[0]
        if len(handles) <= 1:
            return
        target_url = ""
        for h in handles:
            if h == self.main_window_handle:
                continue
            try:
                self.driver.switch_to.window(h)
                try:
                    url = self.driver.current_url or ""
                except Exception:
                    url = ""
                if url and url not in ("about:blank", "chrome://newtab/"):
                    target_url = url
                self.driver.close()
            except Exception:
                pass
        try:
            self.driver.switch_to.window(self.main_window_handle)
        except Exception:
            pass
        if target_url:
            try:
                self.driver.get(target_url)
            except Exception:
                pass

    def _select_main_window_handle(self, expected_url: str = ""):
        if not self.driver:
            return None
        try:
            handles = list(self.driver.window_handles)
        except Exception:
            return None
        if not handles:
            return None
        if len(handles) == 1:
            self.main_window_handle = handles[0]
            return handles[0]

        target_url = (expected_url or self.url_var.get() or self.settings.get("last_url") or "").strip()
        if target_url and not (target_url.startswith("http://") or target_url.startswith("https://")):
            target_url = "https://" + target_url
        target_url_l = target_url.lower()
        target_host = ""
        try:
            target_host = urllib.parse.urlparse(target_url).netloc.lower()
        except Exception:
            target_host = ""

        best = None
        best_score = -1
        for h in handles:
            try:
                self.driver.switch_to.window(h)
                cur_url = self.driver.current_url or ""
            except Exception:
                continue
            cur_url_l = cur_url.lower()
            cur_host = ""
            try:
                cur_host = urllib.parse.urlparse(cur_url).netloc.lower()
            except Exception:
                cur_host = ""
            score = 0
            if target_host and cur_host == target_host:
                score += 2
            if target_url_l and cur_url_l.startswith(target_url_l):
                score += 2
            if cur_url and cur_url not in ("about:blank", "chrome://newtab/"):
                score += 1
            if score > best_score:
                best_score = score
                best = h

        if not best:
            best = handles[0]
        try:
            self.driver.switch_to.window(best)
        except Exception:
            pass
        self.main_window_handle = best
        return best

    def ensure_chrome_hwnd(self, force=False):
        if not self.proc and not self.driver:
            return self.chrome_hwnd

        pid = self.proc.pid if self.proc else 0

        if self.port:
            port_pid = get_pid_by_port(self.port)
            if port_pid:
                pid = port_pid

        title_hint = ""
        host_hint = ""
        try:
            title_hint = self.driver.title or ""
        except Exception:
            pass
        try:
            host_hint = urllib.parse.urlparse(self.get_current_url()).netloc or ""
        except Exception:
            pass

        desired_title = (self.settings.get("browser_title") or "").strip()

        if not force and self.chrome_hwnd:
            try:
                text_l = get_window_text(self.chrome_hwnd).lower()
                if desired_title and desired_title.lower() == text_l.strip():
                    return self.chrome_hwnd
                if (title_hint and title_hint.lower() in text_l) or (host_hint and host_hint.lower() in text_l):
                    return self.chrome_hwnd
                if IsWindowVisible(self.chrome_hwnd) and not (title_hint or host_hint or desired_title):
                    return self.chrome_hwnd
            except Exception:
                pass

        best = pick_main_hwnd(
            pid,
            title_hint=title_hint,
            host_hint=host_hint,
            size_hint=(self.win_w, self.win_h),
            include_all=bool(pid),
        )

        if not best and desired_title:
            try:
                candidates = find_chrome_hwnds_by_title(desired_title, exact=True, include_hidden=True)
                if not candidates:
                    candidates = find_chrome_hwnds_by_title(desired_title, exact=False, include_hidden=True)
                if candidates:
                    best_area = -1
                    best_hwnd = None
                    for hwnd in candidates:
                        r = RECT()
                        if not GetWindowRect(hwnd, ctypes.byref(r)):
                            continue
                        area = max(0, (r.right - r.left)) * max(0, (r.bottom - r.top))
                        if area > best_area:
                            best_area = area
                            best_hwnd = hwnd
                    best = best_hwnd
            except Exception:
                pass

        self.chrome_hwnd = best
        return self.chrome_hwnd

    def get_browser_hwnds(self, include_all: bool = True):
        hwnds = []
        pid = 0
        if self.proc or self.driver:
            pid = self.proc.pid if self.proc else 0
            if self.port:
                port_pid = get_pid_by_port(self.port)
                if port_pid:
                    pid = port_pid

        if pid:
            hwnds = get_pid_hwnds(pid)
            if not hwnds:
                try:
                    hwnds = get_pid_hwnds(get_related_pids(pid))
                except Exception:
                    hwnds = []

        if not hwnds:
            desired_title = (self.settings.get("browser_title") or "").strip()
            if desired_title:
                try:
                    hwnds = find_chrome_hwnds_by_title(desired_title, exact=True, include_hidden=True)
                    if not hwnds:
                        hwnds = find_chrome_hwnds_by_title(desired_title, exact=False, include_hidden=True)
                except Exception:
                    hwnds = []

        if not hwnds and include_all:
            hwnds = get_chrome_hwnds()
        return hwnds

    # ---------- hotkeys ----------
    def _parse_hotkey(self, s: str):
        s = (s or "").strip()
        if not s:
            return None
        parts = [p for p in re.split(r"[+\s]+", s.lower()) if p]
        mods = 0
        keys = []
        for p in parts:
            if p in ("ctrl", "control"):
                mods |= MOD_CONTROL
            elif p == "shift":
                mods |= MOD_SHIFT
            elif p in ("alt", "menu"):
                mods |= MOD_ALT
            elif p in ("win", "windows", "meta", "super"):
                mods |= MOD_WIN
            else:
                keys.append(p)
        if not keys:
            return None
        if len(keys) != 1:
            return {"error": "只支持一个主键（例如 Ctrl+Win+Alt+0 或 Ctrl+Win+Alt+.）"}
        key = keys[0]

        key_map = {
            ".": ("period", 0xBE, "."),
            "period": ("period", 0xBE, "."),
            "dot": ("period", 0xBE, "."),
            ",": ("comma", 0xBC, ","),
            "comma": ("comma", 0xBC, ","),
            "-": ("minus", 0xBD, "-"),
            "minus": ("minus", 0xBD, "-"),
            "=": ("equal", 0xBB, "="),
            "+": ("equal", 0xBB, "+"),
            "equal": ("equal", 0xBB, "="),
            "/": ("slash", 0xBF, "/"),
            "slash": ("slash", 0xBF, "/"),
            "\\": ("backslash", 0xDC, "\\"),
            "backslash": ("backslash", 0xDC, "\\"),
            "space": ("space", 0x20, "Space"),
        }
        if key in key_map:
            keysym, vk, label = key_map[key]
        elif len(key) == 1 and key.isdigit():
            keysym, vk, label = key, ord(key), key
        elif len(key) == 1 and key.isalpha():
            keysym, vk, label = key.lower(), ord(key.upper()), key.upper()
        else:
            return {"error": f"不支持的按键: {key}"}

        mods_label = []
        if mods & MOD_CONTROL:
            mods_label.append("Ctrl")
        if mods & MOD_SHIFT:
            mods_label.append("Shift")
        if mods & MOD_ALT:
            mods_label.append("Alt")
        if mods & MOD_WIN:
            mods_label.append("Win")
        display = "+".join(mods_label + [label]) if mods_label else label
        return {"mods": mods, "vk": vk, "keysym": keysym, "display": display}

    def apply_hotkeys(self, toggle_str: str, lock_str: str, close_str: str, save=True):
        toggle_info = self._parse_hotkey(toggle_str)
        lock_info = self._parse_hotkey(lock_str)
        close_info = self._parse_hotkey(close_str)
        errors = []
        if not toggle_info:
            errors.append("最小化: 快捷键格式无效")
        elif "error" in toggle_info:
            errors.append(f"最小化: {toggle_info['error']}")
        if not lock_info:
            errors.append("恢复: 快捷键格式无效")
        elif "error" in lock_info:
            errors.append(f"恢复: {lock_info['error']}")
        if not close_info:
            errors.append("关闭全部: 快捷键格式无效")
        elif "error" in close_info:
            errors.append(f"关闭全部: {close_info['error']}")
        if errors:
            self._show_warning("\n".join(errors))
            return False

        self.hotkey_toggle = toggle_info["display"]
        self.hotkey_lock = lock_info["display"]
        self.hotkey_close = close_info["display"]
        if save:
            self.settings["hotkey_toggle"] = self.hotkey_toggle
            self.settings["hotkey_lock"] = self.hotkey_lock
            self.settings["hotkey_close"] = self.hotkey_close
            save_settings(self.settings)

        self._clear_local_shortcuts()
        self._set_local_shortcut(self.hotkey_toggle, self.minimize_all)
        self._set_local_shortcut(self.hotkey_lock, self.restore_all)
        self._set_local_shortcut(self.hotkey_close, self.close_all)
        self._set_local_shortcut(HOTKEY_BROWSER_REFRESH, self.browser_refresh)
        self._set_local_shortcut(HOTKEY_BROWSER_BACK, self.browser_back)
        self._set_local_shortcut(HOTKEY_PAGE_ZOOM_IN, self.page_zoom_in)
        self._set_local_shortcut(HOTKEY_PAGE_ZOOM_OUT, self.page_zoom_out)
        self._set_local_shortcut(HOTKEY_TOGGLE_MUTE, self.toggle_mute_shortcut)
        self._set_local_shortcut(HOTKEY_WINDOW_SCALE_UP, self.window_scale_up)
        self._set_local_shortcut(HOTKEY_WINDOW_SCALE_DOWN, self.window_scale_down)

        if self._ensure_ahk_running():
            self._clear_global_hotkeys()
            self._sync_ahk_config()
            return True

        failed = self._register_global_hotkeys(toggle_info, lock_info, close_info)
        if failed:
            self._show_warning("以下快捷键注册失败，可能被系统或其他程序占用：\n" + "\n".join(failed))
        return True

    def _register_global_hotkeys(self, toggle_info, lock_info, close_info):
        hwnd = self._get_hwnd()
        for hk_id in list(self._global_hotkeys.keys()):
            try:
                UnregisterHotKey(hwnd, hk_id)
            except Exception:
                pass
        self._global_hotkeys = {}
        failed = []
        if not hwnd:
            failed.append(f"最小化: {toggle_info.get('display','')}")
            failed.append(f"恢复: {lock_info.get('display','')}")
            failed.append(f"关闭全部: {close_info.get('display','')}")
            return failed
        try:
            if RegisterHotKey(hwnd, 1, toggle_info["mods"] | MOD_NOREPEAT, toggle_info["vk"]) or \
               RegisterHotKey(hwnd, 1, toggle_info["mods"], toggle_info["vk"]):
                self._global_hotkeys[1] = "minimize"
            else:
                failed.append(f"最小化: {toggle_info.get('display','')}")
            if RegisterHotKey(hwnd, 2, lock_info["mods"] | MOD_NOREPEAT, lock_info["vk"]) or \
               RegisterHotKey(hwnd, 2, lock_info["mods"], lock_info["vk"]):
                self._global_hotkeys[2] = "restore"
            else:
                failed.append(f"恢复: {lock_info.get('display','')}")
            if RegisterHotKey(hwnd, 3, close_info["mods"] | MOD_NOREPEAT, close_info["vk"]) or \
               RegisterHotKey(hwnd, 3, close_info["mods"], close_info["vk"]):
                self._global_hotkeys[3] = "close"
            else:
                failed.append(f"关闭全部: {close_info.get('display','')}")
        except Exception:
            pass
        return failed

    def _clear_global_hotkeys(self):
        hwnd = self._get_hwnd()
        for hk_id in list(self._global_hotkeys.keys()):
            try:
                UnregisterHotKey(hwnd, hk_id)
            except Exception:
                pass
        self._global_hotkeys = {}

    def _place_dialog_next_to_panel(self, win, gap=8, size=None):
        try:
            self.adjustSize()
            win.adjustSize()
            panel_geo = self.frameGeometry()
            if size:
                ww, wh = size
            else:
                ww = win.width() or win.sizeHint().width()
                wh = win.height() or win.sizeHint().height()
            screen = self.screen() or QtGui.QGuiApplication.primaryScreen()
            if not screen:
                return
            screen_geo = screen.availableGeometry()
            ww = max(1, min(int(ww), screen_geo.width()))
            wh = max(1, min(int(wh), screen_geo.height()))
            x_right = panel_geo.right() + gap
            x_left = panel_geo.left() - ww - gap
            if x_right + ww <= screen_geo.right():
                x = x_right
            else:
                x = max(screen_geo.left(), x_left)
            y = max(screen_geo.top(), min(panel_geo.top(), screen_geo.bottom() - wh))
            win.setGeometry(int(x), int(y), int(ww), int(wh))
        except Exception:
            pass

    def open_hotkey_dialog(self):
        self._show_warning("快捷键冲突，可能被系统或其他程序占用。")
        win = QtWidgets.QDialog(self)
        win.setWindowTitle("修改快捷键")
        win.setModal(False)
        layout = QtWidgets.QVBoxLayout(win)

        form = QtWidgets.QFormLayout()
        toggle_edit = QtWidgets.QLineEdit(self.hotkey_toggle)
        lock_edit = QtWidgets.QLineEdit(self.hotkey_lock)
        close_edit = QtWidgets.QLineEdit(self.hotkey_close)
        form.addRow("最小化", toggle_edit)
        form.addRow("恢复", lock_edit)
        form.addRow("关闭全部", close_edit)
        top_on_edit = QtWidgets.QLineEdit("Ctrl+Win+Alt+T")
        top_on_edit.setReadOnly(True)
        top_off_edit = QtWidgets.QLineEdit("Ctrl+Shift+Win+T")
        top_off_edit.setReadOnly(True)
        form.addRow("置顶快捷键", top_on_edit)
        form.addRow("取消置顶", top_off_edit)
        refresh_edit = QtWidgets.QLineEdit(HOTKEY_BROWSER_REFRESH)
        refresh_edit.setReadOnly(True)
        back_edit = QtWidgets.QLineEdit(HOTKEY_BROWSER_BACK)
        back_edit.setReadOnly(True)
        zoom_in_edit = QtWidgets.QLineEdit(HOTKEY_PAGE_ZOOM_IN)
        zoom_in_edit.setReadOnly(True)
        zoom_out_edit = QtWidgets.QLineEdit(HOTKEY_PAGE_ZOOM_OUT)
        zoom_out_edit.setReadOnly(True)
        mute_edit = QtWidgets.QLineEdit(HOTKEY_TOGGLE_MUTE)
        mute_edit.setReadOnly(True)
        win_scale_up_edit = QtWidgets.QLineEdit(HOTKEY_WINDOW_SCALE_UP)
        win_scale_up_edit.setReadOnly(True)
        win_scale_down_edit = QtWidgets.QLineEdit(HOTKEY_WINDOW_SCALE_DOWN)
        win_scale_down_edit.setReadOnly(True)
        form.addRow("浏览器刷新", refresh_edit)
        form.addRow("浏览器返回", back_edit)
        form.addRow("页面放大", zoom_in_edit)
        form.addRow("页面缩小", zoom_out_edit)
        form.addRow("静音开关", mute_edit)
        form.addRow("窗口放大", win_scale_up_edit)
        form.addRow("窗口缩小", win_scale_down_edit)
        layout.addLayout(form)

        tip = QtWidgets.QLabel("示例: Ctrl+Win+Alt+0（只支持一个主键）")
        layout.addWidget(tip)

        btns = QtWidgets.QHBoxLayout()
        btns.addStretch(1)
        btn_apply = QtWidgets.QPushButton("应用")
        btn_cancel = QtWidgets.QPushButton("取消")
        btns.addWidget(btn_apply)
        btns.addWidget(btn_cancel)
        layout.addLayout(btns)

        def apply_and_close():
            ok = self.apply_hotkeys(toggle_edit.text(), lock_edit.text(), close_edit.text(), save=True)
            if ok:
                win.close()

        btn_apply.clicked.connect(apply_and_close)
        btn_cancel.clicked.connect(win.close)
        win.adjustSize()
        self._place_dialog_next_to_panel(win)
        win.show()

    def on_browser_top_toggle(self, checked):
        if checked:
            self._show_topmost_invalid_hint()
        self.apply_browser_topmost(force=True)

    # ---------- hide/lock ----------
    def minimize_all(self):
        try:
            self.showMinimized()
        except Exception:
            pass
        try:
            self.ensure_chrome_hwnd()
            minimize_window(self.chrome_hwnd)
        except Exception:
            pass

    def restore_all(self):
        try:
            self.showNormal()
            self.raise_()
        except Exception:
            pass
        try:
            self.ensure_chrome_hwnd()
            restore_window(self.chrome_hwnd)
        except Exception:
            pass

    def close_all(self):
        try:
            self.safe_quit_driver()
            self.safe_kill_proc()
            for sess in self.extra_sessions:
                self._quit_driver_obj(sess.get("driver"))
                self._kill_proc_and_profile(sess.get("proc"), sess.get("profile_dir"))
            self.extra_sessions = []
        except Exception:
            pass
        try:
            self.request_exit()
        except Exception:
            pass

    def request_exit(self):
        self._force_close = True
        try:
            self.close()
        finally:
            self._force_close = False

    def hide_all(self):
        try:
            self.hide()
        except Exception:
            pass
        try:
            self.ensure_chrome_hwnd()
            hide_window(self.chrome_hwnd)
        except Exception:
            pass

    def toggle_hide(self):
        if self.lock_hidden:
            return
        if not self.hidden_toggle:
            self.hidden_toggle = True
            self.minimize_all()
        else:
            self.hidden_toggle = False
            self.restore_all()

    def toggle_lock(self):
        if self.lock_hidden:
            self.lock_hidden = False
            self.hidden_toggle = False
            self.restore_all()
            return
        self.lock_hidden = True
        self.hidden_toggle = True
        self.hide_all()

    def center_panel(self):
        try:
            self.adjustSize()
            geo = self.frameGeometry()
            screen = self.screen() or QtGui.QGuiApplication.primaryScreen()
            if not screen:
                return
            center = screen.availableGeometry().center()
            geo.moveCenter(center)
            self.move(geo.topLeft())
        except Exception:
            pass

    # ---------- window size ----------
    def get_ratio_key(self):
        label = (self.browser_ratio_var.get() or "").strip()
        return RATIO_LABEL_TO_KEY.get(label, "4:3")

    def get_base_window_size(self):
        ratio = self.get_ratio_key()
        level = self.browser_size_level_var.get() or "S"
        preset = WINDOW_SIZE_PRESETS.get(ratio, WINDOW_SIZE_PRESETS["4:3"])
        return preset.get(level, preset["S"])

    def get_browser_position_key(self):
        label = (self.browser_pos_var.get() or "").strip()
        return BROWSER_POS_LABEL_TO_KEY.get(label, "bottom_right")

    def calc_browser_position(self, w: int, h: int):
        try:
            screen = self.screen() or QtGui.QGuiApplication.primaryScreen()
            if not screen:
                return 0, 0
            geo = screen.availableGeometry()
            left = int(geo.left())
            top = int(geo.top())
            sw = int(geo.width())
            sh = int(geo.height())
        except Exception:
            return 0, 0

        key = self.get_browser_position_key()
        m = BROWSER_POS_MARGIN

        if key == "bottom_left":
            x = left + m
            y = top + sh - h - m
        elif key == "top_right":
            x = left + sw - w - m
            y = top + m
        elif key == "top_left":
            x = left + m
            y = top + m
        elif key == "center":
            x = left + int((sw - w) / 2)
            y = top + int((sh - h) / 2)
        elif key == "right_center":
            x = left + sw - w - m
            y = top + int((sh - h) / 2)
        elif key == "left_center":
            x = left + m
            y = top + int((sh - h) / 2)
        elif key == "top_center":
            x = left + int((sw - w) / 2)
            y = top + m
        elif key == "bottom_center":
            x = left + int((sw - w) / 2)
            y = top + sh - h - m
        else:
            x = left + sw - w - m
            y = top + sh - h - m

        max_x = left + max(0, sw - w)
        max_y = top + max(0, sh - h)
        x = max(left, min(int(x), max_x))
        y = max(top, min(int(y), max_y))
        return x, y

    def on_browser_position_change(self):
        self.settings["browser_position"] = self.get_browser_position_key()
        save_settings(self.settings)
        self.resize_browser_window(self.win_w, self.win_h)

    def raise_browser_above_panel(self):
        try:
            self.ensure_chrome_hwnd()
            hwnds = self.get_browser_hwnds(include_all=False)
            if not hwnds and self.chrome_hwnd:
                hwnds = [self.chrome_hwnd]
            if not hwnds:
                return
            panel_top = bool(self.panel_top_var.get())
            browser_top = bool(self.browser_top_var.get())
            if panel_top and browser_top:
                return
            if browser_top:
                for hwnd in hwnds:
                    set_window_topmost(hwnd, True, force=True)
                return
            if panel_top:
                return
            flags = SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE
            for hwnd in hwnds:
                SetWindowPos(hwnd, HWND_TOP, 0, 0, 0, 0, flags)
        except Exception:
            pass

    def get_attach_side(self):
        label = ""
        try:
            label = self.attach_side_combo.currentText()
        except Exception:
            label = "左侧" if self.attach_side == "left" else "右侧"
        return "left" if "左" in label else "right"

    def on_attach_toggle(self, *_):
        self.settings["attach_enabled"] = bool(self.attach_checkbox.isChecked())
        save_settings(self.settings)
        if self.attach_checkbox.isChecked():
            self.sync_attach_positions(force=True)

    def on_attach_side_change(self, label: str):
        side = "left" if "左" in (label or "") else "right"
        self.settings["attach_side"] = side
        self.attach_side = side
        save_settings(self.settings)
        self.sync_attach_positions(force=True)

    def apply_taskbar_merge(self, *_):
        self.merge_taskbar = bool(self.merge_checkbox.isChecked())
        self.settings["merge_taskbar"] = self.merge_taskbar
        save_settings(self.settings)
        try:
            self.ensure_chrome_hwnd()
            if self.chrome_hwnd:
                owner = self._get_hwnd() if self.merge_taskbar else 0
                set_window_owner(self.chrome_hwnd, owner)
        except Exception:
            pass

    def _panel_geo(self):
        geo = self.frameGeometry()
        return int(geo.left()), int(geo.top()), int(geo.width()), int(geo.height())

    def _get_screen_for_point(self, x: int, y: int):
        try:
            screen = QtGui.QGuiApplication.screenAt(QtCore.QPoint(int(x), int(y)))
        except Exception:
            screen = None
        if not screen:
            screen = self.screen() or QtGui.QGuiApplication.primaryScreen()
        return screen

    def _clamp_panel_pos(self, x: int, y: int, panel_w: int, panel_h: int):
        screen = self._get_screen_for_point(x + int(panel_w / 2), y + int(panel_h / 2))
        if not screen:
            return int(x), int(y)
        geo = screen.availableGeometry()
        left = int(geo.left())
        top = int(geo.top())
        max_x = left + max(0, int(geo.width()) - int(panel_w))
        max_y = top + max(0, int(geo.height()) - int(panel_h))
        x = max(left, min(int(x), max_x))
        y = max(top, min(int(y), max_y))
        return x, y

    def _calc_attach_side_from_positions(self, panel_left: int, panel_w: int, browser_left: int, browser_w: int):
        if panel_left + panel_w <= browser_left:
            return "left"
        if panel_left >= browser_left + browser_w:
            return "right"
        return self.get_attach_side()

    def _calc_panel_pos_for_browser(self, browser_left: int, browser_top: int, browser_w: int, browser_h: int,
                                    panel_w: int, panel_h: int, side_pref: str):
        screen = self._get_screen_for_point(browser_left + int(browser_w / 2), browser_top + int(browser_h / 2))
        if not screen:
            return browser_left, browser_top, side_pref
        geo = screen.availableGeometry()
        left = int(geo.left())
        top = int(geo.top())
        right = left + int(geo.width())
        bottom = top + int(geo.height())
        m = ATTACH_MARGIN

        x_left = browser_left - panel_w - m
        x_right = browser_left + browser_w + m

        def fits(xv):
            return xv >= left and (xv + panel_w) <= right

        side_used = side_pref
        x = x_left if side_pref == "left" else x_right
        if not fits(x):
            alt_side = "right" if side_pref == "left" else "left"
            alt_x = x_right if alt_side == "right" else x_left
            if fits(alt_x):
                side_used = alt_side
                x = alt_x
            else:
                x = max(left, min(int(x), right - panel_w))

        y = max(top, min(int(browser_top), bottom - panel_h))
        return int(x), int(y), side_used

    def _move_browser_window(self, x: int, y: int):
        if not self.chrome_hwnd:
            return
        flags = SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE
        try:
            SetWindowPos(self.chrome_hwnd, 0, int(x), int(y), 0, 0, flags)
        except Exception:
            pass

    def sync_attach_positions(self, force: bool = False):
        if not self.attach_checkbox.isChecked():
            return
        if self._syncing_attach:
            return
        try:
            self.ensure_chrome_hwnd()
        except Exception:
            return
        if not self.chrome_hwnd:
            return
        rect = get_window_rect(self.chrome_hwnd)
        if not rect:
            return
        panel_left, panel_top, panel_w, panel_h = self._panel_geo()
        browser_left, browser_top, browser_right, browser_bottom = rect
        browser_w = max(0, browser_right - browser_left)
        browser_h = max(0, browser_bottom - browser_top)

        if self._last_panel_pos is None:
            self._last_panel_pos = (panel_left, panel_top)
        if self._last_browser_rect is None:
            self._last_browser_rect = (browser_left, browser_top, browser_w, browser_h)

        panel_moved = (panel_left, panel_top) != self._last_panel_pos
        browser_moved = (browser_left, browser_top, browser_w, browser_h) != self._last_browser_rect

        if force:
            panel_moved = False
            browser_moved = True

        side = self.get_attach_side()
        m = ATTACH_MARGIN

        if browser_moved and not panel_moved:
            target_x, target_y, _ = self._calc_panel_pos_for_browser(
                browser_left, browser_top, browser_w, browser_h, panel_w, panel_h, side
            )
            self._syncing_attach = True
            try:
                self.move(int(target_x), int(target_y))
            finally:
                self._syncing_attach = False
            panel_left, panel_top, panel_w, panel_h = self._panel_geo()
        elif panel_moved and not browser_moved:
            side = self._calc_attach_side_from_positions(panel_left, panel_w, browser_left, browser_w)
            if side == "left":
                target_x = panel_left + panel_w + m
            else:
                target_x = panel_left - browser_w - m
            target_y = panel_top
            self._move_browser_window(int(target_x), int(target_y))
            browser_left, browser_top = target_x, target_y

        self._last_panel_pos = (panel_left, panel_top)
        self._last_browser_rect = (browser_left, browser_top, browser_w, browser_h)

    def arrange_zorder(self):
        try:
            self.ensure_chrome_hwnd()
            if not self.chrome_hwnd:
                return
            panel_hwnd = self._get_hwnd()
            if not panel_hwnd:
                return
            flags = SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE
            panel_top = bool(self.panel_top_var.get())
            browser_top = bool(self.browser_top_var.get())
            if panel_top and browser_top:
                return
            if panel_top:
                SetWindowPos(panel_hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
                if browser_top:
                    SetWindowPos(self.chrome_hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
                else:
                    SetWindowPos(self.chrome_hwnd, HWND_TOP, 0, 0, 0, 0, flags)
                SetWindowPos(panel_hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
            else:
                if browser_top:
                    SetWindowPos(self.chrome_hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
                    SetWindowPos(panel_hwnd, HWND_TOP, 0, 0, 0, 0, flags)
                else:
                    SetWindowPos(self.chrome_hwnd, HWND_TOP, 0, 0, 0, 0, flags)
                    SetWindowPos(panel_hwnd, HWND_TOP, 0, 0, 0, 0, flags)
        except Exception:
            pass

    def force_restack(self):
        self.ensure_chrome_hwnd(force=True)
        self.apply_panel_topmost()
        self.apply_browser_topmost(force=True, save=False)
        self.apply_taskbar_merge()
        self.arrange_zorder()
        self.sync_attach_positions(force=True)

    def apply_browser_window_size(self, resize_now=True):
        base_w, base_h = self.get_base_window_size()
        scale = float(self.window_scale_var.get())
        scale = max(BROWSER_SCALE_MIN, min(BROWSER_SCALE_MAX, scale))
        if abs(scale - float(self.window_scale_var.get())) > 1e-6:
            self.window_scale_var.set(scale)
        w = max(100, int(base_w * scale))
        h = max(100, int(base_h * scale))
        self.win_w, self.win_h = w, h
        self.settings["browser_ratio"] = self.get_ratio_key()
        self.settings["browser_size_level"] = self.browser_size_level_var.get() or "M"
        self.settings["browser_scale"] = scale
        save_settings(self.settings)
        if self.window_scale_label:
            self.window_scale_label.setText(f"{scale:.2f}x")
        if self.window_size_label:
            level = self.settings["browser_size_level"]
            level_label = SIZE_LEVEL_LABEL.get(level, level)
            self.window_size_label.setText(f"{level_label} {w}x{h}")
        if resize_now:
            self.resize_browser_window(w, h)

    def resize_browser_window(self, w: int, h: int):
        if w <= 0 or h <= 0:
            return
        x, y = self.calc_browser_position(w, h)
        try:
            self.ensure_chrome_hwnd()
            if not self.chrome_hwnd:
                raise RuntimeError("no hwnd")
            flags = SWP_NOZORDER | SWP_NOACTIVATE
            SetWindowPos(self.chrome_hwnd, 0, x, y, w, h, flags)
            self.raise_browser_above_panel()
            return
        except Exception:
            pass
        if self.driver:
            try:
                self.driver.set_window_rect(x=x, y=y, width=w, height=h)
                self.raise_browser_above_panel()
            except Exception:
                pass

    def position_browser_bottom_right(self):
        self.resize_browser_window(self.win_w, self.win_h)

    def on_browser_ratio_change(self):
        self.apply_browser_window_size(resize_now=True)

    def on_browser_size_level_change(self):
        self.apply_browser_window_size(resize_now=True)

    def on_window_scale(self, _):
        self.apply_browser_window_size(resize_now=True)

    # ---------- real-time zoom ----------
    def apply_zoom(self):
        z = float(self.zoom_var.get())
        self.zoom_label.setText(f"{z:.2f}x")
        self.settings["remember_zoom"] = z
        save_settings(self.settings)
        if not self.driver:
            return
        self._ensure_driver_window()
        ok = False
        try:
            self.driver.execute_cdp_cmd("Page.setZoomFactor", {"zoomFactor": z})
            ok = True
        except Exception:
            pass
        if not ok:
            try:
                self.driver.execute_script(
                    "document.documentElement.style.zoom = arguments[0];"
                    "document.body && (document.body.style.zoom = arguments[0]);",
                    z,
                )
            except Exception:
                pass

    def on_zoom(self, _):
        self.apply_zoom()

    # ---------- transparency ----------
    def apply_alpha(self):
        a = float(self.alpha_var.get())
        self.alpha_label.setText(f"{a:.2f}")
        self.settings["remember_alpha"] = a
        save_settings(self.settings)
        try:
            self.ensure_chrome_hwnd()
            if self.chrome_hwnd:
                set_window_alpha(self.chrome_hwnd, a)
        except Exception:
            pass

    def on_alpha(self, _):
        self.apply_alpha()

    def on_panel_alpha(self, _):
        a = float(self.panel_alpha_var.get())
        self.panel_alpha_label.setText(f"{a:.2f}")
        self.settings["remember_panel_alpha"] = a
        save_settings(self.settings)
        try:
            self.setWindowOpacity(a)
        except Exception:
            pass

    # ---------- audio ----------
    def _audio_toggle_js(self, muted: bool) -> str:
        flag = "true" if muted else "false"
        return f"""
        (function(){{
          const muted = {flag};
          try {{ window.__miniFishMuted = muted; }} catch(e) {{}}
          function apply() {{
            try {{
              const els = document.querySelectorAll('video, audio');
              for (const el of els) {{
                try {{ el.muted = muted; }} catch(e) {{}}
                if (!muted) {{
                  try {{ if (el.volume === 0) el.volume = 1; }} catch(e) {{}}
                }}
              }}
            }} catch(e) {{}}
          }}
          apply();
          try {{
            if (!window.__miniFishAudioObserver) {{
              window.__miniFishAudioObserver = new MutationObserver(() => {{
                if (window.__miniFishMuted) {{
                  apply();
                }}
              }});
              window.__miniFishAudioObserver.observe(document.documentElement || document.body, {{childList:true, subtree:true}});
            }}
          }} catch(e) {{}}
        }})();
        """

    def _apply_audio_state(self, silent: bool = False) -> bool:
        if not self.driver:
            return False
        self._ensure_driver_window()
        muted = not self.audio_enabled
        applied = False
        try:
            self.driver.execute_cdp_cmd("Page.setAudioMuted", {"muted": muted})
            applied = True
        except Exception:
            pass
        try:
            self.driver.execute_script(self._audio_toggle_js(muted))
            applied = True
        except Exception:
            if not silent:
                self._show_warning("音频切换失败，请刷新页面")
        return applied

    def on_audio_toggle(self, checked: bool):
        self.audio_enabled = bool(checked)
        self.settings["audio_enabled"] = self.audio_enabled
        save_settings(self.settings)
        if self.proc or self.driver:
            self._apply_audio_state(silent=True)
        try:
            self.status_var.set("声音已启用" if self.audio_enabled else "已静音")
        except Exception:
            pass

    # ---------- topmost ----------
    def apply_panel_topmost(self):
        v = bool(self.panel_top_var.get())
        self.settings["panel_topmost"] = v
        save_settings(self.settings)
        try:
            was_minimized = self.isMinimized()
            was_maximized = self.isMaximized()
            self.setWindowFlag(QtCore.Qt.WindowStaysOnTopHint, v)
            if was_minimized:
                self.showMinimized()
            elif was_maximized:
                self.showMaximized()
            else:
                self.show()
                self.raise_()
                self.activateWindow()
        except Exception:
            pass
        try:
            hwnd = self._get_hwnd()
            if hwnd:
                set_window_topmost(hwnd, v, force=True)
        except Exception:
            pass
        if not v:
            if not self.browser_top_var.get():
                try:
                    self.ensure_chrome_hwnd()
                    hwnds = self.get_browser_hwnds(include_all=False)
                    if not hwnds and self.chrome_hwnd:
                        hwnds = [self.chrome_hwnd]
                    for hwnd in hwnds:
                        set_window_topmost(hwnd, False, force=True)
                except Exception:
                    pass
        self.arrange_zorder()

    def apply_browser_topmost(self, force=False, save=True):
        v = bool(self.browser_top_var.get())
        if save:
            self.settings["browser_topmost"] = v
            save_settings(self.settings)
        force_flag = force or v
        try:
            self._send_ahk_cmd("top_on" if v else "top_off")
        except Exception:
            pass
        try:
            self.ensure_chrome_hwnd(force=force_flag)
            hwnds = self.get_browser_hwnds(include_all=False)
            if not hwnds and self.chrome_hwnd:
                hwnds = [self.chrome_hwnd]
            for hwnd in hwnds:
                set_window_topmost(hwnd, v, force=force_flag)
        except Exception:
            pass
        self.arrange_zorder()

    # ---------- status + favicon ----------
    def clear_custom_status(self):
        self.custom_status_var.set("")

    def on_custom_status_change(self, *_):
        text = (self.custom_status_var.get() or "").strip()
        self.settings["custom_status"] = text
        save_settings(self.settings)
        self.refresh_status(force_icon=False)

    def manual_refresh(self):
        self.on_custom_status_change()
        self.apply_panel_topmost()
        self.ensure_chrome_hwnd(force=True)
        self.apply_browser_window_size(resize_now=True)
        self.apply_zoom()
        self.apply_alpha()
        self.apply_browser_topmost(force=True)
        self.apply_browser_title()
        self.apply_browser_icon()
        self.refresh_status(force_icon=True)
        self.apply_taskbar_merge()

    def refresh_status(self, force_icon=False):
        custom = (self.custom_status_var.get() or "").strip()
        if custom:
            self.status_var.set(custom[:80])
        else:
            if not self.driver:
                self.status_var.set("disconnected")
                return
            try:
                title = self.driver.title or ""
            except Exception:
                title = ""
            cur = self.get_current_url()
            host = ""
            try:
                host = urllib.parse.urlparse(cur).netloc
            except Exception:
                pass
            show = (host + "  " + title).strip() or "ready"
            self.status_var.set(show[:80])

        if force_icon and self.driver:
            QtCore.QTimer.singleShot(100, self.update_favicon)

    def update_favicon(self):
        if not self.driver:
            return
        self._ensure_driver_window()
        cur = self.get_current_url()
        try:
            host = urllib.parse.urlparse(cur).netloc or "site"
        except Exception:
            host = "site"

        icon_url = get_best_icon_url(self.driver) or fallback_favicon_url(cur)
        if not icon_url:
            self.icon_label.setPixmap(QtGui.QPixmap())
            self.icon_label.setText("◎")
            self.site_icon_pixmap = None
            self._update_tray_icons()
            return

        # cache as png path; if it isn't png, Tk may fail and we fallback
        fn = safe_filename(host) + ".png"
        path = os.path.join(CACHE_DIR, fn)
        try:
            download_to(path, icon_url)
            pix = QtGui.QPixmap(path)
            if not pix.isNull():
                pix = pix.scaled(16, 16, QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation)
                self.site_icon_pixmap = pix
                self.icon_label.setPixmap(pix)
                self.icon_label.setText("")
                self._update_tray_icons()
                return
        except Exception:
            pass
        self.icon_label.setPixmap(QtGui.QPixmap())
        self.icon_label.setText("◎")
        self.site_icon_pixmap = None
        self._update_tray_icons()

    # ---------- safe minimize/restore ----------
    def minimize_both(self):
        try:
            self.showMinimized()
        except Exception:
            pass
        try:
            self.ensure_chrome_hwnd()
            minimize_window(self.chrome_hwnd)
        except Exception:
            pass

    def restore_both(self):
        try:
            self.showNormal()
        except Exception:
            pass
        try:
            self.ensure_chrome_hwnd()
            restore_window(self.chrome_hwnd)
        except Exception:
            pass

    # ---------- poll: auto remember url + keep topmost alive ----------
    def poll_state(self):
        self._check_ahk_events()
        if self.lock_hidden:
            self.hide_all()
            return
        if (self.proc or self.driver or self.port) and not self._starting_browser and not self._attach_timer:
            if self._is_browser_alive():
                self._browser_gone_ticks = 0
            else:
                self._browser_gone_ticks += 1
                if self._browser_gone_ticks >= 3:
                    self._stop_attach_timer()
                    self.safe_quit_driver()
                    self.safe_kill_proc()
                    self.port = None
                    self.chrome_hwnd = None
                    self.main_window_handle = None
                    try:
                        self.status_var.set("浏览器已关闭，点击Go重新启动")
                    except Exception:
                        pass
                    return
        else:
            self._browser_gone_ticks = 0
        try:
            self.enforce_single_window()
        except Exception:
            pass
        # 1) auto remember when user clicks inside page
        cur = self.get_current_url()
        if cur:
            if cur != self.settings.get("last_url"):
                self.remember_url(cur)
                self.url_var.set(cur)
                self.refresh_status(force_icon=False)
            if not self.audio_enabled and cur != self._last_audio_url:
                self._apply_audio_state(silent=True)
            self._last_audio_url = cur
        else:
            self._last_audio_url = ""

        # 2) ensure browser hwnd stays correct (sometimes changes after navigation)
        self.ensure_chrome_hwnd()
        self.apply_taskbar_merge()
        self.sync_attach_positions()

        # 3) re-apply topmost if enabled (some windows lose it)
        if self.browser_top_var.get():
            try:
                if not self.panel_top_var.get():
                    self.apply_browser_topmost(force=True, save=False)
            except Exception:
                pass
        elif self.panel_top_var.get():
            try:
                self.raise_browser_above_panel()
            except Exception:
                pass
        self.arrange_zorder()

        # 4) keep browser title alive (sites may override)
        self.apply_browser_title()
        if self.browser_icon_data_url:
            self.apply_browser_icon()

        return

    def on_close(self):
        hwnd = self._get_hwnd()
        try:
            UnregisterHotKey(hwnd, 1)
            UnregisterHotKey(hwnd, 2)
            UnregisterHotKey(hwnd, 3)
        except Exception:
            pass
        try:
            self._stop_ahk()
        except Exception:
            pass
        try:
            self._stop_attach_timer()
        except Exception:
            pass
        self.safe_quit_driver()
        self.safe_kill_proc()
        for sess in self.extra_sessions:
            self._quit_driver_obj(sess.get("driver"))
            self._kill_proc_and_profile(sess.get("proc"), sess.get("profile_dir"))
        self.extra_sessions = []
        self._destroy_tray_icons()
        try:
            QtCore.QTimer.singleShot(0, QtWidgets.QApplication.quit)
        except Exception:
            pass

    def run(self):
        self.show()

if __name__ == "__main__":
    try:
        if not acquire_single_instance():
            notify_already_running()
            sys.exit(0)
        app = QtWidgets.QApplication(sys.argv)
        win = MiniFish()
        win.show()
        sys.exit(app.exec())
    except Exception as e:
        try:
            QtWidgets.QMessageBox.critical(None, "错误", str(e))
        except Exception:
            pass
        raise
