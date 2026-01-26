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
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

APP_NAME = "牛马神器"
APP_VERSION = "4.0.25"
UPDATE_PRODUCT_KEY = "niuma_shenqi"
GITEE_RELEASE_URL = "https://gitee.com/api/v5/repos/chen-bin98/Auto-ALL_for-Ai/releases/latest"
GITHUB_RELEASE_URL = "https://api.github.com/repos/JerryC0820/Auto-ALL_for-Ai/releases/latest"
DEFAULT_SETTINGS_URL_GITEE = "https://gitee.com/chen-bin98/Auto-ALL_for-Ai/raw/main/default_settings.json"
DEFAULT_SETTINGS_URL_GITHUB = "https://raw.githubusercontent.com/JerryC0820/Auto-ALL_for-Ai/main/default_settings.json"
AHK_INSTALLER_NAME = "AutoHotkey_2.0.19_setup.exe"
AHK_INSTALLER_URL_GITEE = "https://gitee.com/chen-bin98/Auto-ALL_for-Ai/raw/main/AutoHotkey_2.0.19_setup.exe"
AHK_INSTALLER_URL_GITHUB = "https://raw.githubusercontent.com/JerryC0820/Auto-ALL_for-Ai/main/AutoHotkey_2.0.19_setup.exe"
DOWNLOAD_TIMEOUT = 20
CHUNK_SIZE = 1024 * 512


def _fetch_json(url: str):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=10) as resp:
        data = resp.read()
    return json.loads(data.decode("utf-8", errors="ignore"))


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


def fetch_latest_release():
    for source, url in (("gitee", GITEE_RELEASE_URL), ("github", GITHUB_RELEASE_URL)):
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
    raise RuntimeError("无法获取最新安装包，请检查网络")


def _download_file(url: str, dest: str, progress_cb=None):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=DOWNLOAD_TIMEOUT) as resp:
        total = int(resp.headers.get("Content-Length") or 0)
        done = 0
        with open(dest, "wb") as f:
            while True:
                chunk = resp.read(CHUNK_SIZE)
                if not chunk:
                    break
                f.write(chunk)
                done += len(chunk)
                if progress_cb:
                    progress_cb(done, total)


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


def _copy_package(src_root: str, exe_name: str, install_dir: str):
    os.makedirs(install_dir, exist_ok=True)
    internal_src = os.path.join(src_root, "_internal")
    if os.path.isdir(internal_src):
        shutil.copytree(internal_src, os.path.join(install_dir, "_internal"), dirs_exist_ok=True)
    assets_src = os.path.join(src_root, "assets")
    if os.path.isdir(assets_src):
        shutil.copytree(assets_src, os.path.join(install_dir, "assets"), dirs_exist_ok=True)
    exe_src = os.path.join(src_root, exe_name)
    exe_dst = os.path.join(install_dir, exe_name)
    shutil.copy2(exe_src, exe_dst)
    return exe_dst


def _download_default_settings(install_dir: str):
    settings_path = os.path.join(install_dir, "_mini_fish_settings.json")
    for url in (DEFAULT_SETTINGS_URL_GITEE, DEFAULT_SETTINGS_URL_GITHUB):
        try:
            _download_file(url, settings_path)
            return True
        except Exception:
            continue
    return False


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


def _download_ahk_installer(dest_path: str):
    for url in (AHK_INSTALLER_URL_GITEE, AHK_INSTALLER_URL_GITHUB):
        try:
            _download_file(url, dest_path)
            return True
        except Exception:
            continue
    return False


def _ensure_ahk_installed(temp_dir: str):
    if _find_ahk_exe():
        return True
    installer_path = os.path.join(temp_dir, AHK_INSTALLER_NAME)
    if not os.path.exists(installer_path):
        ok = _download_ahk_installer(installer_path)
        if not ok:
            return False
    arg_sets = [
        ["/VERYSILENT", "/SUPPRESSMSGBOXES", "/NORESTART", "/SP-"],
        ["/SILENT", "/SUPPRESSMSGBOXES", "/NORESTART", "/SP-"],
        ["/S"],
        ["/silent"],
    ]
    for args in arg_sets:
        try:
            subprocess.run([installer_path] + args, check=False)
        except Exception:
            continue
        time.sleep(1.0)
        if _find_ahk_exe():
            return True
    return bool(_find_ahk_exe())

def _create_first_run_flag(install_dir: str):
    try:
        flag_path = os.path.join(install_dir, "_mini_fish_first_run.flag")
        with open(flag_path, "w", encoding="utf-8") as f:
            f.write("1")
    except Exception:
        pass


def _create_desktop_shortcut(exe_path: str):
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    if not os.path.isdir(desktop):
        return
    name = os.path.splitext(os.path.basename(exe_path))[0] + ".lnk"
    link_path = os.path.join(desktop, name)
    cmd = (
        "$WScriptShell = New-Object -ComObject WScript.Shell;"
        f"$Shortcut = $WScriptShell.CreateShortcut('{link_path}');"
        f"$Shortcut.TargetPath = '{exe_path}';"
        "$Shortcut.WorkingDirectory = '{os.path.dirname(exe_path)}';"
        "$Shortcut.Save()"
    )
    subprocess.run(["powershell", "-NoProfile", "-Command", cmd], check=False, creationflags=0x08000000)


class InstallerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME} 安装器 {APP_VERSION}")
        self.geometry("560x420")
        self.resizable(False, False)

        self.queue = queue.Queue()
        self.installing = False
        self.install_done = False
        self.installed_exe_path = ""

        self.agree_var = tk.BooleanVar(value=False)
        self.path_var = tk.StringVar(value=self._default_install_dir())
        self.shortcut_var = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value="等待开始安装")
        self.progress_var = tk.DoubleVar(value=0.0)

        self.container = tk.Frame(self)
        self.container.pack(fill="both", expand=True)

        self.frames = {}
        self._build_step1()
        self._build_step2()
        self._build_step3()
        self._show_step(1)

        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _default_install_dir(self):
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        base = desktop if os.path.isdir(desktop) else os.path.expanduser("~")
        return os.path.join(base, APP_NAME)

    def _build_step1(self):
        frame = tk.Frame(self.container)
        self.frames[1] = frame

        title = tk.Label(frame, text=f"欢迎使用 {APP_NAME}", font=("Segoe UI", 16, "bold"))
        title.pack(pady=12)

        info = tk.Label(
            frame,
            text=(
                "本安装器将下载最新版并安装到指定目录。\n"
                "本软件仅用于测试学习，不得用于商业用途。"
            ),
            font=("Segoe UI", 10),
            justify="center",
        )
        info.pack(pady=6)

        text = tk.Text(frame, height=12, wrap="word")
        text.insert(
            "1.0",
            "使用协议\n\n"
            "1. 本软件仅用于测试学习用途，不得用于商业用途。\n"
            "2. 使用者需自行承担因使用本软件产生的全部责任。\n"
            "3. 如不同意上述条款，请立即退出并删除本软件。\n"
            "\n感谢理解与支持。",
        )
        text.configure(state="disabled")
        text.pack(fill="both", expand=True, padx=16, pady=6)

        agree = tk.Checkbutton(
            frame,
            text="我已阅读并同意以上协议",
            variable=self.agree_var,
            command=self._update_step1_buttons,
        )
        agree.pack(pady=6)

        btn_row = tk.Frame(frame)
        btn_row.pack(fill="x", pady=10, padx=16)
        self.step1_next = ttk.Button(btn_row, text="下一步", command=self._step1_next)
        self.step1_next.pack(side="right")
        ttk.Button(btn_row, text="取消", command=self.destroy).pack(side="right", padx=6)
        self._update_step1_buttons()

    def _update_step1_buttons(self):
        state = "normal" if self.agree_var.get() else "disabled"
        self.step1_next.configure(state=state)

    def _build_step2(self):
        frame = tk.Frame(self.container)
        self.frames[2] = frame

        title = tk.Label(frame, text="选择安装位置", font=("Segoe UI", 14, "bold"))
        title.pack(pady=12)

        path_row = tk.Frame(frame)
        path_row.pack(fill="x", padx=16, pady=6)
        tk.Label(path_row, text="安装目录:").pack(side="left")
        entry = ttk.Entry(path_row, textvariable=self.path_var)
        entry.pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(path_row, text="浏览...", command=self._browse_dir).pack(side="left")

        shortcut = tk.Checkbutton(frame, text="创建桌面快捷方式", variable=self.shortcut_var)
        shortcut.pack(anchor="w", padx=16, pady=6)

        btn_row = tk.Frame(frame)
        btn_row.pack(fill="x", pady=10, padx=16)
        ttk.Button(btn_row, text="上一步", command=lambda: self._show_step(1)).pack(side="right")
        ttk.Button(btn_row, text="下一步", command=self._step2_next).pack(side="right", padx=6)
        ttk.Button(btn_row, text="取消", command=self.destroy).pack(side="right", padx=6)

    def _build_step3(self):
        frame = tk.Frame(self.container)
        self.frames[3] = frame

        title = tk.Label(frame, text="正在安装", font=("Segoe UI", 14, "bold"))
        title.pack(pady=12)

        status = tk.Label(frame, textvariable=self.status_var, font=("Segoe UI", 10))
        status.pack(pady=4)

        bar = ttk.Progressbar(frame, maximum=100, variable=self.progress_var)
        bar.pack(fill="x", padx=24, pady=8)

        self.finish_button = ttk.Button(frame, text="完成", command=self._finish, state="disabled")
        self.finish_button.pack(pady=12)

    def _show_step(self, step: int):
        for f in self.frames.values():
            f.pack_forget()
        self.frames[step].pack(fill="both", expand=True)
        if step == 3 and not self.installing and not self.install_done:
            self._start_install()

    def _step1_next(self):
        self._show_step(2)

    def _step2_next(self):
        path = self.path_var.get().strip()
        if not path:
            messagebox.showwarning("提示", "请选择安装目录")
            return
        if os.path.isdir(path) and os.listdir(path):
            ok = messagebox.askyesno("提示", "目标目录非空，是否继续安装？")
            if not ok:
                return
        self._show_step(3)

    def _browse_dir(self):
        initial = self.path_var.get().strip() or self._default_install_dir()
        target = filedialog.askdirectory(initialdir=initial, title="选择安装目录")
        if target:
            self.path_var.set(target)

    def _start_install(self):
        self.installing = True
        self.status_var.set("正在获取最新版本...")
        self.progress_var.set(0)
        thread = threading.Thread(target=self._install_worker, daemon=True)
        thread.start()
        self.after(100, self._poll_queue)

    def _install_worker(self):
        try:
            info = fetch_latest_release()
            self.queue.put(("status", f"正在下载 {info['name']} ({info['source']})"))
            temp_dir = tempfile.mkdtemp(prefix="niuma_installer_")
            zip_path = os.path.join(temp_dir, info["name"] or "package.zip")

            def _progress(done, total):
                self.queue.put(("progress", done, total))

            _download_file(info["url"], zip_path, _progress)
            self.queue.put(("status", "正在解压安装包..."))
            extract_dir = os.path.join(temp_dir, "extracted")
            os.makedirs(extract_dir, exist_ok=True)
            with zipfile.ZipFile(zip_path, "r") as zf:
                zf.extractall(extract_dir)
            root, exe_name = _find_package_root(extract_dir)
            if not root or not exe_name:
                raise RuntimeError("安装包结构异常")
            install_dir = self.path_var.get().strip()
            exe_path = _copy_package(root, exe_name, install_dir)

            self.queue.put(("status", "正在下载配置文件..."))
            _download_default_settings(install_dir)
            _create_first_run_flag(install_dir)

            self.queue.put(("status", "正在安装 AutoHotkey..."))
            if not _ensure_ahk_installed(temp_dir):
                self.queue.put(("status", "AutoHotkey 安装失败，快捷键可能不可用"))

            if self.shortcut_var.get():
                self.queue.put(("status", "正在创建桌面快捷方式..."))
                _create_desktop_shortcut(exe_path)

            self.queue.put(("done", exe_path))
        except Exception as e:
            self.queue.put(("error", str(e)))
        finally:
            self.queue.put(("cleanup",))

    def _poll_queue(self):
        try:
            while True:
                msg = self.queue.get_nowait()
                if not msg:
                    continue
                if msg[0] == "status":
                    self.status_var.set(msg[1])
                elif msg[0] == "progress":
                    done, total = msg[1], msg[2]
                    if total > 0:
                        percent = int(done * 100 / total)
                        self.progress_var.set(percent)
                        self.status_var.set(
                            f"正在下载... {percent}% ({done/1024/1024:.1f}MB/{total/1024/1024:.1f}MB)"
                        )
                    else:
                        self.status_var.set(f"正在下载... {done/1024/1024:.1f}MB")
                elif msg[0] == "done":
                    self.install_done = True
                    self.installing = False
                    self.installed_exe_path = msg[1]
                    self.progress_var.set(100)
                    self.status_var.set("安装完成")
                    self.finish_button.configure(state="normal")
                elif msg[0] == "error":
                    self.installing = False
                    self.status_var.set("安装失败")
                    messagebox.showerror("错误", f"安装失败: {msg[1]}")
                    self.finish_button.configure(state="normal")
                elif msg[0] == "cleanup":
                    pass
        except queue.Empty:
            pass
        if self.installing:
            self.after(100, self._poll_queue)

    def _finish(self):
        if self.install_done and self.installed_exe_path:
            launch = messagebox.askyesno("安装完成", "安装完成，是否立即启动？")
            if launch:
                try:
                    subprocess.Popen([self.installed_exe_path], cwd=os.path.dirname(self.installed_exe_path))
                except Exception:
                    pass
        self.destroy()

    def _on_close(self):
        if self.installing:
            ok = messagebox.askyesno("提示", "正在安装中，确定要退出吗？")
            if not ok:
                return
        self._finish()


if __name__ == "__main__":
    app = InstallerApp()
    app.mainloop()
