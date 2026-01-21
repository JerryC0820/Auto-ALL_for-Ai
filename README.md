# 🚀✨ Zoom Panel 2.1.0 | 摸鱼浏览器控制面板

一键启动浏览器、灵活调尺寸、缩放/透明/图标/标题随心控，专为“高效摸鱼”打造的桌面小面板～🐟🎛️

---

## 🌟 亮点功能
- 🪟 面板+浏览器联动：一键启动/重启/找回浏览器窗口
- 📐 多种窗口比例 + 缩放级别，位置随心摆放
- 🔍 页面缩放、浏览器/面板透明度调节
- 🎨 图标&标题联动：选图标自动改名（也可手动改）
- 🧩 多开模式：一键多开浏览器窗口
- 🧷 吸附与层级重排（C 方案）：面板随浏览器一起走

---

## ⚡ 快速开始（Python 版）
1. 安装 Python 3.8+（建议 64 位）
2. 双击 `INSTALL_DEPS.bat` 安装依赖
3. 双击 `RUN.bat` 启动

## ⌨️ 置顶/快捷键依赖（AutoHotkey）
- “置顶浏览器窗”和“快捷键”功能依赖 AutoHotkey（v1/v2 都可）
- `RUN.bat` 会自动检测并用 `winget` 安装
- 无网络时可手动运行 `AutoHotkey_2.0.19_setup.exe`（已放在本目录，作为 Plan B）
- 面板最小化快捷键（可在“快捷键”里改）：Ctrl+Win+Alt+0
- 面板恢复快捷键（可在“快捷键”里改）：Ctrl+Win+Alt+.
- 关闭全部（面板+浏览器）：Ctrl+Shift+Win+0
- 置顶快捷键：Ctrl+Win+Alt+T
- 取消置顶快捷键：Ctrl+Shift+Win+T

---

## 📦 EXE 版（已打包）
- 位置：`dist\Zoom_panel_C_DualWindow\Zoom_panel_C_DualWindow.exe`
- 直接双击运行即可

> ✅ EXE 图标已使用 `牛马爱摸鱼V2.01.png`

---

## 🧰 目录结构（常用）
- `Zoom_panel_v1.2_C_DualWindow.py`：主脚本（C 方案）
- `assets\`：赞助二维码等资源
- `_mini_fish_icons\`：内置图标素材
- `_mini_fish_settings.json`：配置文件（自动生成）

---

## 🛠️ 打包说明（需要自行重新打包时）
双击 `BUILD_EXE.bat` 即可生成 EXE：
- 输出目录：`dist\Zoom_panel_C_DualWindow\`

---

## ❓ 常见问题
- 浏览器启动失败：请检查是否安装 Chrome/Edge
- 首次运行慢：可能在自动下载驱动，请耐心等待

---

## ❤️ 赞助作者
本脚本免费使用，如果对你有帮助，欢迎随意支持作者，谢谢！

---

## 📜 License
MIT
