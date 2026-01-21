Zoom Panel 2.1.0 说明

快速开始
1) 安装 Python 3.8+（建议 64 位）。
2) 运行 INSTALL_DEPS.bat 安装依赖。
3) 运行 RUN.bat 启动。

运行环境
- Windows 10/11
- Python 3.8+（系统可用 py 命令）
- 已安装 Chrome 或 Edge
- 首次运行需要联网，Selenium Manager 会下载浏览器驱动

离线/其他电脑
- 如果目标电脑无法联网，请把匹配版本的驱动放到本目录后运行 RUN.bat：
  chromedriver.exe（Chrome）或 msedgedriver.exe（Edge）
  RUN.bat 会把当前目录加入 PATH

快捷键（默认）
- 最小化：Ctrl+Win+Alt+0
- 恢复：Ctrl+Win+Alt+.
- 关闭全部（面板+浏览器）：Ctrl+Shift+Win+0
- 置顶：Ctrl+Win+Alt+T
- 取消置顶：Ctrl+Shift+Win+T

AutoHotkey
- 置顶与快捷键依赖 AutoHotkey（v1/v2 都可）
- 有网可用 RUN.bat 自动安装，无网用 AutoHotkey_2.0.19_setup.exe

说明
- 多开按钮：点亮为蓝色时，点击 Go 会新开浏览器窗口。
- pillow 仅在赞助弹窗需要显示 JPG/ICO/GIF 时才需要。
- 设置保存在 _mini_fish_settings.json。
