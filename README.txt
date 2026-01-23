牛马神器 4.0.10 说明

更新内容
- 缺少配置时弹窗选择部署位置（当前目录/新目录），可自动迁移
- 缺少配置且有新版本时强制更新，支持安装到新目录
- 更新源可在“关于”内切换（国内优先）
- 新增轻量安装器脚本（可选择安装位置并下载最新包）

快速开始
1) 安装 Python 3.8+（建议 64 位）。
2) 运行 INSTALL_DEPS.bat 安装依赖。
3) 运行 RUN.bat 启动。

国内下载/镜像（可选）
- 镜像地址：https://gitee.com/chen-bin98/Auto-ALL_for-Ai
- 终端克隆（GitHub）：git clone https://github.com/JerryC0820/Auto-ALL_for-Ai.git
- 终端克隆（Gitee）：git clone https://gitee.com/chen-bin98/Auto-ALL_for-Ai.git
- 进入目录：cd Auto-ALL_for-Ai

终端一键安装/启动（在仓库目录内执行）
py -3 -m pip install -r requirements.txt && py 牛马神器_v4.0.10.py

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
