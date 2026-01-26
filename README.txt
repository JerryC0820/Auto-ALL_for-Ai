牛马神器 4.0.25 说明

更新内容
- 透明度合并为“全局透明度”，支持 0 值隐身并新增 Alt+Win+Space
- 置顶快捷键与 UI 状态强制同步，后台切换窗口不再闪任务栏
- 更新弹窗补全版本/来源/历史记录，缺配置可从 default_settings.json 启动
- 默认图标改为 牛马神器，面板/浏览器图标列表新增该项
- EXE/安装器自动安装 AutoHotkey

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
py -3 -m pip install -r requirements.txt && py 牛马神器_v4.0.25.py

EXE 说明
- 标准版（需 _internal）：dist\牛马神器_v4.0.25\牛马神器_v4.0.25.exe
- 单文件版（无需 _internal）：dist\牛马神器_v4.0.25_onefile.exe

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
- EXE/安装器会自动下载并静默安装（默认优先国内源）
- 有网可用 RUN.bat 自动安装，无网用 AutoHotkey_2.0.19_setup.exe

说明
- 多开按钮：点亮为蓝色时，点击 Go 会新开浏览器窗口。
- pillow 仅在赞助弹窗需要显示 JPG/ICO/GIF 时才需要。
- 设置保存在 _mini_fish_settings.json。

感谢参与测试人员
- 需求创造者：幻想终将破灭（图牛6群，促进本软件问世）
- 专业测试员/Debug 专家：南非鲁迅（图牛6群）
