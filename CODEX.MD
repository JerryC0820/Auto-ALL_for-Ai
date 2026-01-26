![alt text](image.png)# CODEX Rules

## Workflow (must follow)
- 每次改动前确认是否需要更新版本号；用户若明确“不改版本号”，按其指示。
- 改动主 Python 后先删除旧测试版本，再复制到 `D:\用户文件\Desktop\项目开发-JERRY&Codex\牛马神器\测试` 并启动供用户验证。
- 用户确认 OK 后再执行：打包 EXE、生成源码/运行包压缩、更新 GitHub 覆盖/Release。
- 每次新增规则、错误或教训都追加到本文件。
- 打包包体不要包含 dist 目录；压缩包根目录直接放 exe 和 `_internal+default_settings.json`。
- `Package/` 目录仅用于本地上传备份，不提交到仓库。
- 测试目录内容规则：仅放主脚本、`assets`、`default_settings.json`，以及同级的 EXE + `_internal+default_settings.json`；测试脚本/EXE 版本号固定为 v2.0.10，文件命名为 `v2.0.10测试版_vx.x.x`（vx.x.x 为最新版版本号），并额外放最新版安装器 EXE 供测试。
- 每次打包完成后删除 `build/`、`_build_tmp/`、`_dist_tmp/` 与所有 `.spec`；`dist/` 与 `Package/` 仅保留最新版本。
- 打包 EXE 完成后把成品（运行文件夹、安装器 EXE、单文件 EXE）统一放入 `Application/`；打包前先清理旧成品，`dist/` 仅作为临时目录，复制完成后删除。

## Logs
- 2026-01-22: 用户要求固定流程：先测试再打包/上传；每次询问是否更新版本号。
- 2026-01-22: 新增打包结构要求（zip 根目录放 exe + _internal），并新增本地 Package 目录规则。
- 2026-01-22: 测试目录改为项目根目录下 `测试`，每次测试前需先删除旧测试版本。
- 2026-01-23: Chrome 启动后 launcher 进程会很快退出，不能仅靠 `proc.poll()` 判断存活，应以 DevTools 端口/窗口为准。
- 2026-01-23: 启动失败会残留占用 profile 的浏览器进程，需按 `--user-data-dir` 清理避免失控窗口。
- 2026-01-23 17:08: Gitee Release 上传曾遇限流，当前改动尚未重新上传；后续任务需提醒补传。
- 2026-01-23: 更新流程要求强制关闭进程、清理缓存并更新桌面快捷方式。
- 2026-01-23: 更新脚本需删除旧 `_internal`/`assets` 后替换，并清理 `_mini_fish_update` 缓存目录。
- 2026-01-23: PowerShell 更新脚本执行易被阻断，改用 cmd + VBS 更新器以确保覆盖替换与快捷方式更新。
- 2026-01-24: 本次允许跳过用户手动测试，但仍需清理测试目录并放入测试包；后续继续按“先测试后打包/上传”的流程。
- 2026-01-24: Gitee Release 创建成功但附件上传失败（code 40001 登录失效），需有效 Gitee Token 重新上传附件。
- 2026-01-26: EXE 内不应执行 pip 安装；依赖安装失败时允许跳过（仅 Selenium），缺少 PySide6 则提示并退出。
- 2026-01-26: 打包后清理 build/spec，dist/Package 仅保留最新版本。
- 2026-01-26: 新增 Application 成品目录规则，dist 改为临时目录。
- 2026-01-26: Gitee Release 附件上传提示仓库附件配额已满（1GB），需先在 Gitee 页面手动删除旧附件再上传。
- 2026-01-27: Gitee Release 附件上传再次提示仓库附件配额已满（1GB），需清理附件后再补传。
