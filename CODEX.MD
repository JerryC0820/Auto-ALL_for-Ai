# CODEX Rules

## Workflow (must follow)
- 每次改动前确认是否需要更新版本号；用户若明确“不改版本号”，按其指示。
- 改动后先复制 Python 版到 `D:\用户文件\Desktop\摸鱼浏览器\Zoom_panel_2.1.0\测试` 并启动供用户验证。
- 用户确认 OK 后再执行：打包 EXE、生成源码/运行包压缩、更新 GitHub 覆盖/Release。
- 每次新增规则、错误或教训都追加到本文件。
- 打包包体不要包含 dist 目录；压缩包根目录直接放 exe 和 `_internal`。
- `Package/` 目录仅用于本地上传备份，不提交到仓库。

## Logs
- 2026-01-22: 用户要求固定流程：先测试再打包/上传；每次询问是否更新版本号。
- 2026-01-22: 新增打包结构要求（zip 根目录放 exe + _internal），并新增本地 Package 目录规则。
