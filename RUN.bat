@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"
set "PATH=%~dp0;%PATH%"

set "AHK_EXE="
if exist "C:\Program Files\AutoHotkey\v2\AutoHotkey.exe" set "AHK_EXE=C:\Program Files\AutoHotkey\v2\AutoHotkey.exe"
if exist "C:\Program Files\AutoHotkey\AutoHotkey.exe" set "AHK_EXE=C:\Program Files\AutoHotkey\AutoHotkey.exe"
if exist "C:\Program Files (x86)\AutoHotkey\AutoHotkey.exe" set "AHK_EXE=C:\Program Files (x86)\AutoHotkey\AutoHotkey.exe"

if not defined AHK_EXE (
  where /q winget
  if errorlevel 1 (
    echo AutoHotkey not found. Install from https://www.autohotkey.com/
  ) else (
    echo Installing AutoHotkey...
    winget install -e --id AutoHotkey.AutoHotkey --source winget --accept-package-agreements --accept-source-agreements --silent
  )
)

set "APP_EXE="
for /f "delims=" %%F in ('dir /b /s /o:-d "%~dp0dist*\\牛马神器_v4.0.11.exe" 2^>nul') do (
  set "APP_EXE=%%F"
  goto :runexe
)
:runexe
if defined APP_EXE (
  start "" "%APP_EXE%"
  exit /b 0
)

if exist "牛马神器_v4.0.11.py" (
  where /q py
  if not errorlevel 1 (
    start "" py "牛马神器_v4.0.11.py"
    exit /b 0
  )
  where /q python
  if not errorlevel 1 (
    start "" python "牛马神器_v4.0.11.py"
    exit /b 0
  )
)

echo App not found.
pause
