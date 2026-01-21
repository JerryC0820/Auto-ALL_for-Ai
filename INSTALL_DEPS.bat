@echo off
setlocal
cd /d "%~dp0"
py -m pip install -U pip
py -m pip install -U selenium PySide6 pillow
pause
