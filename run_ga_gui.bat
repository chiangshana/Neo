@echo off
chcp 65001 >nul
cd /d "%~dp0"
python -B ga_tkinter_app.py
pause
