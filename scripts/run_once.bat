@echo off
chcp 936 >nul
cd /d "%~dp0.."
python -m src.watcher --once
pause
