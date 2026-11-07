@echo off
chcp 936 >nul
cd /d "%~dp0.."
echo ============================================
echo  Feishu - ERP Bridge  (loop mode)
echo  Ctrl+C to stop
echo ============================================
python -m src.watcher --loop
pause
