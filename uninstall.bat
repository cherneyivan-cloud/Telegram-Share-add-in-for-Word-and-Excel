@echo off
chcp 65001 >nul
echo Удаление Telegram Share...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\uninstall.ps1" %*
echo.
pause
