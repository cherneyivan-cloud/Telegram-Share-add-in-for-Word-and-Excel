@echo off
chcp 65001 >nul
echo Установка Telegram Share (Word + Excel)...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\install.ps1" %*
echo.
pause
