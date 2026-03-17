@echo off
chcp 65001 > nul
echo.
echo  Запуск сервера водителей...
echo.
cd /d "%~dp0"
node server.js
pause
