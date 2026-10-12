@echo off
title 🥊 Соревнования по боксу
cd /d "%~dp0"
echo ============================================================
echo   🥊 СОРЕВНОВАНИЯ ПО БОКСУ — ЗАПУСК
echo ============================================================
echo.
echo  Запуск сервера...
start /MIN "Boxing Server" cmd /c "cd /d "%~dp0" && .venv\Scripts\python.exe server.py"
echo.
echo  ✅ Сервер запущен!
echo  🌐 Откройте браузер: http://localhost:8080
echo.
echo  Сервер работает в отдельном окне.
echo  Закройте окно сервера для остановки.
echo.
pause