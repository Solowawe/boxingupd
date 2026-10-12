@echo off
rem ============================================================
rem   Установка приложения «Соревнования по боксу» на новом ПК
rem   Запустите этот файл ОДИН раз: создаст venv и поставит зависимости
rem ============================================================
cd /d "%~dp0"
echo.
echo  Установка зависимостей...
if not exist ".venv\Scripts\python.exe" (
    echo  Создаю виртуальное окружение...
    python -m venv .venv
    if errorlevel 1 (
        echo  ОШИБКА: не найден Python. Установите Python 3.9+ с python.org
        echo  и поставьте галочку "Add to PATH".
        pause
        exit /b 1
    )
)
echo  Устанавливаю пакеты...
.venv\Scripts\python.exe -m pip install --upgrade pip
.venv\Scripts\python.exe -m pip install -r requirements.txt
echo.
echo  ✅ Готово! Запустите run.bat чтобы стартовать сервер.
echo  Сайт откроется по адресу http://localhost:8080
echo.
pause