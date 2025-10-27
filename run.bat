@echo off
chcp 65001 > nul
echo ========================================
echo  КОМПАС-3D Project Manager
echo ========================================
echo.

REM Проверка наличия Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Python не найден! Установите Python 3.8+
    pause
    exit /b 1
)

echo ✅ Python найден
echo.
echo 📦 Проверка зависимостей...

REM Установка зависимостей если нужно
pip show customtkinter >nul 2>&1
if errorlevel 1 (
    echo.
    echo 📥 Установка зависимостей...
    pip install -r requirements.txt
    if errorlevel 1 (
        echo ❌ Ошибка установки зависимостей!
        pause
        exit /b 1
    )
)

echo ✅ Зависимости установлены
echo.
echo 🚀 Запуск приложения...
echo.

python gui_kompas_manager.py

if errorlevel 1 (
    echo.
    echo ❌ Ошибка запуска!
    pause
)

