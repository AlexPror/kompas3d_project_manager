@echo off
chcp 65001 > nul
echo ========================================
echo  Анализ конфигурации чертежа
echo ========================================
echo.
echo 🔍 Запуск анализа...
echo.

python analyze_drawing_config.py

echo.
echo ========================================
echo Готово! Результаты в drawing_config_analysis.txt
pause > nul
