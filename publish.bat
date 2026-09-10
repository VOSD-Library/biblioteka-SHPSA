@echo off
chcp 65001 >nul
cd /d E:\BibCatalog
echo ============================== >> publish_log.txt
echo %date% %time% >> publish_log.txt
python publish.py
echo.
echo Готово. Нажмите любую клавишу для выхода...
pause >nul