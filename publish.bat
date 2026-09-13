@echo off
chcp 65001 >nul
cd /d E:\BibCatalog

if exist publish_log.txt del publish_log.txt

echo ============================================================ >> publish_log.txt
echo   Публикация каталога библиотеки ШПСА >> publish_log.txt
echo   %date% %time% >> publish_log.txt
echo ============================================================ >> publish_log.txt
echo. >> publish_log.txt

echo ============================================================
echo   Публикация каталога библиотеки ШПСА
echo ============================================================
echo.

echo [1/4] Пересборка index.html из Каталог.xlsx ...
python update_catalog_from_xlsx.py Каталог.xlsx catalog_template.html index.html >> publish_log.txt 2>&1

echo [2/4] git add ...
git add -A >> publish_log.txt 2>&1

echo [3/4] git commit ...
git commit -m "Обновление каталога" >> publish_log.txt 2>&1

echo [4/4] git push ...
git push >> publish_log.txt 2>&1

echo.
echo ============================================================
echo   Готово! Сайт обновится через 1-2 минуты:
echo   https://biblioteka-shpsa.pages.dev
echo ============================================================
echo.
pause