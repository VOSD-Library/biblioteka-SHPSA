@echo off
chcp 65001 >nul
cd /d E:\BibCatalog

echo ============================================================
echo   Публикация каталога библиотеки ШПСА
echo ============================================================
echo.

echo [1/4] Пересборка index.html из Каталог.xlsx ...
python update_catalog_from_xlsx.py Каталог.xlsx catalog_template.html index.html

echo.
echo [2/4] git add ...
git add -A

echo.
echo [3/4] git commit ...
git commit -m "Обновление каталога"

echo.
echo [4/4] git push ...
git push

echo.
echo ============================================================
echo   Готово! Сайт обновится через 1-2 минуты:
echo   https://biblioteka-shpsa.pages.dev
echo ============================================================
echo.
pause