@echo off
chcp 65001 >nul
cd /d E:\BibCatalog

if exist publish_log.txt del publish_log.txt

echo ============================================================
echo   Publishing library catalog SHPSA
echo   %date% %time%
echo ============================================================
echo.

echo [1/4] Rebuilding index.html ...
python update_catalog_from_xlsx.py Каталог.xlsx catalog_template.html index.html

echo.
echo [2/4] git add ...
git add -A

echo.
echo [3/4] git commit ...
git commit -m "Update catalog"

echo.
echo [4/4] git push ...
git push

echo.
echo ============================================================
echo   Done! Site will update in 1-2 minutes:
echo   https://biblioteka-shpsa.pages.dev
echo ============================================================
echo.
pause