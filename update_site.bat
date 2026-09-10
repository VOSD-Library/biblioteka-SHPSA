@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
cd /d E:\BibCatalog
echo ============================== >> update_log.txt
echo %date% %time% >> update_log.txt
python update_catalog_from_xlsx.py Каталог.xlsx catalog_template.html index.html >> update_log.txt 2>&1

echo Попытка деплоя #1 >> update_log.txt
netlify deploy --prod --dir=E:\BibCatalog >> update_log.txt 2>&1
if %ERRORLEVEL% EQU 0 goto DONE

echo Деплой #1 не удался, жду 30 секунд... >> update_log.txt
timeout /t 30 /nobreak >nul
echo Попытка деплоя #2 >> update_log.txt
netlify deploy --prod --dir=E:\BibCatalog >> update_log.txt 2>&1
if %ERRORLEVEL% EQU 0 goto DONE

echo Деплой #2 не удался, жду 30 секунд... >> update_log.txt
timeout /t 30 /nobreak >nul
echo Попытка деплоя #3 >> update_log.txt
netlify deploy --prod --dir=E:\BibCatalog >> update_log.txt 2>&1

:DONE
echo Готово. >> update_log.txt
