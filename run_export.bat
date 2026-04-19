@echo off
setlocal
cd /d %~dp0
python data_api\export_master_jsons.py
if errorlevel 1 (
  echo.
  echo El exportador finalizo con error. Revisa la consola y el manifest si existe.
  pause
  exit /b %errorlevel%
)
echo.
echo Exportacion completada. Revisa data_api\out\meta\manifest.json
pause
