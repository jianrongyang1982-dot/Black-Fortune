@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo  WorkPrice Importer - Build EXE
echo ============================================
set "PY=%~dp0.venv\Scripts\python.exe"
if not exist "%PY%" (
  echo [WARN] .venv not found, falling back to system python
  set "PY=python"
)
echo [1/3] Installing runtime dependencies...
"%PY%" -m pip install -r requirements.txt --quiet
echo [2/3] Installing PyInstaller...
"%PY%" -m pip install pyinstaller --quiet
echo [3/3] Building (onedir / no console)...
REM Back up the previous build so PyInstaller does not wipe the whole dist folder
if exist "dist\WorkPrice-Importer" (
  move "dist\WorkPrice-Importer" "dist\WorkPrice-Importer.bak.%RANDOM%" >nul
  echo Previous build backed up as dist\WorkPrice-Importer.bak.*
)
"%PY%" -m PyInstaller --noconfirm --clean --onedir --noconsole ^
  --name WorkPrice-Importer ^
  --hidden-import pymysql ^
  --hidden-import pymysql.cursors ^
  import_workprice.py
echo.
if exist "dist\WorkPrice-Importer\WorkPrice-Importer.exe" (
  copy /Y "db_config.json" "dist\WorkPrice-Importer\db_config.json" >nul
  echo Build complete: dist\WorkPrice-Importer\WorkPrice-Importer.exe
  echo db_config.json is copied next to the exe and can be edited directly.
  echo Distribute the WHOLE dist\WorkPrice-Importer folder, including _internal.
) else (
  echo Build failed, check the output above.
)
pause
