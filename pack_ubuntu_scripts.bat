@echo off
setlocal EnableExtensions

cd /d "%~dp0"
set "PROJECT_ROOT=%cd%"
set "OUT_DIR=%PROJECT_ROOT%\dist\SyntheticMonitoringTool-scripts"
set "ZIP_PATH=%PROJECT_ROOT%\dist\SyntheticMonitoringTool-scripts.zip"

set "FILES=probe_core.py config_loader.py runtime_env.py run_cli.py run_cli.sh setup_ubuntu.sh install_server_deps.sh ubuntu_apt_deps.sh pw_browser_sync.py config.example.json requirements-linux.txt DEPLOY-ubuntu.md"

echo [1/3] Preparing script package directory...
if exist "%OUT_DIR%" rmdir /s /q "%OUT_DIR%"
mkdir "%OUT_DIR%" 2>nul

echo [2/3] Copying scripts and config files...
set "MISSING="
for %%F in (%FILES%) do (
  if not exist "%PROJECT_ROOT%\%%F" (
    echo Missing file: %%F
    set "MISSING=1"
  ) else (
    copy /Y "%PROJECT_ROOT%\%%F" "%OUT_DIR%\" >nul
  )
)
if defined MISSING goto :error

echo [2.5/3] Normalizing shell scripts to Unix line endings (LF)...
python "%PROJECT_ROOT%\fix_sh_lf.py" "%OUT_DIR%"
if errorlevel 1 goto :error

echo [3/3] Creating zip archive...
if exist "%ZIP_PATH%" del /f /q "%ZIP_PATH%"
powershell -NoProfile -Command "Compress-Archive -Path '%OUT_DIR%\*' -DestinationPath '%ZIP_PATH%' -Force"
if errorlevel 1 goto :error

echo.
echo Package created:
echo   Folder: %OUT_DIR%
echo   Zip:    %ZIP_PATH%
echo.
echo Deploy on Ubuntu:
echo   unzip SyntheticMonitoringTool-scripts.zip -d /opt/SyntheticMonitoringTool
echo   cd /opt/SyntheticMonitoringTool
echo   chmod +x setup_ubuntu.sh run_cli.sh
echo   ./setup_ubuntu.sh
echo   cp config.example.json config.json
echo   ./run_cli.sh
goto :eof

:error
echo.
echo pack_ubuntu_scripts.bat failed.
exit /b 1
