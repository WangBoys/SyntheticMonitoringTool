@echo off
setlocal

cd /d "%~dp0"
set "PROJECT_ROOT=%cd%"
set "SHORT_DIST=%PROJECT_ROOT%\dist"
set "SHORT_WORK=%PROJECT_ROOT%\build"
set "BUNDLE_DIR="
set "ISCC_EXE="
set "FALLBACK_USED=0"
set "FORCE_DEPS_INSTALL=0"
set "FORCE_BROWSER_INSTALL=0"
set "FULL_CLEAN=0"
set "SKIP_INSTALLER=0"
set "ADD_DEFENDER_EXCLUSIONS=0"
set "PY_SITE_PACKAGES="
set "ICON_PNG=%PROJECT_ROOT%\ico.png"
set "ICON_ICO=%PROJECT_ROOT%\app_icon.ico"

for %%A in (%*) do (
  if /I "%%~A"=="--force-deps" set "FORCE_DEPS_INSTALL=1"
  if /I "%%~A"=="--force-browser" set "FORCE_BROWSER_INSTALL=1"
  if /I "%%~A"=="--full-clean" set "FULL_CLEAN=1"
  if /I "%%~A"=="--skip-installer" set "SKIP_INSTALLER=1"
  if /I "%%~A"=="--add-defender-exclusions" set "ADD_DEFENDER_EXCLUSIONS=1"
)

set "BUILD_MODE=INCREMENTAL"
if "%FULL_CLEAN%"=="1" set "BUILD_MODE=FULL_CLEAN"
echo [Mode] %BUILD_MODE%
call :get_python_site_packages
if not defined PY_SITE_PACKAGES (
  echo Warning: failed to detect Python site-packages path.
)
echo [Tip] To speed up build on Windows Defender, add exclusions:
echo   "%PROJECT_ROOT%\build"
echo   "%PROJECT_ROOT%\dist"
echo   "%PROJECT_ROOT%\pw-browsers"
if defined PY_SITE_PACKAGES echo   "%PY_SITE_PACKAGES%"
if "%ADD_DEFENDER_EXCLUSIONS%"=="1" call :add_defender_exclusions

echo [0/5] Checking file/process occupation...
tasklist /FI "IMAGENAME eq SyntheticMonitoringTool.exe" | find /I "SyntheticMonitoringTool.exe" >nul
if not errorlevel 1 (
  echo Detected running process: SyntheticMonitoringTool.exe
  echo Please close the app before running build_exe.bat
  goto :error
)

echo [1/5] Checking build dependencies...
if "%FORCE_DEPS_INSTALL%"=="1" goto :install_deps
python -c "import PySide6,pandas,openpyxl,playwright,msoffcrypto,PIL,PyInstaller" >nul 2>nul
if errorlevel 1 goto :install_deps
echo Dependencies already satisfied, skip pip install.
goto :deps_ready

:install_deps
echo Installing/repairing Python dependencies...
python -m pip install -r requirements.txt
if errorlevel 1 goto :error

:deps_ready

if not exist "%ICON_PNG%" (
  echo Missing icon source file: %ICON_PNG%
  goto :error
)
python -c "from PIL import Image; img=Image.open(r'%ICON_PNG%').convert('RGBA'); img.save(r'%ICON_ICO%', format='ICO', sizes=[(256,256),(128,128),(64,64),(48,48),(32,32),(16,16)])"
if errorlevel 1 (
  echo Failed to convert ico.png to app_icon.ico
  goto :error
)

echo [2/5] Syncing Playwright Chromium into pw-browsers...
set "SYNC_FORCE_FLAG="
if "%FORCE_BROWSER_INSTALL%"=="1" set "SYNC_FORCE_FLAG=--force"
python "%PROJECT_ROOT%\pw_browser_sync.py" sync --path "%PROJECT_ROOT%\pw-browsers" %SYNC_FORCE_FLAG%
if errorlevel 1 goto :error

echo [3/5] Preparing build paths...
call :prepare_paths
if errorlevel 2 (
  echo Detected locked dist/build output. Switching to fallback paths in project root...
  set "FALLBACK_USED=1"
  set "SHORT_DIST=%PROJECT_ROOT%\dist_fallback_%RANDOM%%RANDOM%"
  set "SHORT_WORK=%PROJECT_ROOT%\build_fallback_%RANDOM%%RANDOM%"
  call :prepare_paths
  if errorlevel 1 goto :error
)
if errorlevel 1 goto :error

echo [4/5] Building executable...
set "PYI_CLEAN_FLAG="
if "%FULL_CLEAN%"=="1" set "PYI_CLEAN_FLAG=--clean"
python -m PyInstaller --noconfirm %PYI_CLEAN_FLAG% --distpath "%SHORT_DIST%" --workpath "%SHORT_WORK%" SyntheticMonitoringTool.spec
if errorlevel 1 goto :error

for /f %%i in ('python -c "import sys; print(f'python{sys.version_info.major}{sys.version_info.minor}.dll')"') do set "PY_DLL_NAME=%%i"
for /f %%i in ('python -c "import pathlib,sys; print(pathlib.Path(sys.executable).parent)"') do set "PY_HOME=%%i"
if not exist "%SHORT_DIST%\SyntheticMonitoringTool\_internal\%PY_DLL_NAME%" (
  copy /y "%PY_HOME%\%PY_DLL_NAME%" "%SHORT_DIST%\SyntheticMonitoringTool\_internal\%PY_DLL_NAME%"
  if errorlevel 1 goto :error
)

set "BUNDLE_DIR=%SHORT_DIST%\SyntheticMonitoringTool"
python "%PROJECT_ROOT%\pw_browser_sync.py" check --path "%BUNDLE_DIR%\_internal\pw-browsers"
if errorlevel 1 (
  echo Missing bundled Playwright Chromium in %BUNDLE_DIR%\_internal\pw-browsers.
  goto :error
)

if "%SKIP_INSTALLER%"=="1" goto :done_no_installer

echo [5/5] Building installer exe...
if exist "%USERPROFILE%\AppData\Local\Programs\Inno Setup 6\ISCC.exe" set "ISCC_EXE=%USERPROFILE%\AppData\Local\Programs\Inno Setup 6\ISCC.exe"
if not defined ISCC_EXE if exist "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" set "ISCC_EXE=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not defined ISCC_EXE if exist "%ProgramFiles%\Inno Setup 6\ISCC.exe" set "ISCC_EXE=%ProgramFiles%\Inno Setup 6\ISCC.exe"

if defined ISCC_EXE goto :found_iscc

if not defined ISCC_EXE (
  for /f "delims=" %%i in ('where ISCC 2^>nul') do (
    set "ISCC_EXE=%%i"
  )
)

if defined ISCC_EXE goto :found_iscc
goto :need_iscc

:found_iscc
"%ISCC_EXE%" /DSourceDir="%BUNDLE_DIR%" /DSetupIcon="%ICON_ICO%" "%PROJECT_ROOT%\SyntheticMonitoringToolInstaller.iss"
if errorlevel 1 goto :error

:done_no_installer
echo Done.
echo EXE path: %BUNDLE_DIR%\SyntheticMonitoringTool.exe
if not "%SKIP_INSTALLER%"=="1" (
  echo Installer path: installer\SyntheticMonitoringTool-Setup.exe
) else (
  echo Installer build skipped by --skip-installer.
)
if "%FALLBACK_USED%"=="1" (
  echo NOTE: dist/build was locked; this build used fallback paths:
  echo   dist: %SHORT_DIST%
  echo   build: %SHORT_WORK%
)
exit /b 0

:get_python_site_packages
for /f "delims=" %%i in ('python -c "import site,sysconfig; paths=site.getsitepackages(); c=[p for p in paths if 'site-packages' in p.lower()]; print(c[0] if c else sysconfig.get_paths().get('purelib',''))" 2^>nul') do set "PY_SITE_PACKAGES=%%i"
exit /b 0

:add_defender_exclusions
if /I not "%OS%"=="Windows_NT" exit /b 0
echo [Defender] Adding exclusion paths (may require Administrator)...
set "DEFENDER_PS=$ErrorActionPreference='Stop'; $paths=@('%PROJECT_ROOT%\build','%PROJECT_ROOT%\dist','%PROJECT_ROOT%\pw-browsers'"
if defined PY_SITE_PACKAGES set "DEFENDER_PS=%DEFENDER_PS%,'%PY_SITE_PACKAGES%'"
set "DEFENDER_PS=%DEFENDER_PS%); foreach ($p in $paths) { if (Test-Path $p) { Add-MpPreference -ExclusionPath $p -ErrorAction Stop } }; Write-Host 'Defender exclusions added.'"
powershell -NoProfile -ExecutionPolicy Bypass -Command "%DEFENDER_PS%"
if errorlevel 1 (
  echo [Defender] Failed to add exclusions automatically. Run this script as Administrator or add exclusions manually.
)
exit /b 0

:prepare_paths
if not exist "%SHORT_DIST%" mkdir "%SHORT_DIST%"
if errorlevel 1 exit /b 1
if not exist "%SHORT_WORK%" mkdir "%SHORT_WORK%"
if errorlevel 1 exit /b 1

rem Pre-clean old build outputs to avoid PyInstaller lock errors.
if exist "%SHORT_DIST%\SyntheticMonitoringTool\SyntheticMonitoringTool.exe" (
  del /f /q "%SHORT_DIST%\SyntheticMonitoringTool\SyntheticMonitoringTool.exe" >nul 2>nul
  if exist "%SHORT_DIST%\SyntheticMonitoringTool\SyntheticMonitoringTool.exe" exit /b 2
)
if exist "%SHORT_DIST%\SyntheticMonitoringTool" (
  rmdir /s /q "%SHORT_DIST%\SyntheticMonitoringTool" >nul 2>nul
  if exist "%SHORT_DIST%\SyntheticMonitoringTool" exit /b 2
)
if "%FULL_CLEAN%"=="1" (
  if exist "%SHORT_WORK%\SyntheticMonitoringTool" (
    rmdir /s /q "%SHORT_WORK%\SyntheticMonitoringTool" >nul 2>nul
    if exist "%SHORT_WORK%\SyntheticMonitoringTool" exit /b 2
  )
)
exit /b 0

:need_iscc
echo Inno Setup Compiler (ISCC) not found.
echo Please install Inno Setup 6 first:
echo   winget install --id JRSoftware.InnoSetup -e --silent --scope user --accept-source-agreements --accept-package-agreements
pause
exit /b 1

:error
echo Build failed. Please check the errors above.
pause
exit /b 1
