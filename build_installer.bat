@echo off
setlocal EnableExtensions

cd /d "%~dp0"
if errorlevel 1 goto :preflight_error

set "PROJECT_ROOT=%cd%"
set "ISCC_EXE="

echo [Preflight] Checking build environment...

where python >nul 2>nul
if errorlevel 1 goto :err_python

if not exist "%PROJECT_ROOT%\build_exe.bat" goto :err_missing_build_exe
if not exist "%PROJECT_ROOT%\SyntheticMonitoringTool.spec" goto :err_missing_spec
if not exist "%PROJECT_ROOT%\SyntheticMonitoringToolInstaller.iss" goto :err_missing_iss

python -c "import PySide6,pandas,openpyxl,playwright,msoffcrypto,PIL" >nul 2>nul
if errorlevel 1 goto :err_deps

python -m PyInstaller --version >nul 2>nul
if errorlevel 1 goto :err_pyinstaller

call :find_iscc
if errorlevel 1 goto :err_iscc

echo [Preflight] OK. Starting build...
echo Building installer via build_exe.bat...
call "%PROJECT_ROOT%\build_exe.bat"
if errorlevel 1 goto :error

exit /b 0

:find_iscc
set "ISCC_EXE="
if exist "%USERPROFILE%\AppData\Local\Programs\Inno Setup 6\ISCC.exe" set "ISCC_EXE=%USERPROFILE%\AppData\Local\Programs\Inno Setup 6\ISCC.exe"
if defined ISCC_EXE exit /b 0
if exist "%ProgramFiles%\Inno Setup 6\ISCC.exe" set "ISCC_EXE=%ProgramFiles%\Inno Setup 6\ISCC.exe"
if defined ISCC_EXE exit /b 0
if exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" set "ISCC_EXE=C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if defined ISCC_EXE exit /b 0
for /f "delims=" %%i in ('where ISCC 2^>nul') do set "ISCC_EXE=%%i"
if defined ISCC_EXE exit /b 0
exit /b 1

:err_python
echo Python not found in PATH.
goto :preflight_error

:err_missing_build_exe
echo Missing file: build_exe.bat
goto :preflight_error

:err_missing_spec
echo Missing file: SyntheticMonitoringTool.spec
goto :preflight_error

:err_missing_iss
echo Missing file: SyntheticMonitoringToolInstaller.iss
goto :preflight_error

:err_deps
echo Missing Python dependencies. Please run: pip install -r requirements.txt
goto :preflight_error

:err_pyinstaller
echo PyInstaller is unavailable. Please install it first.
goto :preflight_error

:err_iscc
echo Inno Setup Compiler (ISCC) not found.
echo Install command:
echo   winget install --id JRSoftware.InnoSetup -e --silent --scope user --accept-source-agreements --accept-package-agreements
goto :preflight_error

:preflight_error
echo Preflight check failed. Environment is not ready for installer build.
pause
exit /b 1

:error
echo Build installer failed. Please check the errors above.
pause
exit /b 1
