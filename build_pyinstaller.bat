@echo off
setlocal

set "ROOT_DIR=%~dp0"
set "DIST_DIR=%ROOT_DIR%dist"
set "BUILD_DIR=%ROOT_DIR%build_pyinstaller"
set "SPEC_DIR=%ROOT_DIR%build_pyinstaller"
set "MAIN_FILE=%ROOT_DIR%main.py"

if not exist "%MAIN_FILE%" (
    echo [ERROR] Cannot find main.py in %ROOT_DIR%
    exit /b 1
)

where python >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python is not available in PATH.
    exit /b 1
)

python -m PyInstaller --version >nul 2>nul
if errorlevel 1 (
    echo [INFO] PyInstaller is not installed. Installing build dependency...
    python -m pip install -U pyinstaller
    if errorlevel 1 (
        echo [ERROR] Failed to install PyInstaller.
        exit /b 1
    )
)

if not exist "%BUILD_DIR%" mkdir "%BUILD_DIR%"
if not exist "%SPEC_DIR%" mkdir "%SPEC_DIR%"

echo [INFO] Building RATools for PDF with PyInstaller onedir...
python -m PyInstaller "%MAIN_FILE%" ^
  --noconfirm ^
  --clean ^
  --console ^
  --onedir ^
  --name RATools-for-PDF ^
  --distpath "%DIST_DIR%" ^
  --workpath "%BUILD_DIR%" ^
  --specpath "%SPEC_DIR%" ^
  --noupx ^
  --icon "%ROOT_DIR%icon.png" ^
  --version-file "%ROOT_DIR%build_version_info.txt" ^
  --exclude-module torch ^
  --exclude-module torchvision ^
  --exclude-module easyocr ^
  --exclude-module tensorflow ^
  --exclude-module pandas ^
  --exclude-module numpy ^
  --exclude-module scipy ^
  --exclude-module PIL ^
  --add-data "%ROOT_DIR%LICENSE;." ^
  --add-data "%ROOT_DIR%THIRD_PARTY_NOTICES.md;." ^
  --add-data "%ROOT_DIR%icon.png;." ^
  --add-data "%ROOT_DIR%plugins;plugins"

if errorlevel 1 (
    echo [ERROR] PyInstaller build failed.
    exit /b 1
)

set "OUTPUT_EXE=%DIST_DIR%\RATools-for-PDF\RATools-for-PDF.exe"
python "%ROOT_DIR%patch_pe_subsystem.py" "%OUTPUT_EXE%" --windows-gui
if errorlevel 1 (
    echo [ERROR] Failed to patch PE subsystem.
    exit /b 1
)

set "OPENSSL_PLUGIN=%DIST_DIR%\RATools-for-PDF\_internal\PySide6\plugins\tls\qopensslbackend.dll"
if exist "%OPENSSL_PLUGIN%" (
    echo [INFO] Removing Qt OpenSSL TLS plugin to avoid startup DLL conflicts...
    del /q "%OPENSSL_PLUGIN%"
)

echo.
echo [OK] Build completed.
echo [OK] Output folder: %DIST_DIR%\RATools-for-PDF
echo [OK] Run: %OUTPUT_EXE%
endlocal
