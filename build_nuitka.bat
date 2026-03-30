@echo off
setlocal

set "ROOT_DIR=%~dp0"
set "OUT_DIR=%ROOT_DIR%build"
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

python -m nuitka --version >nul 2>nul
if errorlevel 1 (
    echo [INFO] Nuitka is not installed. Installing build dependencies...
    python -m pip install -U nuitka ordered-set zstandard
    if errorlevel 1 (
        echo [ERROR] Failed to install Nuitka build dependencies.
        exit /b 1
    )
)

if not exist "%OUT_DIR%" mkdir "%OUT_DIR%"

echo [INFO] Building RATools for PDF with Nuitka...
python -m nuitka "%MAIN_FILE%" ^
  --standalone ^
  --assume-yes-for-downloads ^
  --enable-plugin=pyside6 ^
  --windows-console-mode=disable ^
  --include-module=fitz ^
  --include-module=app_paths ^
  --include-data-files="%ROOT_DIR%icon.png=icon.png" ^
  --include-data-dir="%ROOT_DIR%plugins=plugins" ^
  --output-dir="%OUT_DIR%" ^
  --company-name="RATools" ^
  --product-name="RATools for PDF" ^
  --file-description="RA PDF batch processing tool" ^
  --file-version="0.1.0"

if errorlevel 1 (
    echo [ERROR] Nuitka build failed.
    exit /b 1
)

echo.
echo [OK] Build completed.
echo [OK] Output folder: %OUT_DIR%\main.dist
echo [OK] Run: %OUT_DIR%\main.dist\main.exe
endlocal
