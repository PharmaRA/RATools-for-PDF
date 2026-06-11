@echo off
setlocal

set "ROOT_DIR=%~dp0"
set "OUT_DIR=%ROOT_DIR%build"
set "MAIN_FILE=%ROOT_DIR%main.py"
set "ICON_FILE=%ROOT_DIR%icon.ico"
set "PLUGINS_DIR=%ROOT_DIR%plugins"

if not exist "%MAIN_FILE%" (
    echo [ERROR] Cannot find main.py in %ROOT_DIR%
    exit /b 1
)

if not exist "%ICON_FILE%" (
    echo [ERROR] Cannot find icon.ico in %ROOT_DIR%
    exit /b 1
)

if not exist "%PLUGINS_DIR%" (
    echo [ERROR] Cannot find plugins directory in %ROOT_DIR%
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

for /f "delims=" %%A in ('python -c "from app_version import APP_VERSION_STR; print(APP_VERSION_STR)"') do set "APP_VERSION_STR=%%A"
for /f "delims=" %%A in ('python -c "from app_version import APP_COMPANY; print(APP_COMPANY)"') do set "APP_COMPANY=%%A"
for /f "delims=" %%A in ('python -c "from app_version import APP_NAME; print(APP_NAME)"') do set "APP_NAME=%%A"
set "VERSIONED_DIR_NAME=main_v%APP_VERSION_STR%.dist"
set "RAW_OUTPUT_DIR=%OUT_DIR%\main.dist"
set "VERSIONED_OUTPUT_DIR=%OUT_DIR%\%VERSIONED_DIR_NAME%"

echo [INFO] Building RATools for PDF with Nuitka...
python -m nuitka "%MAIN_FILE%" ^
  --standalone ^
  --assume-yes-for-downloads ^
  --enable-plugin=pyside6 ^
  --windows-console-mode=disable ^
  --include-module=fitz ^
  --include-module=app_paths ^
  --include-data-files="%ROOT_DIR%LICENSE=LICENSE" ^
  --include-data-files="%ROOT_DIR%THIRD_PARTY_NOTICES.md=THIRD_PARTY_NOTICES.md" ^
  --include-data-files="%ICON_FILE%=icon.ico" ^
  --include-data-dir="%PLUGINS_DIR%=plugins" ^
  --output-dir="%OUT_DIR%" ^
  --company-name="%APP_COMPANY%" ^
  --product-name="%APP_NAME%" ^
  --file-description="RA PDF batch processing tool" ^
  --file-version="%APP_VERSION_STR%"

if errorlevel 1 (
    echo [ERROR] Nuitka build failed.
    exit /b 1
)

if exist "%VERSIONED_OUTPUT_DIR%" rmdir /s /q "%VERSIONED_OUTPUT_DIR%"
ren "%RAW_OUTPUT_DIR%" "%VERSIONED_DIR_NAME%"
if errorlevel 1 (
    echo [ERROR] Failed to rename Nuitka output folder to %VERSIONED_DIR_NAME%.
    exit /b 1
)

echo.
echo [OK] Build completed.
echo [OK] Output folder: %VERSIONED_OUTPUT_DIR%
echo [OK] Run: %OUT_DIR%\main_v%APP_VERSION_STR%.dist\main.exe
endlocal
