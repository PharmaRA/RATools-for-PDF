@echo off
setlocal

set "ROOT_DIR=%~dp0"
set "DIST_DIR=%ROOT_DIR%dist"
set "MAIN_FILE=%ROOT_DIR%main.py"
set "NO_UPDATE_MAIN_FILE=%ROOT_DIR%main_no_update.py"
set "ICON_FILE=%ROOT_DIR%icon.ico"
set "PLUGINS_DIR=%ROOT_DIR%plugins"

if not exist "%MAIN_FILE%" (
    echo [ERROR] Cannot find main.py in %ROOT_DIR%
    exit /b 1
)

if not exist "%NO_UPDATE_MAIN_FILE%" (
    echo [ERROR] Cannot find main_no_update.py in %ROOT_DIR%
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

if not exist "%DIST_DIR%" mkdir "%DIST_DIR%"

for /f "delims=" %%A in ('python -c "from ratools_pdf.config.version import APP_VERSION_STR; print(APP_VERSION_STR)"') do set "APP_VERSION_STR=%%A"
for /f "delims=" %%A in ('python -c "from ratools_pdf.config.version import APP_COMPANY; print(APP_COMPANY)"') do set "APP_COMPANY=%%A"
for /f "delims=" %%A in ('python -c "from ratools_pdf.config.version import APP_NAME; print(APP_NAME)"') do set "APP_NAME=%%A"

if not defined APP_COMPANY (
    echo [ERROR] Failed to read APP_COMPANY from ratools_pdf.config.version.
    exit /b 1
)
if not defined APP_NAME (
    echo [ERROR] Failed to read APP_NAME from ratools_pdf.config.version.
    exit /b 1
)
if not defined APP_VERSION_STR (
    echo [ERROR] Failed to read APP_VERSION_STR from ratools_pdf.config.version.
    exit /b 1
)

set "PYTHON3_DLL="
for /f "delims=" %%A in ('python -c "import sys, pathlib; p = pathlib.Path(sys.base_prefix) / 'python3.dll'; print(p if p.is_file() else '')"') do set "PYTHON3_DLL=%%A"

set "PYMUPDF_DIR="
for /f "delims=" %%A in ('python -c "import pymupdf, pathlib; print(pathlib.Path(pymupdf.__file__).parent)"') do set "PYMUPDF_DIR=%%A"

set "BUILD_TARGET=%~1"
if "%BUILD_TARGET%"=="" set "BUILD_TARGET=all"

if /i "%BUILD_TARGET%"=="main" goto :do_main_only
if /i "%BUILD_TARGET%"=="no_update" goto :do_no_update_only
if /i "%BUILD_TARGET%"=="all" goto :do_all

echo [ERROR] Unknown build target '%BUILD_TARGET%'. Use 'all', 'main', or 'no_update'.
exit /b 1

:do_main_only
call :build_variant "%MAIN_FILE%" "RATools-for-PDF" ""
if errorlevel 1 exit /b 1
goto :build_finish

:do_no_update_only
call :build_variant "%NO_UPDATE_MAIN_FILE%" "RATools-for-PDF-NoUpdate" "--nofollow-import-to=ratools_pdf.services.update_checker"
if errorlevel 1 exit /b 1
goto :build_finish

:do_all
call :build_variant "%MAIN_FILE%" "RATools-for-PDF" ""
if errorlevel 1 exit /b 1

call :build_variant "%NO_UPDATE_MAIN_FILE%" "RATools-for-PDF-NoUpdate" "--nofollow-import-to=ratools_pdf.services.update_checker"
if errorlevel 1 exit /b 1
goto :build_finish

:build_variant
set "ENTRY_FILE=%~1"
set "EXE_NAME=%~2"
set "VARIANT_EXTRA_ARGS=%~3"
set "ENTRY_BASENAME=%~n1"
set "VERSIONED_DIR_NAME=%EXE_NAME%_v%APP_VERSION_STR%"
set "RAW_OUTPUT_DIR=%DIST_DIR%\%ENTRY_BASENAME%.dist"
set "VERSIONED_OUTPUT_DIR=%DIST_DIR%\%VERSIONED_DIR_NAME%"

if exist "%RAW_OUTPUT_DIR%" rmdir /s /q "%RAW_OUTPUT_DIR%"
if exist "%VERSIONED_OUTPUT_DIR%" rmdir /s /q "%VERSIONED_OUTPUT_DIR%"

echo.
echo [INFO] Building %VERSIONED_DIR_NAME% with Nuitka standalone...
python -m nuitka "%ENTRY_FILE%" ^
  --standalone ^
  --assume-yes-for-downloads ^
  --enable-plugin=pyside6 ^
  --windows-console-mode=disable ^
  --windows-icon-from-ico="%ICON_FILE%" ^
  --include-module=fitz ^
  --include-module=ratools_pdf.config.paths ^
  --include-module=statistics ^
  --no-deployment-flag=excluded-module-usage ^
  --nofollow-import-to=pymupdf ^
  --nofollow-import-to=pandas ^
  --nofollow-import-to=numpy ^
  --nofollow-import-to=scipy ^
  --nofollow-import-to=matplotlib ^
  --nofollow-import-to=torch ^
  --nofollow-import-to=torchvision ^
  --nofollow-import-to=tensorflow ^
  --nofollow-import-to=easyocr ^
  --nofollow-import-to=IPython ^
  --nofollow-import-to=jupyter ^
  --nofollow-import-to=flask ^
  --nofollow-import-to=sqlalchemy ^
  --nofollow-import-to=openai ^
  --nofollow-import-to=pygame ^
  --nofollow-import-to=wxauto ^
  --nofollow-import-to=uiautomation ^
  --nofollow-import-to=snownlp ^
  --nofollow-import-to=apscheduler ^
  --nofollow-import-to=dateparser ^
  --nofollow-import-to=comtypes ^
  %VARIANT_EXTRA_ARGS% ^
  --include-data-files="%ROOT_DIR%LICENSE=LICENSE" ^
  --include-data-files="%ROOT_DIR%THIRD_PARTY_NOTICES.md=THIRD_PARTY_NOTICES.md" ^
  --include-data-files="%ICON_FILE%=icon.ico" ^
  --include-raw-dir="%PLUGINS_DIR%=plugins" ^
  --lto=no ^
  --jobs=8 ^
  --output-dir="%DIST_DIR%" ^
  --output-filename="%EXE_NAME%.exe" ^
  --company-name="%APP_COMPANY%" ^
  --product-name="%APP_NAME%" ^
  --file-description="%APP_NAME%" ^
  --file-version="%APP_VERSION_STR%" ^
  --product-version="%APP_VERSION_STR%"

if errorlevel 1 (
    echo [ERROR] Nuitka build failed for %EXE_NAME%.
    exit /b 1
)

if not exist "%RAW_OUTPUT_DIR%" (
    echo [ERROR] Nuitka raw output directory does not exist: %RAW_OUTPUT_DIR%
    exit /b 1
)

if not exist "%RAW_OUTPUT_DIR%\plugins\qpdf\qpdf.exe" (
    echo [INFO] Copying plugins to %RAW_OUTPUT_DIR%\plugins...
    xcopy /E /I /Y /Q "%PLUGINS_DIR%" "%RAW_OUTPUT_DIR%\plugins" >nul
    if errorlevel 1 (
        echo [ERROR] Failed to copy plugins directory.
        exit /b 1
    )
)

if defined PYMUPDF_DIR (
    if exist "%PYMUPDF_DIR%" (
        echo [INFO] Copying PyMuPDF runtime package to %RAW_OUTPUT_DIR%\pymupdf...
        xcopy /E /I /Y /Q "%PYMUPDF_DIR%" "%RAW_OUTPUT_DIR%\pymupdf" >nul
        if errorlevel 1 (
            echo [ERROR] Failed to copy PyMuPDF runtime package.
            exit /b 1
        )
    )
)

if defined PYTHON3_DLL (
    if exist "%PYTHON3_DLL%" (
        echo [INFO] Copying python3.dll to output directory...
        copy /Y "%PYTHON3_DLL%" "%RAW_OUTPUT_DIR%\" >nul
    )
)

set "OPENSSL_PLUGIN=%RAW_OUTPUT_DIR%\PySide6\qt-plugins\tls\qopensslbackend.dll"
if exist "%OPENSSL_PLUGIN%" (
    echo [INFO] Removing Qt OpenSSL TLS plugin for %EXE_NAME% to avoid startup DLL conflicts...
    del /q "%OPENSSL_PLUGIN%"
)

echo [INFO] Renaming %RAW_OUTPUT_DIR% to %VERSIONED_DIR_NAME%...
ren "%RAW_OUTPUT_DIR%" "%VERSIONED_DIR_NAME%"
if errorlevel 1 (
    timeout /t 1 /nobreak >nul
    ren "%RAW_OUTPUT_DIR%" "%VERSIONED_DIR_NAME%"
)
if errorlevel 1 (
    echo [ERROR] Failed to rename output folder to %VERSIONED_DIR_NAME%.
    exit /b 1
)

set "TARGET_EXE=%VERSIONED_OUTPUT_DIR%\%EXE_NAME%.exe"
if not exist "%TARGET_EXE%" (
    echo [ERROR] Target executable not found: %TARGET_EXE%
    exit /b 1
)

echo [OK] %EXE_NAME% build finished: %TARGET_EXE%
exit /b 0

:build_finish
echo.
if exist "%DIST_DIR%\main.build" rmdir /s /q "%DIST_DIR%\main.build"
if exist "%DIST_DIR%\main_no_update.build" rmdir /s /q "%DIST_DIR%\main_no_update.build"
echo [OK] Build completed.
if exist "%DIST_DIR%\RATools-for-PDF_v%APP_VERSION_STR%\RATools-for-PDF.exe" (
    echo [OK] Update-enabled output folder: %DIST_DIR%\RATools-for-PDF_v%APP_VERSION_STR%
)
if exist "%DIST_DIR%\RATools-for-PDF-NoUpdate_v%APP_VERSION_STR%\RATools-for-PDF-NoUpdate.exe" (
    echo [OK] No-update output folder: %DIST_DIR%\RATools-for-PDF-NoUpdate_v%APP_VERSION_STR%
)
endlocal
exit /b 0
