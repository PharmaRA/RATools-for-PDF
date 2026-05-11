@echo off
setlocal

set "ROOT_DIR=%~dp0"
set "DIST_DIR=%ROOT_DIR%dist"
set "BUILD_DIR=%ROOT_DIR%build_pyinstaller"
set "SPEC_DIR=%ROOT_DIR%build_pyinstaller"
set "MAIN_FILE=%ROOT_DIR%main.py"
set "NO_UPDATE_MAIN_FILE=%ROOT_DIR%main_no_update.py"

if not exist "%MAIN_FILE%" (
    echo [ERROR] Cannot find main.py in %ROOT_DIR%
    exit /b 1
)

if not exist "%NO_UPDATE_MAIN_FILE%" (
    echo [ERROR] Cannot find main_no_update.py in %ROOT_DIR%
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
set "VERSION_INFO_FILE=%BUILD_DIR%\build_version_info_generated.txt"

for /f "delims=" %%A in ('python -c "from app_version import APP_COMPANY; print(APP_COMPANY)"') do set "APP_COMPANY=%%A"
for /f "delims=" %%A in ('python -c "from app_version import APP_NAME; print(APP_NAME)"') do set "APP_NAME=%%A"
for /f "delims=" %%A in ('python -c "from app_version import APP_VERSION_STR; print(APP_VERSION_STR)"') do set "APP_VERSION_STR=%%A"
for /f "delims=" %%A in ('python -c "from app_version import APP_WINDOWS_VERSION; print(*APP_WINDOWS_VERSION, sep=chr(44)+chr(32))"') do set "APP_WINDOWS_VERSION=%%A"

if not defined APP_COMPANY (
    echo [ERROR] Failed to read APP_COMPANY from app_version.py.
    exit /b 1
)
if not defined APP_NAME (
    echo [ERROR] Failed to read APP_NAME from app_version.py.
    exit /b 1
)
if not defined APP_VERSION_STR (
    echo [ERROR] Failed to read APP_VERSION_STR from app_version.py.
    exit /b 1
)
if not defined APP_WINDOWS_VERSION (
    echo [ERROR] Failed to read APP_WINDOWS_VERSION from app_version.py.
    exit /b 1
)

> "%VERSION_INFO_FILE%" echo VSVersionInfo(
>> "%VERSION_INFO_FILE%" echo   ffi=FixedFileInfo(
>> "%VERSION_INFO_FILE%" echo     filevers=(%APP_WINDOWS_VERSION%),
>> "%VERSION_INFO_FILE%" echo     prodvers=(%APP_WINDOWS_VERSION%),
>> "%VERSION_INFO_FILE%" echo     mask=0x3f,
>> "%VERSION_INFO_FILE%" echo     flags=0x0,
>> "%VERSION_INFO_FILE%" echo     OS=0x40004,
>> "%VERSION_INFO_FILE%" echo     fileType=0x1,
>> "%VERSION_INFO_FILE%" echo     subtype=0x0,
>> "%VERSION_INFO_FILE%" echo     date=(0, 0)
>> "%VERSION_INFO_FILE%" echo   ),
>> "%VERSION_INFO_FILE%" echo   kids=[
>> "%VERSION_INFO_FILE%" echo     StringFileInfo([
>> "%VERSION_INFO_FILE%" echo       StringTable(
>> "%VERSION_INFO_FILE%" echo         '080404B0',
>> "%VERSION_INFO_FILE%" echo         [
>> "%VERSION_INFO_FILE%" echo           StringStruct('CompanyName', '%APP_COMPANY%'),
>> "%VERSION_INFO_FILE%" echo           StringStruct('FileDescription', '%APP_NAME%'),
>> "%VERSION_INFO_FILE%" echo           StringStruct('FileVersion', '%APP_VERSION_STR%'),
>> "%VERSION_INFO_FILE%" echo           StringStruct('InternalName', 'RATools-for-PDF'),
>> "%VERSION_INFO_FILE%" echo           StringStruct('OriginalFilename', 'RATools-for-PDF.exe'),
>> "%VERSION_INFO_FILE%" echo           StringStruct('ProductName', '%APP_NAME%'),
>> "%VERSION_INFO_FILE%" echo           StringStruct('ProductVersion', '%APP_VERSION_STR%')
>> "%VERSION_INFO_FILE%" echo         ]
>> "%VERSION_INFO_FILE%" echo       )
>> "%VERSION_INFO_FILE%" echo     ]),
>> "%VERSION_INFO_FILE%" echo     VarFileInfo([VarStruct('Translation', [2052, 1200])])
>> "%VERSION_INFO_FILE%" echo   ]
>> "%VERSION_INFO_FILE%" echo )

call :build_variant "%MAIN_FILE%" "RATools-for-PDF" ""
if errorlevel 1 (
    call :cleanup_version_info
    exit /b 1
)

call :build_variant "%NO_UPDATE_MAIN_FILE%" "RATools-for-PDF-NoUpdate" "--exclude-module update_checker"
if errorlevel 1 (
    call :cleanup_version_info
    exit /b 1
)

call :cleanup_version_info

echo.
echo [OK] Build completed.
echo [OK] Update-enabled output folder: %DIST_DIR%\RATools-for-PDF
echo [OK] No-update output folder: %DIST_DIR%\RATools-for-PDF-NoUpdate
endlocal
exit /b 0

:build_variant
set "ENTRY_FILE=%~1"
set "EXE_NAME=%~2"
set "VARIANT_EXTRA_ARGS=%~3"

echo [INFO] Building %EXE_NAME% with PyInstaller onedir...
python -m PyInstaller "%ENTRY_FILE%" ^
  --noconfirm ^
  --clean ^
  --console ^
  --onedir ^
  --name "%EXE_NAME%" ^
  --distpath "%DIST_DIR%" ^
  --workpath "%BUILD_DIR%" ^
  --specpath "%SPEC_DIR%" ^
  --noupx ^
  --icon "%ROOT_DIR%icon.ico" ^
  --version-file "%VERSION_INFO_FILE%" ^
  --exclude-module torch ^
  --exclude-module torchvision ^
  --exclude-module easyocr ^
  --exclude-module tensorflow ^
  --exclude-module pandas ^
  --exclude-module numpy ^
  --exclude-module scipy ^
  --exclude-module PIL ^
  %VARIANT_EXTRA_ARGS% ^
  --add-data "%ROOT_DIR%LICENSE;." ^
  --add-data "%ROOT_DIR%THIRD_PARTY_NOTICES.md;." ^
  --add-data "%ROOT_DIR%icon.ico;." ^
  --add-data "%ROOT_DIR%plugins;plugins"

if errorlevel 1 (
    echo [ERROR] PyInstaller build failed for %EXE_NAME%.
    exit /b 1
)

set "OUTPUT_EXE=%DIST_DIR%\%EXE_NAME%\%EXE_NAME%.exe"
python "%ROOT_DIR%patch_pe_subsystem.py" "%OUTPUT_EXE%" --windows-gui
if errorlevel 1 (
    echo [ERROR] Failed to patch PE subsystem for %EXE_NAME%.
    exit /b 1
)

set "OPENSSL_PLUGIN=%DIST_DIR%\%EXE_NAME%\_internal\PySide6\plugins\tls\qopensslbackend.dll"
if exist "%OPENSSL_PLUGIN%" (
    echo [INFO] Removing Qt OpenSSL TLS plugin for %EXE_NAME% to avoid startup DLL conflicts...
    del /q "%OPENSSL_PLUGIN%"
)

echo [OK] %EXE_NAME% output folder: %DIST_DIR%\%EXE_NAME%
exit /b 0

:cleanup_version_info
if exist "%VERSION_INFO_FILE%" del /q "%VERSION_INFO_FILE%"
exit /b 0
