@echo off
setlocal
python -B "%~dp0scripts\build_pyinstaller_safely.py" --install-dependency %*
exit /b %errorlevel%
