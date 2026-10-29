@echo off
:: ─────────────────────────────────────────────────────────────────────────────
::  build.bat  —  Build OutlookBulkMailer.exe with PyInstaller
::  Run this script from the project root directory on Windows.
:: ─────────────────────────────────────────────────────────────────────────────

echo [1/4] Installing dependencies...
pip install -r requirements.txt
if errorlevel 1 goto :error

echo [2/4] Cleaning previous build...
if exist build  rmdir /s /q build
if exist dist   rmdir /s /q dist

echo [3/4] Building executable...
pyinstaller ^
    --noconfirm ^
    --onefile ^
    --windowed ^
    --name "OutlookBulkMailer" ^
    --icon "assets\icon.ico" ^
    --add-data "modules;modules" ^
    --add-data "assets;assets" ^
    --hidden-import "msal" ^
    --hidden-import "msal.application" ^
    --hidden-import "msal.token_cache" ^
    --hidden-import "openpyxl" ^
    --hidden-import "requests" ^
    --hidden-import "PySide6.QtSvg" ^
    --hidden-import "PySide6.QtXml" ^
    app.py
if errorlevel 1 goto :error

echo [4/4] Done!
echo Executable: dist\OutlookBulkMailer.exe
pause
goto :eof

:error
echo BUILD FAILED. See errors above.
pause
exit /b 1
