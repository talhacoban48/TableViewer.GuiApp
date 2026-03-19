@echo off
setlocal

echo ============================================================
echo  Table Viewer -- PyInstaller Build
echo ============================================================

:: Change to the directory where this script lives
cd /d "%~dp0"

:: ── Clean previous build ──────────────────────────────────────
echo [1/3] Cleaning previous build...
if exist build  rmdir /s /q build
if exist dist   rmdir /s /q dist
if exist TableViewer.spec del /q TableViewer.spec

:: ── Run PyInstaller ──────────────────────────────────────────
echo [2/3] Running PyInstaller...

pyinstaller ^
    --name "TableViewer" ^
    --icon "assets\favicon.ico" ^
    --windowed ^
    --onedir ^
    --add-data "assets;assets" ^
    --add-data "tableviewer;tableviewer" ^
    --hidden-import "openpyxl" ^
    --hidden-import "pandas" ^
    --hidden-import "PyQt5" ^
    --noconfirm ^
    main.py

if errorlevel 1 (
    echo.
    echo [ERROR] PyInstaller failed!
    pause
    exit /b 1
)

:: ── Done ─────────────────────────────────────────────────────
echo [3/3] Build complete.
echo.
echo Output folder: dist\TableViewer\
echo Executable   : dist\TableViewer\TableViewer.exe
echo.
echo You can now run Inno Setup on setup.iss to create the installer.
echo.
pause
