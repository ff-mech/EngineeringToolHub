@echo off
setlocal EnableDelayedExpansion
pushd "%~dp0"

:: ================================================================
::  Engineering Tool Hub  —  build.bat
::  Builds Engineering Tool Hub.exe using PyInstaller (--onedir)
::  Fast launch: no extraction needed, all files stay in the dist folder.
:: ================================================================

echo.
echo  ============================================================
echo   Engineering Tool Hub  —  Build Script
echo  ============================================================
echo.

:: Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python not found. Make sure Python is on your PATH.
    popd
    pause
    exit /b 1
)

:: Install / upgrade dependencies
echo  [1/4]  Installing dependencies...
python -m pip install --quiet --upgrade pyinstaller xlwings pypdf pywin32 pymupdf openpyxl
if errorlevel 1 (
    echo  [ERROR] pip install failed. Check your internet connection or proxy.
    pause
    exit /b 1
)
echo         Done.

:: Clean previous build
echo  [2/4]  Cleaning previous build...
if exist "dist\Engineering Tool Hub" (
    rmdir /s /q "dist\Engineering Tool Hub"
)
if exist "build" rmdir /s /q "build"
echo         Done.

:: Locate es.exe to bundle it
set "ES_EXE="
if exist "tools\BomFiller\es.exe"  set "ES_EXE=tools\BomFiller\es.exe"
if exist "es.exe"                  set "ES_EXE=es.exe"

if "!ES_EXE!"=="" (
    echo  [WARN]  es.exe not found — Bom Filler stock search will rely on system PATH.
    echo          Expected location: tools\BomFiller\es.exe
    set "ADD_ES="
) else (
    echo  [INFO]  Bundling es.exe from: !ES_EXE!
    set "ADD_ES=--add-data "!ES_EXE!;.""
)

:: Build
echo  [3/4]  Running PyInstaller...
echo.

python -m PyInstaller ^
    --noconfirm ^
    --onedir ^
    --windowed ^
    --name "Engineering Tool Hub" ^
    --icon NONE ^
    --distpath "dist" ^
    --workpath "build" ^
    --hidden-import "xlwings" ^
    --hidden-import "win32com.client" ^
    --hidden-import "win32print" ^
    --hidden-import "pythoncom" ^
    --hidden-import "pywintypes" ^
    --hidden-import "pypdf" ^
    --hidden-import "pypdf._reader" ^
    --hidden-import "pypdf._writer" ^
    --hidden-import "tkinter" ^
    --hidden-import "tkinter.ttk" ^
    --hidden-import "tkinter.filedialog" ^
    --hidden-import "tkinter.messagebox" ^
    --hidden-import "fitz" ^
    --hidden-import "openpyxl" ^
    !ADD_ES! ^
    app.py

if errorlevel 1 (
    echo.
    echo  [ERROR] PyInstaller build failed. See output above.
    pause
    exit /b 1
)

:: Copy es.exe into dist folder (next to the exe) as a fallback
if not "!ES_EXE!"=="" (
    copy /y "!ES_EXE!" "dist\Engineering Tool Hub\es.exe" >nul 2>&1
)

echo.
echo  [4/4]  Build complete.
echo.
echo  ============================================================
echo   Output:  dist\Engineering Tool Hub\Engineering Tool Hub.exe
echo.
echo   To distribute: copy the entire
echo     dist\Engineering Tool Hub\
echo   folder to the target machine. Do NOT move just the .exe.
echo  ============================================================
echo.
pause
