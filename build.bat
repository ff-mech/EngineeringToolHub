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

:: ── Copy all companion files next to the .exe ───────────────────────
echo  [4/5]  Copying companion files...

:: es.exe (Bom Filler stock search)
if not "!ES_EXE!"=="" (
    copy /y "!ES_EXE!" "dist\Engineering Tool Hub\es.exe" >nul 2>&1
    echo         Copied es.exe
)

:: User manual (How to Use tab)
if exist "Engineering_Tool_Hub.pdf" (
    copy /y "Engineering_Tool_Hub.pdf" "dist\Engineering Tool Hub\Engineering_Tool_Hub.pdf" >nul 2>&1
    echo         Copied Engineering_Tool_Hub.pdf
) else (
    echo  [WARN]  Engineering_Tool_Hub.pdf not found - How to Use tab will show an error.
)

:: Engineering Design Package PDFs (Training Materials tab)
if exist "tools\EngineeringDesignPackage" (
    xcopy /s /i /y "tools\EngineeringDesignPackage" "dist\Engineering Tool Hub\tools\EngineeringDesignPackage\" >nul 2>&1
    echo         Copied tools\EngineeringDesignPackage\
) else (
    echo  [WARN]  tools\EngineeringDesignPackage not found - Training Materials print will show an error.
)

:: FoxFab Design Tips (Training Materials tab)
if exist "tools\FoxFab_Design_Tips.docx" (
    if not exist "dist\Engineering Tool Hub\tools" mkdir "dist\Engineering Tool Hub\tools"
    copy /y "tools\FoxFab_Design_Tips.docx" "dist\Engineering Tool Hub\tools\FoxFab_Design_Tips.docx" >nul 2>&1
    echo         Copied tools\FoxFab_Design_Tips.docx
) else (
    echo  [WARN]  tools\FoxFab_Design_Tips.docx not found - Training Materials button will show an error.
)

:: File Logger script (File Logger tab)
if exist "tools\File Logger" (
    xcopy /s /i /y "tools\File Logger" "dist\Engineering Tool Hub\tools\File Logger\" >nul 2>&1
    echo         Copied tools\File Logger\
) else (
    echo  [WARN]  tools\File Logger not found - File Logger tab will show an error.
)

:: SolidWorks macro (SW Batch Update tab - Open Folder button)
if exist "tools\SolidworksBatchUpdate" (
    xcopy /s /i /y "tools\SolidworksBatchUpdate" "dist\Engineering Tool Hub\tools\SolidworksBatchUpdate\" >nul 2>&1
    echo         Copied tools\SolidworksBatchUpdate\
) else (
    echo  [WARN]  tools\SolidworksBatchUpdate not found - SW Batch Update folder link will not work.
)

echo         Done.

echo.
echo  [5/5]  Build complete.
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
