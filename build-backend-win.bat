@echo off
echo ============================================================
echo   Excel AI Analyzer -- Backend Compiler (Windows)
echo ============================================================
echo.

REM Check for Python using 'py' launcher (standard on Windows)
py --version >nul 2>&1
if errorlevel 1 (
    echo X Python not found.
    echo   Please install Python 3 from https://www.python.org
    echo   Make sure to check "Add Python to PATH" during installation.
    pause
    exit /b 1
)

echo Found Python:
py --version
echo.

REM Install PyInstaller
echo Installing PyInstaller...
py -m pip install pyinstaller --quiet
if errorlevel 1 (
    echo X Failed to install PyInstaller.
    pause
    exit /b 1
)

REM Install backend dependencies
echo Installing backend dependencies...
py -m pip install flask flask-cors pandas openpyxl xlrd numpy requests --quiet
if errorlevel 1 (
    echo X Failed to install dependencies.
    pause
    exit /b 1
)

echo.
echo Compiling backend to server.exe...
echo This will take 2-3 minutes...
echo.

REM Run PyInstaller
py -m PyInstaller backend\server.spec --distpath backend\dist --workpath backend\build --noconfirm

if errorlevel 1 (
    echo X Build failed.
    pause
    exit /b 1
)

echo.
echo ============================================================
if exist "backend\dist\server\server.exe" (
    echo   Build successful!
    for %%A in ("backend\dist\server\server.exe") do echo   Size: %%~zA bytes
) else (
    echo X server.exe not found after build.
)
echo ============================================================
pause
