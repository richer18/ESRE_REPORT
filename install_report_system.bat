@echo off
setlocal EnableExtensions

set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

set "PYTHON_EXE="
set "REQUIREMENTS=%SCRIPT_DIR%requirements.txt"
set "REPORT_SCRIPT=%SCRIPT_DIR%run_collection_query.py"
set "FIREBIRD_CLIENT_64=C:\Program Files\Firebird\Firebird_2_5\bin\fbclient.dll"
set "ODBC_DSN=itaxzamboanguita"

echo ==================================================
echo   ESRE Report System Installer / Checker
echo ==================================================
echo.

where python >nul 2>nul
if not errorlevel 1 (
    for /f "delims=" %%P in ('where python') do (
        if not defined PYTHON_EXE set "PYTHON_EXE=%%P"
    )
)

if not defined PYTHON_EXE (
    echo STATUS: Python was not found in PATH.
    echo STATUS: Trying to install Python 3 using winget...
    where winget >nul 2>nul
    if errorlevel 1 (
        echo ERROR: winget is not available.
        echo Please install Python manually from https://www.python.org/downloads/
        echo Make sure to check "Add python.exe to PATH" during installation.
        pause
        exit /b 1
    )
    winget install --id Python.Python.3 --source winget --accept-package-agreements --accept-source-agreements
    if errorlevel 1 (
        echo ERROR: Python installation failed.
        pause
        exit /b 1
    )
    echo STATUS: Python install finished. Re-open PowerShell/CMD, then run this installer again.
    pause
    exit /b 0
)

echo STATUS: Python found:
echo %PYTHON_EXE%
python --version
echo.

if not exist "%REPORT_SCRIPT%" (
    echo ERROR: run_collection_query.py was not found.
    echo Expected: %REPORT_SCRIPT%
    pause
    exit /b 1
)

if not exist "%REQUIREMENTS%" (
    echo ERROR: requirements.txt was not found.
    echo Expected: %REQUIREMENTS%
    pause
    exit /b 1
)

echo STATUS: Upgrading pip...
python -m pip install --upgrade pip
if errorlevel 1 (
    echo WARNING: pip upgrade failed. Continuing to package install...
)

echo.
echo STATUS: Installing required Python packages...
python -m pip install -r "%REQUIREMENTS%"
if errorlevel 1 (
    echo ERROR: Python package installation failed.
    pause
    exit /b 1
)

echo.
echo STATUS: Checking Firebird 64-bit client DLL...
if exist "%FIREBIRD_CLIENT_64%" (
    echo OK: %FIREBIRD_CLIENT_64%
) else (
    echo WARNING: Firebird 64-bit client DLL was not found:
    echo %FIREBIRD_CLIENT_64%
    echo Install Firebird 2.5 64-bit client/tools, then run this installer again.
)

echo.
echo STATUS: Checking ODBC DSN name: %ODBC_DSN%
python -c "import pyodbc; dsns=[x for x in pyodbc.dataSources().keys()]; print('ODBC DSNs:', ', '.join(dsns)); raise SystemExit(0 if '%ODBC_DSN%' in dsns else 2)"
if errorlevel 2 (
    echo WARNING: ODBC DSN "%ODBC_DSN%" was not found for this Python/ODBC environment.
    echo Create it in ODBC Data Sources ^(64-bit^) with the same settings as the old app.
)

echo.
echo STATUS: Testing report database connection through ODBC...
python "%REPORT_SCRIPT%" --connection odbc --test-connection
if errorlevel 1 (
    echo.
    echo WARNING: ODBC connection test failed.
    echo Check DSN, username/password, Firebird client DLL, and main-server access.
    pause
    exit /b 1
)

echo.
echo SUCCESS: ESRE report system dependencies and ODBC connection are ready.
echo You can now run reports, for example:
echo python run_collection_query.py --connection odbc 25 2026-01-01 2026-01-31
echo.
pause
