@echo off
setlocal

set "REPO_URL=https://github.com/richer18/ESRE_REPORT.git"
set "EXPECTED_FOLDER=ESRE_REPORT"

for %%I in ("%CD%") do set "CURRENT_FOLDER=%%~nxI"
if /I not "%CURRENT_FOLDER%"=="%EXPECTED_FOLDER%" (
    echo ERROR: Please run this file from the ESRE_REPORT project folder.
    echo Current folder: %CD%
    pause
    exit /b 1
)

where git >nul 2>nul
if errorlevel 1 (
    echo ERROR: Git is not installed or not available in PATH.
    pause
    exit /b 1
)

if not exist ".git" (
    echo Initializing local Git repository in this ESRE_REPORT folder...
    git init -b main
    if errorlevel 1 (
        git init
        git branch -M main
    )
) else (
    echo Local Git repository already exists in this folder.
)

git remote get-url origin >nul 2>nul
if errorlevel 1 (
    echo Adding GitHub remote origin...
    git remote add origin "%REPO_URL%"
) else (
    echo Updating GitHub remote origin...
    git remote set-url origin "%REPO_URL%"
)

echo.
echo Git setup complete.
echo Remote:
git remote -v
echo.
echo Current status:
git status --short --branch
echo.
pause

