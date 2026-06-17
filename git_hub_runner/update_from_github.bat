@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "REPO_ROOT=%SCRIPT_DIR%.."
cd /d "%REPO_ROOT%"

set "REPO_URL=https://github.com/richer18/ESRE_REPORT.git"

where git >nul 2>nul
if errorlevel 1 (
    echo ERROR: Git is not installed or not available in PATH.
    pause
    exit /b 1
)

if not exist ".git" (
    echo Local repo not found in this folder. Running setup first...
    call "%SCRIPT_DIR%00_setup_github_repo.bat"
    if errorlevel 1 exit /b 1
)

git remote get-url origin >nul 2>nul
if errorlevel 1 (
    git remote add origin "%REPO_URL%"
)

echo Fetching latest changes from GitHub...
git fetch origin
if errorlevel 1 (
    echo ERROR: Fetch failed. Check internet, GitHub access, or repository URL.
    pause
    exit /b 1
)

echo.
echo Updating local files from GitHub main branch...
git pull --rebase origin main
if errorlevel 1 (
    echo ERROR: Pull/rebase failed. Resolve conflicts, then run this again.
    pause
    exit /b 1
)

echo.
echo Local project is updated from GitHub.
git status --short --branch
pause
