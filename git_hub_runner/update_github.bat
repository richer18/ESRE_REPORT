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

set "MSG=%~1"
if "%MSG%"=="" (
    set /p MSG=Commit message [Update ESRE report project]: 
)
if "%MSG%"=="" set "MSG=Update ESRE report project"

echo.
echo Current changes:
git status --short
echo.

echo Staging project files...
git add -A
if errorlevel 1 (
    echo ERROR: git add failed.
    pause
    exit /b 1
)

git diff --cached --quiet
if errorlevel 1 (
    echo Creating commit...
    git commit -m "%MSG%"
    if errorlevel 1 (
        echo ERROR: git commit failed.
        pause
        exit /b 1
    )
) else (
    echo No staged changes to commit.
)

echo.
echo Pushing to GitHub main branch...
git push -u origin main
if errorlevel 1 (
    echo ERROR: Push failed. You may need to sign in to GitHub or create the first remote branch.
    pause
    exit /b 1
)

echo.
echo GitHub is updated.
git status --short --branch
pause
