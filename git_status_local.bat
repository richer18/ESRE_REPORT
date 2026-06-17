@echo off
setlocal

where git >nul 2>nul
if errorlevel 1 (
    echo ERROR: Git is not installed or not available in PATH.
    pause
    exit /b 1
)

if not exist ".git" (
    echo This folder is not initialized as its own Git repository yet.
    echo Run 00_setup_github_repo.bat first.
    pause
    exit /b 1
)

git status --short --branch
echo.
git remote -v
pause

