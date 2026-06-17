@echo off
setlocal
set "SCRIPT_DIR=%~dp0"

:menu
cls
echo ESRE_REPORT GitHub Helper
echo =========================
echo.
echo 1. Setup local repo and GitHub remote
echo 2. Update local files from GitHub
echo 3. Push local changes to GitHub
echo 4. Show local git status
echo 5. Exit
echo.
set /p CHOICE=Choose option: 

if "%CHOICE%"=="1" call "%SCRIPT_DIR%00_setup_github_repo.bat"
if "%CHOICE%"=="2" call "%SCRIPT_DIR%update_from_github.bat"
if "%CHOICE%"=="3" call "%SCRIPT_DIR%update_github.bat"
if "%CHOICE%"=="4" call "%SCRIPT_DIR%git_status_local.bat"
if "%CHOICE%"=="5" exit /b 0

goto menu
