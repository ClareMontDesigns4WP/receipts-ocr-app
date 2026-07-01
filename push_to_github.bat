@echo off
setlocal enabledelayedexpansion
title Purple Cow Accounting - Push to GitHub
color 0A

echo.
echo ============================================================
echo   Purple Cow Accounting - Receipts OCR Tool
echo   GitHub Repository Update
echo ============================================================
echo.

set REPO_DIR=C:\PurpleCow\receipts-ocr-app

cd /d "%REPO_DIR%"
if errorlevel 1 (
    echo ERROR: Repository folder not found at %REPO_DIR%
    pause
    exit /b 1
)

REM Safety check - confirm keygen.py is not present
if exist "keygen.py" (
    echo ERROR: keygen.py detected in this folder.
    echo This file must NEVER be pushed to GitHub.
    echo Please move it to C:\PurpleCow\private\ and try again.
    pause
    exit /b 1
)

echo [1/4] Checking Git status...
git status
echo.

echo [2/4] Staging files...
git add app.py
git add licence_check.py
git add .gitignore
REM Add any other files that belong in the repo below this line:
REM git add build.bat
REM git add installer.nsi
echo Files staged.
echo.

echo [3/4] Committing...
REM Get current date and time for commit message
for /f "tokens=1-3 delims=/ " %%a in ('date /t') do set TODAY=%%c-%%b-%%a
for /f "tokens=1 delims= " %%a in ('time /t') do set NOW=%%a
set COMMIT_MSG=Update %TODAY% %NOW%
git commit -m "%COMMIT_MSG%"
if errorlevel 1 (
    echo Nothing new to commit - files already up to date.
    pause
    exit /b 0
)
echo.

echo [4/4] Pushing to GitHub...
git push origin main
if errorlevel 1 (
    echo ERROR: Push failed.
    echo Check your internet connection and GitHub credentials.
    pause
    exit /b 1
)
echo.

echo ============================================================
echo   SUCCESS - Repository updated on GitHub
echo   Commit: %COMMIT_MSG%
echo ============================================================
echo.
echo You can now run build.bat to create a fresh installer.
echo.
pause