@echo off
:: VCCS VtccpApp — Update from GitHub and rebuild
:: Drop this file anywhere convenient and double-click to sync + build.
:: The repo root is one level above this file (vtccp\tools\update-and-build.bat).

setlocal
set "REPO_ROOT=%~dp0..\.."
set "SRC_ROOT=%~dp0.."

echo.
echo ================================================
echo  VCCS VtccpApp — Update ^& Build
echo ================================================
echo.

:: Pull latest from GitHub
echo [1/2] Pulling latest from GitHub...
cd /d "%REPO_ROOT%"
git pull
if errorlevel 1 (
    echo ERROR: git pull failed. Check your network and credentials.
    pause
    exit /b 1
)
echo.

:: Build
echo [2/2] Building VtccpWindows.sln (Release)...
cd /d "%SRC_ROOT%"
dotnet build VtccpWindows.sln -c Release
if errorlevel 1 (
    echo ERROR: Build failed. See errors above.
    pause
    exit /b 1
)

echo.
echo ================================================
echo  Done. VtccpApp is up to date and built.
echo ================================================
echo.
pause
