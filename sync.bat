@echo off
setlocal

set SOURCE=H:\ShellFolders\Documents\GitHub\ASToolDEV\AS_ToolsDev.extension
set DEST=I:\002_BIM\007_Addins\Aukett Swanke\Extensions\ASToolDEV\AS_ToolsDev.extension
set LOGFILE=%~dp0sync_log.txt

echo ============================================================
echo  ASToolDEV Extension Sync
echo  %DATE% %TIME%
echo ============================================================
echo.

:: ── verify source exists ──────────────────────────────────────
if not exist "%SOURCE%" (
    echo ERROR: Source folder not found:
    echo   %SOURCE%
    echo Check that the H: drive is mapped.
    pause
    exit /b 1
)

:: ── verify destination is reachable ──────────────────────────
if not exist "%DEST%" (
    echo WARNING: Destination folder not found:
    echo   %DEST%
    echo Attempting to create it...
    mkdir "%DEST%"
    if errorlevel 1 (
        echo ERROR: Could not create destination. Check that the I: drive is mapped.
        pause
        exit /b 1
    )
)

:: ── pull latest from GitHub first ────────────────────────────
echo Pulling latest from GitHub...
cd /d H:\ShellFolders\Documents\GitHub\ASToolDEV
git pull
if errorlevel 1 (
    echo WARNING: git pull failed. Continuing with local files.
)

echo.

:: ── sync to shared drive ──────────────────────────────────────
echo Syncing to shared drive...
echo   From: %SOURCE%
echo   To:   %DEST%
echo.

robocopy "%SOURCE%" "%DEST%" ^
    /MIR ^
    /XD __pycache__ .git .vscode .idea ^
    /XF *.pyc *.pyo Thumbs.db .DS_Store *.log ^
    /NP ^
    /NDL ^
    /LOG+:"%LOGFILE%"

:: robocopy exit codes 0-7 are success (8+ are errors)
if %ERRORLEVEL% GEQ 8 (
    echo.
    echo ERROR: Sync failed. Check the log: %LOGFILE%
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo Done! Extension synced to I: drive.
echo Log saved to: %LOGFILE%
echo.
pause
endlocal
