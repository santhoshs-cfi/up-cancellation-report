@echo off
:: ============================================================
::  UrbanPiper Report Automation - Windows Task Scheduler Setup
::  Run this file as Administrator (right-click → Run as administrator)
:: ============================================================

echo.
echo ============================================================
echo  UrbanPiper Report Automation - Scheduler Setup
echo ============================================================
echo.

:: ── Detect Python path ───────────────────────────────────────
for /f "delims=" %%i in ('where python') do set PYTHON_EXE=%%i
if "%PYTHON_EXE%"=="" (
    echo ERROR: Python not found in PATH. Please install Python first.
    pause
    exit /b 1
)
echo Found Python at: %PYTHON_EXE%

:: ── Get the folder where this .bat file lives ────────────────
set SCRIPT_DIR=%~dp0
set SCRIPT_PATH=%SCRIPT_DIR%urbanpiper_report.py

echo Script folder: %SCRIPT_DIR%
echo Script path:   %SCRIPT_PATH%
echo.

:: ── Install required packages ────────────────────────────────
echo Installing required Python packages...
"%PYTHON_EXE%" -m pip install playwright requests --quiet
"%PYTHON_EXE%" -m playwright install chromium
echo.

:: ── Delete old tasks if they exist ───────────────────────────
schtasks /delete /tn "UrbanPiper_DailyReport"   /f >nul 2>&1
schtasks /delete /tn "UrbanPiper_WeeklyReport"  /f >nul 2>&1
schtasks /delete /tn "UrbanPiper_MonthlyReport" /f >nul 2>&1

:: ============================================================
::  TASK 1: Daily at 7:00 AM — Yesterday's data
:: ============================================================
echo Creating Task 1: Daily report at 7:00 AM...
schtasks /create ^
    /tn "UrbanPiper_DailyReport" ^
    /tr "\"%PYTHON_EXE%\" \"%SCRIPT_PATH%\" daily" ^
    /sc DAILY ^
    /st 07:00 ^
    /ru "%USERNAME%" ^
    /rl HIGHEST ^
    /f

if %errorlevel%==0 (
    echo   [OK] Daily report task created.
) else (
    echo   [FAIL] Daily report task creation failed.
)

:: ============================================================
::  TASK 2: Every Monday at 8:00 AM — Last 7 days (prev Mon-Sun)
:: ============================================================
echo Creating Task 2: Weekly report every Monday at 8:00 AM...
schtasks /create ^
    /tn "UrbanPiper_WeeklyReport" ^
    /tr "\"%PYTHON_EXE%\" \"%SCRIPT_PATH%\" weekly" ^
    /sc WEEKLY ^
    /d MON ^
    /st 08:00 ^
    /ru "%USERNAME%" ^
    /rl HIGHEST ^
    /f

if %errorlevel%==0 (
    echo   [OK] Weekly report task created.
) else (
    echo   [FAIL] Weekly report task creation failed.
)

:: ============================================================
::  TASK 3: Every 1st of month at 10:00 AM — Previous full month
::  (Uses monthly scheduler - schtasks /sc MONTHLY /d 1)
:: ============================================================
echo Creating Task 3: Monthly report on 1st of each month at 10:00 AM...
schtasks /create ^
    /tn "UrbanPiper_MonthlyReport" ^
    /tr "\"%PYTHON_EXE%\" \"%SCRIPT_PATH%\" monthly" ^
    /sc MONTHLY ^
    /d 1 ^
    /st 10:00 ^
    /ru "%USERNAME%" ^
    /rl HIGHEST ^
    /f

if %errorlevel%==0 (
    echo   [OK] Monthly report task created.
) else (
    echo   [FAIL] Monthly report task creation failed.
)

:: ── Summary ──────────────────────────────────────────────────
echo.
echo ============================================================
echo  Setup complete! Scheduled tasks:
echo.
echo  [Daily]   Every day at 7:00 AM  - Yesterday's data
echo  [Weekly]  Every Monday at 8:00 AM  - Last 7 days
echo  [Monthly] Every 1st of month at 10:00 AM  - Prev full month
echo ============================================================
echo.
echo You can view/edit tasks in:
echo   Task Scheduler → Task Scheduler Library
echo   Search for "UrbanPiper"
echo.
echo To test immediately, run:
echo   python "%SCRIPT_PATH%" test
echo.
pause
