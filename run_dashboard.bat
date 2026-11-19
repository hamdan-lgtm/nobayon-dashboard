@echo off
setlocal EnableExtensions
title Nobayon Traders - Business Performance Dashboard
cd /d "%~dp0"

rem  Double-click, or:
rem    run_dashboard.bat                 refresh the data, then open the dashboard
rem    run_dashboard.bat /norefresh      open with the data already on disk
rem    run_dashboard.bat /port 8510      force a port

set "PORT="
set "REFRESH=1"
:parse
if "%~1"=="" goto parsed
if /I "%~1"=="/norefresh" set "REFRESH=0"
if /I "%~1"=="/port" set "PORT=%~2"
shift
goto parse
:parsed

echo ============================================================
echo  Nobayon Traders - Monthly Business Performance
echo  %CD%
echo ============================================================
echo.

where python >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python was not found on PATH.
    echo         Install Python 3.12 and tick "Add python.exe to PATH", then run this file again.
    echo.
    pause
    exit /b 1
)

python -c "import streamlit, pandas, plotly, openpyxl" >nul 2>&1
if errorlevel 1 (
    echo [SETUP] Installing the required packages, this runs only once ...
    python -m pip install -r requirements.txt
    if errorlevel 1 (
        echo [ERROR] Package installation failed. Check your internet connection.
        pause
        exit /b 1
    )
)

rem Skip Streamlit's first-run "Welcome / Email" prompt.
if not exist "%USERPROFILE%\.streamlit" mkdir "%USERPROFILE%\.streamlit" >nul 2>&1
if not exist "%USERPROFILE%\.streamlit\credentials.toml" (
    > "%USERPROFILE%\.streamlit\credentials.toml" echo [general]
    >> "%USERPROFILE%\.streamlit\credentials.toml" echo email = ""
)
set "STREAMLIT_BROWSER_GATHER_USAGE_STATS=false"
set "PYTHONUTF8=1"

rem Reuse an already running dashboard instead of starting a twin.
if not defined PORT call :already_running
if defined ALREADY (exit /b 0)

if "%REFRESH%"=="1" (
    echo [1/2] Refreshing Power Query and validating the database ...
    python refresh_pipeline.py --no-git
    if errorlevel 1 (
        echo.
        echo [WARN] Validation reported failures. The dashboard opens with the data on disk -
        echo        review the Audit tab before sharing numbers.
        echo.
    )
) else (
    echo [1/2] Refresh skipped ^(/norefresh^) - using the data already on disk.
)

if defined PORT (
    set "PORTARG=--server.port %PORT%"
) else (
    set "PORTARG="
)

echo.
echo [2/2] Starting the dashboard ...
echo       Your browser opens automatically. Close this window to stop it.
echo.

python -m streamlit run streamlit_app.py %PORTARG% --server.headless false --browser.gatherUsageStats false

if errorlevel 1 (
    echo.
    echo [STOPPED] The dashboard exited with an error - read the messages above.
    echo.
    pause
) else (
    echo.
    echo [STOPPED] The dashboard has closed.
)
exit /b 0

rem ---------------------------------------------------------------------------
:already_running
set "HEALTH="
where curl >nul 2>&1
if errorlevel 1 exit /b 0
for /f "delims=" %%H in ('curl -s -m 2 http://localhost:8501/_stcore/health 2^>nul') do set "HEALTH=%%H"
if /I not "%HEALTH%"=="ok" exit /b 0
echo [INFO] The dashboard is already running on http://localhost:8501
echo        Opening it in your browser ...
start "" "http://localhost:8501"
set "ALREADY=1"
exit /b 0
