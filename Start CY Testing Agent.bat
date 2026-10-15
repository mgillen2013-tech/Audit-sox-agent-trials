@echo off
rem Double-click to start the CY Testing Agent. Runs on this computer only
rem (localhost) -- nobody else can reach it. Close this window to stop it.
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo Python was not found. Install Python 3.11+ from python.org and tick
    echo "Add python.exe to PATH" during installation, then try again.
    pause
    exit /b 1
)

rem First run in a newly unzipped folder: install what the app needs.
if not exist ".installed" (
    echo Installing what the app needs - first run only, this can take a few minutes...
    python -m pip install -r requirements.txt
    if errorlevel 1 (
        echo Installing failed - see the messages above.
        pause
        exit /b 1
    )
    echo done> ".installed"
)

echo Starting the CY Testing Agent - your browser will open in a moment.
echo Keep this window open while you use it. Close it to stop the app.
python -m streamlit run cy_testing_app.py --server.address localhost --browser.gatherUsageStats false
pause
