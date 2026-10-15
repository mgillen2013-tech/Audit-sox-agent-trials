@echo off
rem Double-click to start the CY Testing Agent. Runs on this computer only
rem (localhost) -- nobody else can reach it. Close this window to stop it.
cd /d "%~dp0"

rem It must sit in the unzipped project folder. Opened from inside the ZIP
rem (Windows lets you browse into one), it is copied alone to a temp folder
rem and cannot find the app.
if not exist "cy_testing_app.py" (
    echo This launcher cannot find the app next to it.
    echo.
    echo It must be run from the UNZIPPED project folder:
    echo   1. Right-click the downloaded ZIP and choose "Extract All...", then Extract.
    echo   2. Open the extracted folder - it should contain cy_testing_app.py,
    echo      requirements.txt and an "agent" folder.
    echo   3. Double-click "Start CY Testing Agent.bat" in that folder.
    echo.
    echo This window is running from: %CD%
    pause
    exit /b 1
)

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
