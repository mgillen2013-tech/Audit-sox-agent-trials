@echo off
rem Double-click to start the CY Testing Agent. Runs on this computer only
rem (localhost) -- nobody else can reach it. Close this window to stop it.
rem
rem Written with "goto" instead of ( ... ) blocks on purpose: Windows reads a
rem ")" in the folder name -- e.g. a re-download unzipped as "...vir6t2 (1)" --
rem as the end of a block, and the window then closes before showing anything.
cd /d "%~dp0"

rem It must sit in the unzipped project folder. Opened from inside the ZIP
rem (Windows lets you browse into one), it is copied alone to a temp folder
rem and cannot find the app.
if not exist "cy_testing_app.py" goto :not_unzipped

where python >nul 2>nul
if errorlevel 1 goto :no_python

rem Install what the app needs: on the first run in a folder, and again
rem whenever an update adds something (".installed" keeps a copy of the
rem list that was last installed).
fc /b requirements.txt ".installed" >nul 2>nul
if not errorlevel 1 goto :start
echo Installing what the app needs - this can take a few minutes...
python -m pip install -r requirements.txt
if errorlevel 1 goto :install_failed
copy /y requirements.txt ".installed" >nul

:start
echo Starting the CY Testing Agent - your browser will open in a moment.
echo Keep this window open while you use it. Close it to stop the app.
python -m streamlit run cy_testing_app.py --server.address localhost --browser.gatherUsageStats false
echo.
echo The app has stopped. If that was unexpected, the reason is in the messages above.
pause
exit /b 0

:not_unzipped
echo This launcher cannot find the app next to it.
echo.
echo It must be run from the UNZIPPED project folder:
echo   1. Right-click the downloaded ZIP and choose "Extract All...", then Extract.
echo   2. Open the extracted folder - it should contain cy_testing_app.py,
echo      requirements.txt and an "agent" folder.
echo   3. Double-click "Start CY Testing Agent.bat" in that folder.
echo.
echo This window is running from:
cd
pause
exit /b 1

:no_python
echo Python was not found. Install Python 3.11+ from python.org and tick
echo "Add python.exe to PATH" during installation, then try again.
pause
exit /b 1

:install_failed
echo Installing failed - see the messages above.
pause
exit /b 1
