@echo off
REM Batch script to run maintenance_alert.py with proper working directory
REM Use this file for Windows Task Scheduler to avoid path issues

REM Get the directory where this batch file is located
cd /d "%~dp0"

REM Log the execution start
echo [%date% %time%] Starting maintenance_alert.py from directory: %cd%

REM Run the Python script
python "%~dp0maintenance_alert.py"

REM Log the exit code
echo [%date% %time%] Script finished with exit code: %errorlevel%

REM Pause for debugging (comment out for production)
REM pause