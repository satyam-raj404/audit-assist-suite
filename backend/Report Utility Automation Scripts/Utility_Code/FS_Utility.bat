@echo off
REM Change to the directory where this batch file is located
cd /d "%~dp0"

REM Open the current directory in File Explorer
start .

REM Run the Python script using relative path
python FS_Utility.py

pause
