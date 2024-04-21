@echo off

echo Attempting to set PowerShell execution policy to allow script execution. Administrator privileges required...
PowerShell -Command "Start-Process PowerShell -ArgumentList '-ExecutionPolicy Bypass -Command Set-ExecutionPolicy RemoteSigned -Scope CurrentUser' -Verb RunAs"

echo Creating virtual environment...
python -m venv venv

echo Activating virtual environment...
call venv\Scripts\activate

echo Installing dependencies from requirements.txt...
python -m pip install -r requirements.txt

echo Running main.py...
python main.py

echo Press any key to exit.
pause > nul
