@echo off

set PYTHON_EXE=C:\Users\"Edoardo Menon"\AppData\Local\Programs\Python\Python312\python.exe

echo Creating virtual environment...
%PYTHON_EXE% -m venv venv

echo Activating virtual environment...
call venv\Scripts\activate

echo Installing dependencies from requirements.txt...
%PYTHON_EXE% -m pip install -r requirements.txt

echo Running main.py...
%PYTHON_EXE% main.py

echo Press any key to exit.
pause > nul
