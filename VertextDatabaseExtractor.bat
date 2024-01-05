@echo off
cls
color 0A

:: Activate Virtual Environment
call venv\Scripts\activate

:: Ensure that the virtual environment is activated before proceeding
if ERRORLEVEL 1 (
    echo Failed to activate virtual environment.
    exit /b 1
)

python main.py

PAUSE
