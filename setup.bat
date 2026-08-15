@echo off
REM Creates a virtual environment (.venv), installs requirements.txt into it,
REM and launches the app. Double-click this file, or run it from a terminal
REM inside the app\ folder.
cd /d "%~dp0"

if not exist ".venv" (
    echo Creating virtual environment in .venv ...
    python -m venv .venv
)

echo Installing requirements ...
.venv\Scripts\pip install --upgrade pip
.venv\Scripts\pip install -r requirements.txt

echo Starting Streamlit ...
.venv\Scripts\streamlit run app.py
