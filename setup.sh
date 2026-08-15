#!/usr/bin/env bash
# Creates a virtual environment (.venv), installs requirements.txt into it,
# and launches the app. Run from inside the app/ folder:
#   chmod +x setup.sh && ./setup.sh
set -e
cd "$(dirname "$0")"

if [ ! -d ".venv" ]; then
    echo "Creating virtual environment in .venv ..."
    python3 -m venv .venv
fi

echo "Installing requirements ..."
.venv/bin/pip install --upgrade pip
.venv/bin/pip install -r requirements.txt

echo "Starting Streamlit ..."
.venv/bin/streamlit run app.py
