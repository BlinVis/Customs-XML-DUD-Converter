@echo off
REM Runs the app from source (needs Python 3.10+ installed).
python -m pip install -r requirements.txt >nul
python app.py
