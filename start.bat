@echo off
title Harley's Offer Letter Generator
echo ========================================================
echo     Harley's Automatic Offer Letter Generation Portal
echo ========================================================
echo.
cd /d "%~dp0"

if not exist ".venv\Scripts\activate.bat" (
    echo [INFO] Virtual environment not found. Setting up...
    python -m venv .venv
    call .venv\Scripts\activate.bat
    pip install -r requirements.txt
) else (
    call .venv\Scripts\activate.bat
)

echo [INFO] Launching HR Offer Letter Generator Web App...
echo.
streamlit run app.py
pause
