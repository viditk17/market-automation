@echo off
echo ========================================
echo Market Automation - Local Server
echo ========================================
echo.
echo Step 1: Installing dependencies...
python -m pip install -r requirements.txt --quiet
echo [DONE] Dependencies ready
echo.
echo Step 2: Starting server...
echo.
python app4.py
pause
