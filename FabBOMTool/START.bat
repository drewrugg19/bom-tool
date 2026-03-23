@echo off
title Fabrication BOM Tool
echo.
echo  ============================================
echo   Fabrication BOM Tool
echo  ============================================
echo.
echo  Installing/checking dependencies...
pip install -r requirements.txt --quiet
echo.
echo  Starting server...
echo  Open your browser to: http://localhost:5000
echo.
python app.py
pause
