@echo off
title PPTX Translator Web UI
echo ======================================================
echo    KHOI DONG GIAO DIEN PPTX TRANSLATOR (GEMINI AI)
echo ======================================================
echo.
cd /d "%~dp0"
call .\venv\Scripts\activate
streamlit run app.py --browser.gatherUsageStats false
pause
