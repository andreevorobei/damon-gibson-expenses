@echo off
title Expense Reconciliation - Streamlit App
color 0B

echo.
echo ==========================================
echo    Expense Reconciliation Web App
echo        Powered by Streamlit
echo ==========================================
echo.

REM Проверяем Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ Python not found! Please install Python first.
    echo Download from: https://python.org
    pause
    exit /b
)

REM Проверяем Streamlit
python -c "import streamlit" >nul 2>&1
if %errorlevel% neq 0 (
    echo 📦 Installing Streamlit and dependencies...
    pip install -r streamlit_requirements.txt
    if %errorlevel% neq 0 (
        echo ❌ Failed to install packages!
        pause
        exit /b
    )
)

echo 🚀 Starting Streamlit application...
echo.
echo 🌐 The app will open in your default browser
echo 📍 URL: http://localhost:8501
echo.
echo 💡 Instructions:
echo 1. Upload your CapitalOne file (bank statement)
echo 2. Upload your Jobber file (expenses)  
echo 3. Adjust settings in sidebar if needed
echo 4. Click "Run Reconciliation"
echo 5. Download Excel report
echo.
echo ⏹️ To stop the app: Press Ctrl+C in this window
echo.

REM Запускаем Streamlit
streamlit run streamlit_app.py --server.headless false --browser.gatherUsageStats false

echo.
echo 👋 App closed. Press any key to exit...
pause >nul
