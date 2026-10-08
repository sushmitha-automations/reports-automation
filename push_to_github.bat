@echo off
title Push Automation Code to GitHub
cd /d "%~dp0"
echo ===============================================================================
echo                 PUSHING CODE TO GITHUB REPOSITORY
echo                 https://github.com/sushmitha-automations/reports-automation
echo ===============================================================================
echo.
git push -u origin main
echo.
if %errorlevel% equ 0 (
    echo ===============================================================================
    echo [SUCCESS] Code pushed successfully to GitHub!
    echo.
    echo Next Step - Host for Free on Streamlit Cloud:
    echo 1. Open https://share.streamlit.io in your browser
    echo 2. Sign in with GitHub (sushmitha-automations)
    echo 3. Click "New app"
    echo 4. Repository: sushmitha-automations/reports-automation
    echo 5. Branch: main
    echo 6. Main file path: app.py
    echo 7. Click Deploy!
    echo ===============================================================================
) else (
    echo ===============================================================================
    echo [NOTE] If a browser window opened, please click "Authorize" to sign in to GitHub.
    echo Then run this file again to finish pushing!
    echo ===============================================================================
)
echo.
pause
