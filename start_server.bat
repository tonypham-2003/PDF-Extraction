@echo off
title DP World — PDF to Excel Server
cd /d "%~dp0"
echo.
echo  ==========================================
echo   DP World  ^|  PDF to Excel Converter
echo  ==========================================
echo   Keep this window open while using the app
echo   Close this window to stop the server
echo  ==========================================
echo.
:start
python app.py
echo.
echo  [Server stopped — restarting in 3 seconds...]
echo  [Close this window to exit completely]
timeout /t 3 /nobreak >nul
goto start
