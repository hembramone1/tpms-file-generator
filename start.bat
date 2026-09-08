@echo off
title TPMS File Generator (PWA)
echo ===================================================
echo           TPMS File Generator (PWA)
echo ===================================================
echo.
echo Starting local web server...
echo Serving at: http://localhost:8080
echo.
echo Launching your default browser...
start http://localhost:8080
echo.
echo Press Ctrl+C in this window to stop the server when done.
echo.
python -m http.server 8080
pause
