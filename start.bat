@echo off
echo Starting Sales Presentation Generator...
echo.
cd /d "%~dp0"
call venv\Scripts\activate
start /min pythonw app.py
echo Server started in background.
echo You can close this window. The server will keep running.
echo.
echo To stop the server, run "stop.bat"
pause
