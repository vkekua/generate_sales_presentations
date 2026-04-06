@echo off
echo Setting up Sales Presentation Generator...
echo.

python -m venv venv
call venv\Scripts\activate
pip install -r requirements.txt

echo.
echo Generating SSL certificate...
python generate_cert.py

echo.
echo Setup complete! Run "start.bat" to start the server.
pause
