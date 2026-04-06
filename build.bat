@echo off
echo Building Sales Presentation Generator .exe ...
echo.

call venv\Scripts\activate
pip install pyinstaller

pyinstaller --onefile --noconsole --name "SalesPresentationGenerator" ^
    --add-data "ppt_template.pptx;." ^
    --hidden-import openpyxl ^
    gui.py

echo.
echo Done! The .exe is in the "dist" folder.
echo Share "dist\SalesPresentationGenerator.exe" with the sales team.
pause
