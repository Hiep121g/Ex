@echo off
rem Start PyExcel. You can also drag an .xlsx / .pyxl / .csv file onto this file to open it.
cd /d "%~dp0"
start "" pythonw "%~dp0excel_app.py" %*
