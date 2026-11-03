@echo off
set "PYTHON=%~dp0.venv\Scripts\python.exe"
if not exist "%PYTHON%" (
    echo Python environment not found: "%PYTHON%"
    echo Create it and install dependencies with:
    echo   py -m venv "%~dp0.venv"
    echo   "%~dp0.venv\Scripts\python.exe" -m pip install pandas openpyxl
    pause
    exit /b 1
)
if "%~1"=="" (
    echo Drag one or more .urj files onto this batch file.
    pause
    exit /b 1
)
"%PYTHON%" "%~dp0convert_urj.py" %*
pause
pause