@echo off
cd /d "%~dp0"
if not exist "venv\Scripts\python.exe" (
  echo Creating virtual environment...
  python -m venv venv
  call venv\Scripts\activate.bat
  pip install -r requirements.txt
) else (
  call venv\Scripts\activate.bat
)
echo.
echo Starting Cold Email app at http://127.0.0.1:5000
echo Keep Outlook open.
echo.
python app.py
pause
