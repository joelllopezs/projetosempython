@echo off
cd /d "%~dp0"
python -m pip install -r requirements.txt -q
python app.py
pause
