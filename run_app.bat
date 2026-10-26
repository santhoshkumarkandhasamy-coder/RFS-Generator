@echo off
cd /d "%~dp0"
python -m pip install -r requirements.txt
rem browser used to read NSE / BSE pages (downloads once)
python -m playwright install chromium
python -m streamlit run app.py
pause
