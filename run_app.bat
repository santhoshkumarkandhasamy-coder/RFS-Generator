@echo off
cd /d "%~dp0"
python -m pip install -r requirements.txt
rem browser used to read NSE / BSE pages (downloads once)
python -m playwright install chromium
rem fixed port: if an older copy of the app is still running, this stops with "port 8501 is already in use"
rem instead of quietly starting on 8502 - close the old window (or end python.exe in Task Manager) and run again
python -m streamlit run app.py --server.port 8501
pause
