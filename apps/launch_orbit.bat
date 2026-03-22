@echo off
cd /d "C:\Users\DELL\Desktop\Dashboard"
call venv_orbit\Scripts\activate
streamlit run apps\app_orbit.py
pause