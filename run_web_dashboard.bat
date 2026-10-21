@echo off
chcp 65001 > nul
echo ==============================================================================
echo   HE THONG QUAN LY TRI THUC & RA DE THI TIN HOC GDPT 2018 (NEO4J WEB DASHBOARD)
echo ==============================================================================
echo Dang khoi dong Web Dashboard tren trinh duyet...
python -m streamlit run web_dashboard.py
pause
