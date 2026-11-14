@echo off
title Ung Dung Bao Gia Tu Dong
cd /d "%~dp0"
if exist api_key.txt set /p ANTHROPIC_API_KEY=<api_key.txt
py -c "import anthropic, PIL, reportlab" 2>nul || py -m pip install anthropic pillow reportlab
echo Dang khoi chay ung dung, vui long cho trong giay lat...
py -m streamlit run app.py
pause