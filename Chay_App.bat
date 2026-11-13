@echo off
title Ung Dung Bao Gia Tu Dong
cd /d "%~dp0"
echo Dang khoi chay ung dung, vui long cho trong giay lat...
py -m streamlit run app.py
pause