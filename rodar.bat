@echo off
echo Iniciando Extrator CIP...
streamlit run "%~dp0app.py" --server.headless false
pause
