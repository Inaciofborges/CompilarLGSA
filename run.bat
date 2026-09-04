@echo off
REM Script de atalho para executar o compilador de dados de poços no Windows

REM Primeiro tenta instalar openpyxl se não estiver disponível
python -m pip install openpyxl --user 2>nul

REM Executa o script
python compile_well_data.py

pause
