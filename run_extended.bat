@echo off
REM Script de atalho para executar o compilador extended de dados de poços no Windows

REM Primeiro tenta instalar as dependências se não estiverem disponíveis
python -m pip install openpyxl xlrd --user 2>nul

REM Executa o script
python compile_lgsa_extended.py

pause
