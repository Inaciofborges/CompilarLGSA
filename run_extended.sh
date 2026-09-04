#!/bin/bash
# Script de atalho para executar o compilador extended de dados de poços

# Primeiro tenta instalar as dependências se não estiverem disponíveis
python3 -m pip install openpyxl xlrd --user 2>/dev/null

# Executa o script
python3 compile_lgsa_extended.py
