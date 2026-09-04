#!/bin/bash
# Script de configuração - Instala as dependências necessárias

echo "============================================================"
echo "Instalando dependências do projeto..."
echo "============================================================"
echo ""

python3 -m pip install --upgrade pip
python3 -m pip install -r requirements.txt

echo ""
echo "============================================================"
echo "Instalação concluída!"
echo ""
echo "Agora você pode executar o script com:"
echo "  ./run.sh"
echo "============================================================"
echo ""
