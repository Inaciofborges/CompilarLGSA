#!/usr/bin/env python3
"""
Script para compilar dados de poços a partir de arquivos Excel.

Extrai informações de células específicas e gera um arquivo de saída
com a estrutura: Well | MD (m) | Amostra | Size (mm) | Volume (%)
"""

import os
import sys
from pathlib import Path
from openpyxl import load_workbook
import csv
from statistics import mean


def extract_well_data(file_path):
    """
    Extrai dados de um arquivo Excel.

    Args:
        file_path (str): Caminho do arquivo Excel

    Returns:
        dict: Dicionário com os dados extraídos
    """
    try:
        wb = load_workbook(file_path, data_only=True)
        ws = wb.active

        data = {}

        # Extrai valores das células específicas
        data['Well'] = ws['A3'].value
        data['MD'] = ws['O4'].value
        data['Amostra'] = ws['O3'].value

        # Extrai ranges e calcula a média
        size_values = []
        for row in range(75, 86):  # A75 a A85
            cell_value = ws[f'A{row}'].value
            if cell_value is not None:
                try:
                    size_values.append(float(cell_value))
                except (ValueError, TypeError):
                    pass

        data['Size'] = mean(size_values) if size_values else None

        # Extrai volumes
        volume_values = []
        for row in range(75, 86):  # F75 a F85
            cell_value = ws[f'F{row}'].value
            if cell_value is not None:
                try:
                    volume_values.append(float(cell_value))
                except (ValueError, TypeError):
                    pass

        data['Volume'] = mean(volume_values) if volume_values else None

        return data

    except Exception as e:
        print(f"Erro ao processar {file_path}: {e}", file=sys.stderr)
        return None


def compile_well_data(input_folder, output_file='compiled_wells.csv'):
    """
    Processa todos os arquivos Excel de uma pasta e compila os dados.

    Args:
        input_folder (str): Pasta contendo os arquivos Excel
        output_file (str): Arquivo de saída (CSV)
    """
    input_path = Path(input_folder)

    if not input_path.exists():
        print(f"Erro: A pasta '{input_folder}' não existe.", file=sys.stderr)
        return False

    # Encontra todos os arquivos Excel
    excel_files = list(input_path.glob('*.xlsx')) + list(input_path.glob('*.xls'))

    if not excel_files:
        print(f"Aviso: Nenhum arquivo Excel encontrado em '{input_folder}'", file=sys.stderr)
        return False

    compiled_data = []

    print(f"Processando {len(excel_files)} arquivo(s)...")
    for excel_file in excel_files:
        print(f"  Processando: {excel_file.name}")
        well_data = extract_well_data(str(excel_file))
        if well_data:
            compiled_data.append(well_data)

    if not compiled_data:
        print("Nenhum dado foi extraído.", file=sys.stderr)
        return False

    # Escreve no arquivo de saída
    output_path = Path(output_file)

    try:
        with open(output_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter='|')

            # Cabeçalho com nomes das colunas
            writer.writerow(['Well', 'MD (m)', 'Amostra', 'Size (mm)', 'Volume (%)'])

            # Segunda linha com unidades (já incluídas no cabeçalho)
            # Se quiser uma segunda linha vazia ou apenas com unidades, descomente:
            # writer.writerow(['', 'm', '', 'mm', '%'])

            # Dados compilados
            for data in compiled_data:
                writer.writerow([
                    data.get('Well', ''),
                    data.get('MD', ''),
                    data.get('Amostra', ''),
                    f"{data.get('Size', ''):.2f}" if data.get('Size') else '',
                    f"{data.get('Volume', ''):.2f}" if data.get('Volume') else ''
                ])

        print(f"\n✓ Dados compilados com sucesso em: {output_path.absolute()}")
        return True

    except Exception as e:
        print(f"Erro ao escrever arquivo de saída: {e}", file=sys.stderr)
        return False


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Uso: python compile_well_data.py <pasta_entrada> [arquivo_saida.csv]")
        print("\nExemplo:")
        print("  python compile_well_data.py ./dados compiled_wells.csv")
        sys.exit(1)

    input_folder = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else 'compiled_wells.csv'

    success = compile_well_data(input_folder, output_file)
    sys.exit(0 if success else 1)
