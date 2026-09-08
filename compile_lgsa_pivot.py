#!/usr/bin/env python3
"""
Script para compilar dados de poços com formato pivotado.

Extrai informações de células específicas e gera um arquivo de saída
com uma linha por amostra e múltiplas colunas Size + dados adicionais.

Estrutura:
  Well | MD | Amostra | Size[1] | Size[2] | ... | Size[38] | A75 | A76 | ... | A85
"""

import os
import sys
from pathlib import Path
from openpyxl import load_workbook
import csv

try:
    import xlrd
    HAS_XLRD = True
except ImportError:
    HAS_XLRD = False


def abbreviate_grain_name(name):
    """
    Abrevia nomes de grãos mantendo primeira letra de cada palavra (exceto última).
    Adiciona 'g' após o adjetivo para colunas Sand.
    Exemplo: "Very Fine Sand" -> "Vfg Sand"
    """
    if not name:
        return ''

    words = name.split()
    if len(words) <= 1:
        return name

    abbreviated = []
    for i, word in enumerate(words[:-1]):
        if i == 0:
            abbreviated.append(word[0].upper())
        else:
            abbreviated.append(word[0].lower())

    # Concatena as letras e adiciona 'g' se for Sand
    abbrev_str = ''.join(abbreviated)
    if words[-1].lower() == 'sand':
        abbrev_str += 'g'

    abbreviated_final = [abbrev_str, words[-1]]
    return ' '.join(abbreviated_final)


def extract_well_data(file_path):
    """
    Extrai dados de um arquivo Excel (.xlsx ou .xls).

    Args:
        file_path (str): Caminho do arquivo Excel

    Returns:
        dict: Dicionário com os dados extraídos de uma amostra
    """
    try:
        file_ext = Path(file_path).suffix.lower()

        if file_ext == '.xls':
            # Lê arquivo .xls usando xlrd
            if not HAS_XLRD:
                print(f"Erro: xlrd não está instalado. Use: pip install xlrd --user", file=sys.stderr)
                return None

            import xlrd
            wb = xlrd.open_workbook(file_path)
            ws = wb.sheet_by_index(0)

            # Extrai Well, Amostra e MD (xlrd usa 0-indexed)
            well = ws.cell_value(2, 0)    # A3
            amostra = ws.cell_value(2, 14)  # O3
            md = ws.cell_value(3, 14)    # O4

            # Extrai Size (F71:F35 em ordem decrescente)
            size_values = []
            for row_idx in range(70, 33, -1):  # F71:F35 (0-indexed: 70 até 34)
                try:
                    size_val = ws.cell_value(row_idx, 5)  # Coluna F
                    size_values.append(float(size_val) if size_val else None)
                except (ValueError, TypeError):
                    size_values.append(None)

            # Extrai nomes das colunas (A75:A85)
            column_names = []
            for row_idx in range(74, 85):  # A75:A85 (0-indexed: 74-84)
                try:
                    col_name = ws.cell_value(row_idx, 0)  # Coluna A
                    column_names.append(str(col_name) if col_name else '')
                except (ValueError, TypeError):
                    column_names.append('')

            # Extrai valores (F75:F85)
            values = []
            for row_idx in range(74, 85):  # F75:F85 (0-indexed: 74-84)
                try:
                    val = ws.cell_value(row_idx, 5)  # Coluna F
                    values.append(float(val) if val else None)
                except (ValueError, TypeError):
                    values.append(None)

        else:
            # Lê arquivo .xlsx usando openpyxl
            wb = load_workbook(file_path, data_only=True)
            ws = wb.active

            # Extrai Well, Amostra e MD
            well = ws['A3'].value
            amostra = ws['O3'].value
            md = ws['O4'].value

            # Extrai Size (F71:F35 em ordem decrescente)
            size_values = []
            for row in range(71, 34, -1):  # F71:F35
                try:
                    size_val = ws[f'F{row}'].value
                    size_values.append(float(size_val) if size_val else None)
                except (ValueError, TypeError):
                    size_values.append(None)

            # Extrai nomes das colunas (A75:A85)
            column_names = []
            for row in range(75, 86):  # A75:A85
                try:
                    col_name = ws[f'A{row}'].value
                    column_names.append(str(col_name) if col_name else '')
                except (ValueError, TypeError):
                    column_names.append('')

            # Extrai valores (F75:F85)
            values = []
            for row in range(75, 86):  # F75:F85
                try:
                    val = ws[f'F{row}'].value
                    values.append(float(val) if val else None)
                except (ValueError, TypeError):
                    values.append(None)

        # Debug: Mostra o que foi extraído
        print(f"    Well: {well}, MD: {md}, Amostra: {amostra}")
        print(f"    Size values ({len(size_values)}): {size_values[:3]}...")
        print(f"    Column names ({len(column_names)}): {column_names[:3]}...")
        print(f"    Values ({len(values)}): {values[:3]}...")

        return {
            'Well': well,
            'MD': md,
            'Amostra': amostra,
            'Size': size_values,
            'ColumnNames': column_names,
            'Values': values
        }

    except Exception as e:
        print(f"Erro ao processar {file_path}: {e}", file=sys.stderr)
        return None


def compile_well_data(input_folder, output_file=None):
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
    well_name = None
    column_names = None

    print(f"Processando {len(excel_files)} arquivo(s)...")
    for excel_file in excel_files:
        print(f"  Processando: {excel_file.name}")
        well_data = extract_well_data(str(excel_file))
        if well_data:
            compiled_data.append(well_data)
            # Extrai o nome do poço da primeira entrada
            if well_name is None and well_data.get('Well'):
                well_name = well_data.get('Well')
            # Extrai nomes das colunas da primeira entrada
            if column_names is None and well_data.get('ColumnNames'):
                column_names = well_data.get('ColumnNames')

    if not compiled_data:
        print("Nenhum dado foi extraído.", file=sys.stderr)
        return False

    # Gera o nome do arquivo se não foi fornecido
    if output_file is None:
        if well_name:
            output_file = f"{well_name}_LGSA_Pivot.csv"
        else:
            output_file = 'compiled_wells_pivot.csv'
            print("Aviso: Não foi possível determinar o nome do poço. Usando nome padrão.", file=sys.stderr)

    # Escreve no arquivo de saída (na mesma pasta dos arquivos de entrada)
    output_path = input_path / output_file

    try:
        with open(output_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter='\t')

            # Monta header com nomes das colunas
            header = ['Well', 'MD', 'Amostra', 'Size']
            # Adiciona nomes das colunas de A75:A85 (abreviados)
            abbreviated_names = []
            if column_names:
                for name in column_names:
                    abbreviated_names.append(abbreviate_grain_name(name))
                header.extend(abbreviated_names)
            writer.writerow(header)

            # Monta linha de unidades
            units_row = ['', 'm', '', 'mm']
            # Adiciona "%" para colunas de dados
            for _ in abbreviated_names:
                units_row.append('%')
            writer.writerow(units_row)

            # Dados compilados (múltiplas linhas por amostra)
            for data in compiled_data:
                size_vals = data.get('Size', [])
                values = data.get('Values', [])

                # Uma linha para cada valor de Size
                for i, size_val in enumerate(size_vals):
                    row = [
                        data.get('Well', ''),
                        data.get('MD', ''),
                        data.get('Amostra', ''),
                        size_val if size_val is not None else ''
                    ]

                    # Adiciona valores de F75:F85
                    for val in values:
                        row.append(val if val is not None else '')

                    writer.writerow(row)

        print(f"\n✓ Dados compilados com sucesso em: {output_path.absolute()}")
        return True

    except Exception as e:
        print(f"Erro ao escrever arquivo de saída: {e}", file=sys.stderr)
        return False


if __name__ == '__main__':
    print("=" * 60)
    print("Script de Compilação de Dados de Poços (Pivot)")
    print("=" * 60)

    # Se argumentos foram passados, usa modo linha de comando
    if len(sys.argv) >= 2:
        input_folder = sys.argv[1]
        output_file = sys.argv[2] if len(sys.argv) > 2 else None
    else:
        # Modo interativo
        print("\nDigite o caminho da pasta contendo os arquivos Excel:")
        print("(Exemplo: ./dados ou C:\\Users\\seu_usuario\\dados)")
        input_folder = input("\nCaminho da pasta: ").strip()

        if not input_folder:
            print("\nErro: Caminho não pode estar vazio!")
            sys.exit(1)

        output_file = None
        print("\n" + "=" * 60)

    success = compile_well_data(input_folder, output_file)

    if success:
        print("\n" + "=" * 60)
        print("✓ Processamento concluído com sucesso!")
        print("=" * 60)

    sys.exit(0 if success else 1)
