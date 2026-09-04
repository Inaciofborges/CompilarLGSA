# Script de Compilação de Dados de Poços (Well Data Compiler)

Este script processa arquivos Excel de poços e compila as informações em um arquivo estruturado.

## Requisitos

- Python 3.7+
- Biblioteca `openpyxl` para leitura de arquivos Excel

### Instalação das dependências

```bash
pip install openpyxl
```

## Estrutura esperada dos arquivos de entrada

Os arquivos Excel devem conter os dados nas seguintes células:

| Campo | Célula |
|-------|--------|
| Well | A3 |
| MD (Measured Depth) | O4 |
| Amostra | O3 |
| Size | A75 a A85 (10 valores individuais) |
| Volume | F75 a F85 (10 valores individuais) |

## Uso

### Sintaxe básica

```bash
python compile_well_data.py <pasta_entrada> [arquivo_saida.csv]
```

### Exemplos

**Exemplo 1:** Processar arquivos da pasta `./dados` e gerar `compiled_wells.csv`

```bash
python compile_well_data.py ./dados
```

**Exemplo 2:** Especificar nome do arquivo de saída

```bash
python compile_well_data.py ./dados resultado_compilado.csv
```

## Formato de saída

O arquivo de saída será um CSV com o seguinte formato:

```
Well|MD (m)|Amostra|Size (mm)|Volume (%)
|m||mm|%
WELL-001|1500|Sample-A1|2.10|65.50
WELL-001|1500|Sample-A1|2.30|67.20
WELL-001|1500|Sample-A1|2.20|66.80
...
WELL-001|1500|Sample-A1|2.30|68.00
WELL-002|1650.5|Sample-B1|2.10|65.50
...
```

**Estrutura:**
- **Linha 1**: Nomes das colunas com unidades (Well | MD (m) | Amostra | Size (mm) | Volume (%))
- **Linha 2**: Apenas as unidades (| m | | mm | %)
- **Linhas 3+**: Dados dos poços (cada poço terá 10 linhas de dados)

**Unidades:**
- **MD**: metros (m)
- **Size**: milímetros (mm)
- **Volume**: percentual (%)

## Funcionalidades

✓ Lê múltiplos arquivos Excel (.xlsx, .xls)
✓ Extrai dados de células específicas
✓ Mantém os 10 valores individuais de cada range (não agrega)
✓ Gera arquivo CSV com estrutura: coluna de variáveis + linha de unidades
✓ Cada poço gera 10 linhas de dados no resultado
✓ Tratamento de erros robusto
✓ Mensagens informativas

## Personalização

Se desejar modificar o script para:
- Alterar as células de extração ou ranges
- Mudar o formato de saída (Excel, JSON)
- Adicionar mais colunas ou unidades
- Mudar o delimitador de arquivo (vírgula, ponto-e-vírgula, etc.)

Edite o arquivo `compile_well_data.py` nas seções indicadas.

## Troubleshooting

**Erro: "Nenhum arquivo Excel encontrado"**
- Verifique se os arquivos têm extensão .xlsx ou .xls
- Confirme se estão na pasta especificada

**Erro: "Erro ao processar [arquivo]"**
- Verifique se o arquivo não está corrompido
- Confirme que as células especificadas existem (A3, O3, O4)
- Confirme que os ranges A75:A85 e F75:F85 existem e contêm dados

## Notas

- Os 10 valores de Size (A75:A85) e Volume (F75:F85) são mantidos **individualmente**
- Cada poço gera exatamente **10 linhas** no arquivo de saída
- Valores vazios ou não-numéricos são deixados em branco na saída
- A segunda linha sempre contém as unidades, facilitando importação em ferramentas de análise
