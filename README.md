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
| Size | A75 a A85 (será calculada a média) |
| Volume | F75 a F85 (será calculada a média) |

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
WELL-001|1500|Sample-A1|2.45|65.30
WELL-002|1650|Sample-B1|2.38|68.75
```

As unidades estão indicadas no cabeçalho:
- **MD**: metros (m)
- **Size**: milímetros (mm)
- **Volume**: percentual (%)

## Funcionalidades

✓ Lê múltiplos arquivos Excel (.xlsx, .xls)
✓ Extrai dados de células específicas
✓ Calcula a média dos valores em ranges (Size e Volume)
✓ Gera arquivo CSV estruturado
✓ Tratamento de erros robusto
✓ Mensagens informativas

## Personalizações

Se desejar modificar o script para:
- Alterar as células de extração
- Usar um método diferente de agregação (median, sum, etc.)
- Mudar o formato de saída (Excel, JSON)
- Adicionar mais campos

Edite o arquivo `compile_well_data.py` nas seções indicadas.

## Troubleshooting

**Erro: "Nenhum arquivo Excel encontrado"**
- Verifique se os arquivos têm extensão .xlsx ou .xls
- Confirme se estão na pasta especificada

**Erro: "Erro ao processar [arquivo]"**
- Verifique se o arquivo não está corrompido
- Confirme que as células especificadas existem
- Verifique se os dados são numéricos nos ranges (Size e Volume)

## Notas

- Os ranges de Size (A75:A85) e Volume (F75:F85) terão sua **média** calculada
- Valores vazios ou não-numéricos serão ignorados na média
- Se não houver valores válidos, a célula no resultado ficará vazia
