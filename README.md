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
| Size | A75 a A85 (10 valores individuais) |
| Volume | F75 a F85 (10 valores individuais) |

## Uso

### Modo Interativo (Recomendado)

#### No Linux/macOS:
```bash
./run.sh
```

#### No Windows:
Duplo-clique em `run.bat`

O script pedirá:
1. **Caminho da pasta** com os arquivos Excel
2. **Nome do arquivo de saída** (opcional, padrão: `compiled_wells.csv`)

Exemplo de entrada interativa:
```
Digite o caminho da pasta contendo os arquivos Excel:
./dados

Digite o nome do arquivo de saída (padrão: compiled_wells.csv):
resultado_compilado.csv
```

### Modo Linha de Comando

#### Sintaxe básica

```bash
python compile_well_data.py <pasta_entrada> [arquivo_saida.csv]
```

#### Exemplos

**Exemplo 1:** Processar arquivos da pasta `./dados` com nome padrão

```bash
python compile_well_data.py ./dados
```

**Exemplo 2:** Especificar nome do arquivo de saída

```bash
python compile_well_data.py ./dados resultado_compilado.csv
```

## Formato de saída

O arquivo de saída será um arquivo TAB-delimited com o seguinte formato:

```
Well	MD	Size	Volume
		mm	%
WELL-003	2431.05	2	0
WELL-003	2431.05	1.681793	4.050356
WELL-003	2431.05	0.840896	12.658058
WELL-003	2431.05	0.420448	13.395099
...
WELL-003	2431.05	0.003285	4.903475
WELL-004	2431.05	2	0
WELL-004	2431.05	1.681793	4.050356
...
```

**Estrutura:**
- **Linha 1**: Nomes das colunas (Well, MD, Size, Volume) - sem unidades
- **Linha 2**: Unidades (vazio, vazio, mm, %)
- **Linhas 3+**: Dados dos poços (cada poço terá 10 linhas de dados)

**Unidades:**
- **MD**: metros (m) - sem unidade no cabeçalho
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
