# Script de Compilação de Dados de Poços (Well Data Compiler)

Este script processa arquivos Excel de poços e compila as informações em um arquivo estruturado.

## Requisitos

- Python 3.7+
- Biblioteca `openpyxl` para leitura de arquivos Excel (.xlsx)
- Biblioteca `xlrd` para leitura de arquivos Excel antigos (.xls)

## 🚀 Instalação Rápida

### Windows:
1. Duplo-clique em `setup.bat`
2. Aguarde a instalação das dependências
3. Duplo-clique em `run.bat` para usar o script

### Linux/macOS:
```bash
chmod +x setup.sh
./setup.sh
./run.sh
```

### Instalação Manual (opcional):
```bash
pip install -r requirements.txt
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

## Formatos Suportados

✅ **Arquivos Excel:**
- `.xlsx` (Excel 2010 e posterior) - via openpyxl
- `.xls` (Excel 2003 e anterior) - via xlrd

O script detecta automaticamente o formato e usa a biblioteca apropriada.

## Uso

### Modo Interativo (Recomendado)

#### No Linux/macOS:
```bash
./run.sh
```

#### No Windows:
Duplo-clique em `run.bat`

O script pedirá apenas:
- **Caminho da pasta** com os arquivos Excel

O arquivo de saída será nomeado automaticamente como: **{Nome_do_Poço}_LGSA.csv**

Exemplo de entrada interativa:
```
Digite o caminho da pasta contendo os arquivos Excel:
./dados
```

**Resultado:** Arquivo `WELL-003_LGSA.csv` gerado automaticamente!

### Modo Linha de Comando

#### Sintaxe básica

```bash
python compile_well_data.py <pasta_entrada> [arquivo_saida_personalizado.csv]
```

#### Exemplos

**Exemplo 1:** Processar arquivos da pasta `./dados` (nome gerado automaticamente)

```bash
python compile_well_data.py ./dados
```

Resultado: `WELL-003_LGSA.csv`

**Exemplo 2:** Especificar nome do arquivo de saída personalizado

```bash
python compile_well_data.py ./dados meu_resultado.csv
```

## Formato de saída

O arquivo de saída será um arquivo TAB-delimited com o seguinte formato:

```
Well	MD	Amostra	Size Class	Size	Volume
				mm	%
WELL-003	2431.05	Sample-A1	Granule	2	0
WELL-003	2431.05	Sample-A1	Very Coarse Sand	1.681793	4.050356
WELL-003	2431.05	Sample-A1	Coarse Sand	0.840896	12.658058
WELL-003	2431.05	Sample-A1	Medium Sand	0.420448	13.395099
...
WELL-003	2431.05	Sample-A1	Clay	0.003285	4.903475
WELL-004	2431.05	Sample-B1	Granule	2	0
WELL-004	2431.05	Sample-B1	Very Coarse Sand	1.681793	4.050356
...
```

**Estrutura:**
- **Linha 1**: Nomes das colunas (Well, MD, Amostra, Size Class, Size, Volume) - sem unidades
- **Linha 2**: Unidades (vazio, vazio, vazio, vazio, mm, %)
- **Linhas 3+**: Dados dos poços (cada poço terá 10 linhas de dados)

**Unidades:**
- **MD**: metros (m) - sem unidade no cabeçalho
- **Size**: milímetros (mm)
- **Volume**: percentual (%)

## Classificação de Tamanho (Escala de Wentworth)

A coluna "Size Class" classifica automaticamente o tamanho das partículas:

| Size Class | Intervalo (mm) |
|------------|----------------|
| Granule | ≥ 2.0 |
| Very Coarse Sand | 1.682 - 2.0 |
| Coarse Sand | 0.841 - 1.682 |
| Medium Sand | 0.42 - 0.841 |
| Fine Sand | 0.21 - 0.42 |
| Very Fine Sand | 0.105 - 0.21 |
| Coarse Silt | 0.053 - 0.105 |
| Medium Silt | 0.026 - 0.053 |
| Fine Silt | 0.013 - 0.026 |
| Very Fine Silt | 0.007 - 0.013 |
| Clay | < 0.007 |

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
