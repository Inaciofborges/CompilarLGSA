# Script de Compilação de Dados Excel

## Descrição
Este script automatiza a compilação de dados de múltiplos arquivos Excel. Ele:
- Lê arquivos Excel de uma pasta especificada pelo usuário
- Extrai dados de cada aba (exceto "graph")
- Coleta valores de identificação (Well, MD, CPERM_A, CPOR, HgPOR) das células A9, J9, J10, J11, J12
- Extrai dados da linha 18 (A18:J18)
- Gera um arquivo consolidado com todos os dados

## Requisitos
- Python 3.6+
- pandas
- openpyxl

## Instalação de Dependências
```bash
pip install pandas openpyxl
```

## Como Usar

### Opção 1: Executar direto com Python
```bash
python compilar_dados.py
```

### Opção 2: Executar com Python 3
```bash
python3 compilar_dados.py
```

## Fluxo de Uso
1. Execute o script
2. Uma janela de seleção de pasta aparecerá
3. Selecione a pasta contendo seus arquivos Excel
4. O script processará todos os arquivos `.xlsx`, `.xls` e `.xlsm`
5. Um novo arquivo `dados_compilados.xlsx` será criado na mesma pasta

## Estrutura do Arquivo de Saída

O arquivo gerado terá a seguinte estrutura:

| Well | MD | CPERM_A | CPOR | HgPOR | A | B | C | D | E | F | G | H | I | J |
|------|-------|---------|-------|-------|---|---|---|---|---|---|---|---|---|---|
| | m | mD | v/v | v/v | | | | | | | | | | |
| (valor) | (valor) | (valor) | (valor) | (valor) | ... dados ... |
| ... mais linhas ... |

Onde:
- **Primeira linha**: Cabeçalhos das colunas
- **Segunda linha**: Unidades de medida
- **Linhas seguintes**: Dados compilados de cada aba de cada arquivo

## Exemplo de Estrutura de Entrada

### Arquivo Excel Input
- **Aba 1** (não-graph):
  - A9: "POÇO-001"
  - J9: "1500"
  - J10: "50"
  - J11: "0.2"
  - J12: "0.15"
  - A18:J18: dados numéricos

- **Aba 2** (não-graph):
  - Similar à aba 1 com dados diferentes

- **Aba "graph"**: Ignorada pelo script

## Notas Importantes
- O script ignora automaticamente abas com nome "graph" (case-insensitive)
- Se uma célula não contiver dados, será deixada em branco
- Os arquivos originais não são modificados
- O arquivo de saída substitui versões anteriores

## Resolução de Problemas

### "Nenhum arquivo Excel encontrado"
- Verifique se há arquivos `.xlsx`, `.xls` ou `.xlsm` na pasta selecionada

### Erro de módulo não encontrado (pandas/openpyxl)
- Execute: `pip install pandas openpyxl`

### Células vazias no resultado
- Verifique se os valores estão realmente nas células indicadas (A9, J9, etc.)

## Autor
Criado com Claude Code
