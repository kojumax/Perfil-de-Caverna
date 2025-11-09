
# Descrição

Cav.py é um script Python desenvolvido para processar dados topográficos de cavernas e gerar visualizações gráficas do perfil das cavidades. O programa lê dados de levantamentos topográficos a partir de arquivos Excel (.xls, .xlsx) ou Word (.docx) e cria representações visuais bidimensionais das cavernas.


## Funcionalidades

Processamento de múltiplos formatos: Suporte para arquivos Excel e Word

Cálculo automático de coordenadas: Converte medições de distância e ângulo em coordenadas cartesianas

Visualização completa: Planta baixa com alturas (HT - teto, HB - base)

Processamento em lote: Capacidade de processar múltiplos arquivos automaticamente


# Estrutura dos Dados

### Colunas necessárias:
EST.: Ponto de estação (origem da medição)

PV.: Ponto visado (destino da medição)

DI: Distância inclinada entre os pontos

αc: Ângulo vertical (positivo ou negativo)

HT: Altura total (teto)

HB: Altura da base



## Como usar

1. Processamento automático
Coloque os arquivos na mesma pasta do script e execute:

```bash
python Cav.py
```
2. Processar pasta específica
```
python
process_files(folder_path="caminho/para/sua/pasta")
```
3. Processar arquivos específicos
```
python
process_files(specific_files=["arquivo1.xlsx", "arquivo2.docx"])
```

## Formatos de arquivo suportados
### Arquivo Excel (.xlsx, .xls)
```Planilha deve se chamar "Plan1" (altera linha 62)```

Estrutura de colunas esperada:

Coluna A: EST.

Coluna B: PV

Coluna D: DI

Coluna E: αc

Coluna K: HB

Coluna L: Observações (para extrair HT)

### Arquivo Word (.docx)
```Primeira tabela do documento```
estrutura de colunas esperada:

Coluna 0: EST.

Coluna 1: PV

Coluna 3-4: αc (positivo/negativo)

Coluna 6: DI

Coluna 12: HT

Coluna 13: HB


## Saída do programa
Para cada arquivo processado, o programa gera:

#### 1. Gráfico visual mostrando:

Planta baixa da caverna

Conexões entre pontos

Representação de alturas (HT e HB)

Legenda completa

#### 2. Resumo no console com:

Número de pontos processados

Quantidade de medições

Estatísticas do arquivo

### Símbolos no gráfico
🔴 Ponto vermelho: Estação/Ponto topográfico

📏 Linha preta: Altura total (HT + HB)

🔺 Triângulo vermelho: HT (Teto da caverna)

🔻 Triângulo azul: HB (Base da caverna)

🔷 Linha azul: Conexões horizontais entre pontos

## Dependências
matplotlib - Geração de gráficos

pandas - Processamento de planilhas Excel

python-docx - Leitura de arquivos Word

pathlib - Manipulação de caminhos de arquivos

## Observações importantes
O primeiro ponto do levantamento é considerado como origem (0,0)

Ângulos negativos são automaticamente convertidos para positivos equivalentes

O programa ignora linhas vazias e cabeçalhos automaticamente

Para arquivos Excel, o HT pode ser extraído das observações usando padrões como "Ht. T0 = 6m"

## Exemplo de uso típico
Colete os dados topográficos da caverna

Organize-os no formato Excel ou Word conforme a estrutura esperada

Execute o script

Visualize os gráficos gerados para cada arquivo

Analise a planta baixa e o perfil vertical da caverna

## Limitações
Visualização apenas em 2D (planta baixa com alturas)

Não considera superfícies irregulares entre pontos

Assume medições consecutivas e conectadas

Suporte
Em caso de problemas, verifique:

Formatação correta dos arquivos de entrada

Instalação de todas as dependências

Permissões de leitura dos arquivos

