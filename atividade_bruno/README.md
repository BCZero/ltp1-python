# Atividade Avaliativa 2º Bimestre: Pandas e Matplotlib

| | |
|---|---|
| **Aluno** | Bruno Cesar Sampaio Ribeiro de Assis |
| **Disciplina** | Linguagem Técnica de Programação (Python), SENAC/DF |
| **Tema** | Tema 1: Vendas de uma loja |
| **Planilha** | `brunoCesarSampaioRibeiroDeAssis.xlsx` |
| **Loja analisada** | BS Tech Store (informática e eletrônicos), janeiro a junho |

A planilha de vendas é lida com **Pandas**. O programa cria a coluna `faturamento`, faz as cinco análises obrigatórias com `groupby()`, aplica filtros, gera gráficos com **Matplotlib** (salvos em PNG) e escreve as conclusões a partir dos números.

---

## 1. Estrutura do projeto

```
atividade_bruno/
├── brunoCesarSampaioRibeiroDeAssis.xlsx   # base de dados (60 linhas, cabeçalho do Tema 1)
├── analise_vendas.py                      # programa principal da análise
├── gerar_planilha.py                      # script que criou a planilha (NumPy + Pandas)
├── requirements.txt                       # bibliotecas necessárias
├── README.md                              # este arquivo
├── PLANO_IMPLEMENTACAO.md                 # planejamento da atividade e do código
├── ROTEIRO_APRESENTACAO.md                # roteiro do vídeo de 4 a 7 minutos
└── graficos/                              # PNGs gerados pelo programa
    ├── 01_faturamento_por_categoria.png   # barras (obrigatório)
    ├── 02_faturamento_por_mes.png         # linha  (obrigatório)
    ├── 03_faturamento_por_regiao.png      # barras (obrigatório: "pizza ou barras")
    ├── 04_faturamento_por_produto.png     # barras horizontais (extra)
    └── 05_regiao_centro_oeste.png         # gráfico do filtro por região (extra)
```

### Cabeçalho da planilha (exatamente como pede o Tema 1)

| Coluna | Tipo | O que representa |
|---|---|---|
| `produto` | texto | Nome do produto vendido (ex.: Notebook) |
| `categoria` | texto | Grupo do produto: Informática, Celulares, Periféricos ou Acessórios |
| `mes` | texto | Mês da venda (Janeiro a Junho) |
| `quantidade` | inteiro | Unidades vendidas naquele registro |
| `preco_unitario` | decimal | Preço de uma unidade, em R$ |
| `regiao` | texto | Região do Brasil onde a venda ocorreu |

Coluna calculada pelo programa: **`faturamento = quantidade * preco_unitario`**

---

## 2. Preparando o ambiente

### 2.1 Pré-requisitos
- **Python 3.10 ou superior**. Para conferir, rode `python --version` no terminal.
- **VS Code** com a extensão *Python* (Microsoft).

### 2.2 Abrir a pasta no VS Code
1. `Arquivo > Abrir Pasta...` e escolha a pasta `atividade_bruno`.
2. Abra o terminal: `Terminal > Novo Terminal` (ou `Ctrl + '`).
3. Confira que o terminal está **dentro da pasta do projeto**. O comando `dir` (Windows) ou `ls` (Linux/Mac) deve listar o `.xlsx` e os `.py`.

### 2.3 (Recomendado) Criar um ambiente virtual
O ambiente virtual isola as bibliotecas deste trabalho das demais instalações do Python.

**Windows (PowerShell ou CMD):**
```bash
python -m venv .venv
.venv\Scripts\activate
```
Se o PowerShell bloquear a ativação, rode uma vez:
`Set-ExecutionPolicy -Scope CurrentUser RemoteSigned` e tente de novo.

**Linux / macOS:**
```bash
python3 -m venv .venv
source .venv/bin/activate
```

Com o ambiente ativo, o terminal mostra `(.venv)` no início da linha. No VS Code, aperte `Ctrl + Shift + P`, escolha **Python: Select Interpreter** e selecione o `.venv`.

### 2.4 Instalar as bibliotecas

**Opção A: tudo de uma vez (recomendado):**
```bash
python -m pip install -r requirements.txt
```

**Opção B: comandos do material do professor, um por um:**
```bash
python -m pip install numpy
python -m pip install pandas
python -m pip install openpyxl
python -m pip install matplotlib

# ou, em uma linha só (como na aula de Pandas e Matplotlib)
python -m pip install pandas matplotlib openpyxl
```

**Alternativas comuns** (dependem de como o Python foi instalado):
```bash
# Windows com o "py launcher"
py -m pip install numpy pandas openpyxl matplotlib

# Linux / macOS
python3 -m pip install numpy pandas openpyxl matplotlib

# forma curta (funciona quando o pip está no PATH)
pip install pandas openpyxl
```

Para conferir a instalação:
```bash
python -c "import numpy, pandas, openpyxl, matplotlib; print('Tudo instalado!')"
```

| Biblioteca | Para que serve neste trabalho |
|---|---|
| `numpy` | Gerar os números aleatórios da planilha e localizar o mês de pico (`np.argmax`) |
| `pandas` | `read_excel`, `head`, coluna calculada, filtros, `groupby` |
| `openpyxl` | Permite ao Pandas ler e gravar `.xlsx` |
| `matplotlib` | Gráficos e `fig.savefig()` em PNG |

---

## 3. Como executar

### 3.1 Modo menu (use este na apresentação)
```bash
python analise_vendas.py
```
O programa valida a planilha, mostra o `head()`, cria a coluna `faturamento` e abre o menu:

```
1 - Análises obrigatórias (totais, produto, categoria, mês, região)
2 - Filtro: vendas acima da média
3 - Gráfico de barras: faturamento por categoria
4 - Gráfico de linha: faturamento por mês
5 - Gráfico de barras: faturamento por região
6 - Gráfico extra: faturamento por produto
7 - Filtro por região (você digita) + gráfico
8 - Conclusões
9 - Rodar tudo em sequência
0 - Sair
```
Cada gráfico abre em uma janela. **Feche a janela** para o menu voltar.

### 3.2 Modo automático (gera tudo sem abrir janelas)
```bash
python analise_vendas.py --auto
```
Roda todas as etapas e grava os PNGs na pasta `graficos/`. Use esse modo para conferir se está tudo funcionando antes de gravar o vídeo.

### 3.3 Recriar a planilha (opcional)
```bash
python gerar_planilha.py
```
Esse comando **sobrescreve** o `.xlsx`. A semente aleatória é fixa (`SEMENTE = 2026`), então o arquivo sai sempre igual. Se você editar a planilha à mão no Excel, **não** rode este script depois, ou perderá as edições.

---

## 4. O que o código faz (mapa dos requisitos)

| Requisito do enunciado | Onde está no código |
|---|---|
| Leitura com `pd.read_excel()` | `carregar_dados()` |
| Uso de `head()` | `mostrar_visao_geral()` |
| Mínimo de 30 linhas / cabeçalho correto | `validar_estrutura()` |
| Coluna calculada `faturamento` | `preparar_dados()` |
| `groupby()` | `calcular_indicadores()` (5 agrupamentos + `agg`) |
| Filtro de dados | `filtrar_acima_da_media()`, `filtrar_por_regiao()`, `filtrar_categoria_no_mes()` |
| Gráfico de barras por categoria | `gerar_barras_categoria()` |
| Gráfico de linha por mês | `gerar_linha_mensal()` |
| Pizza ou barras por região | `gerar_barras_regiao()` |
| PNG com `fig.savefig()` | `salvar_e_exibir()` (todos os gráficos) |
| Conclusão interpretativa | `gerar_conclusoes()` |

**Análises obrigatórias do Tema 1:** faturamento total, produto com maior faturamento, categoria com maior faturamento, faturamento por mês e faturamento por região. Todas aparecem na opção 1 do menu.

---

## 5. Resultados obtidos (planilha atual)

| Indicador | Resultado |
|---|---|
| Faturamento total | R$ 844.061,20 |
| Produto com maior faturamento | Notebook (R$ 262.492,50) |
| Categoria com maior faturamento | Informática (R$ 379.778,60, 45,0%) |
| Melhor mês / pior mês | Maio (R$ 180.901,50) / Março (R$ 116.164,50) |
| Região com maior faturamento | Sul (R$ 256.958,20, 30,4%) |

---

## 6. Problemas comuns

| Mensagem / situação | Causa | Solução |
|---|---|---|
| `ModuleNotFoundError: No module named 'pandas'` | Biblioteca instalada em outro Python | Selecione o interpretador certo no VS Code e reinstale com `python -m pip install -r requirements.txt` |
| `ImportError: Missing optional dependency 'openpyxl'` | Falta o openpyxl | `python -m pip install openpyxl` |
| `ERRO: o arquivo '...xlsx' não foi encontrado` | O terminal está em outra pasta | Use `cd` até a pasta do projeto |
| `PermissionError` ao rodar `gerar_planilha.py` | O `.xlsx` está aberto no Excel | Feche o Excel e rode de novo |
| O menu "travou" depois de um gráfico | A janela do gráfico ainda está aberta | Feche a janela do gráfico |
| Meses em ordem alfabética | Coluna `mes` tratada como texto comum | O código já converte para `pd.Categorical` ordenado |
| `'python' não é reconhecido...` | Python fora do PATH | Use `py` no lugar de `python` (Windows) |
