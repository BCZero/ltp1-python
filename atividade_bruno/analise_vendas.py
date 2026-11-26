"""
analise_vendas.py
=================
Atividade Avaliativa 2º Bimestre - LTP (Python) - SENAC/DF
Aluno:   Bruno Cesar Sampaio Ribeiro de Assis
Tema 1:  Vendas de uma loja
Arquivo: brunoCesarSampaioRibeiroDeAssis.xlsx

Objetivo
--------
Analisar o desempenho de vendas da "BS Tech Store" ao longo dos meses.

Regra de trabalho (como na aula):
    o Pandas PREPARA a resposta e o Matplotlib APRESENTA a resposta.

O que o programa faz, na ordem:
    1. Lê o .xlsx com pd.read_excel()                  (leitura)
    2. Confere o cabeçalho e o mínimo de 30 linhas     (validação)
    3. Mostra as primeiras linhas com head()           (visão geral)
    4. Cria a coluna calculada faturamento             (coluna calculada)
    5. Calcula as análises obrigatórias com groupby()  (agrupamentos)
    6. Aplica filtros nos dados                        (filtros)
    7. Gera os gráficos e salva em PNG com savefig()   (visualização)
    8. Escreve conclusões a partir dos números         (interpretação)

Como executar:
    python analise_vendas.py          -> abre o MENU interativo
    python analise_vendas.py --auto   -> roda tudo direto e só salva os PNGs
                                         (sem abrir janelas de gráfico)
"""

# sys: permite ler argumentos digitados no terminal (ex.: --auto).
import sys

# os: permite criar a pasta onde os gráficos serão salvos.
import os

# NumPy: usado para cálculos numéricos auxiliares (ex.: posições das barras).
import numpy as np

# Pandas: leitura da planilha, limpeza, cálculo, filtro e agrupamento.
import pandas as pd

# Matplotlib (módulo pyplot): criação dos gráficos.
import matplotlib.pyplot as plt

# FuncFormatter: permite formatar os números do eixo (ex.: "R$ 100 mil").
from matplotlib.ticker import FuncFormatter

# ---------------------------------------------------------------------------
# CONSTANTES (configurações usadas em todo o programa)
# ---------------------------------------------------------------------------

# Nome da loja fictícia analisada. Aparece nos títulos dos gráficos e no menu.
# Para trocar o nome da loja, basta alterar esta linha.
NOME_LOJA = "BS Tech Store"

# Nome da planilha do aluno (padrão camelCase exigido pelo professor).
ARQUIVO_XLSX = "brunoCesarSampaioRibeiroDeAssis.xlsx"

# Pasta onde os gráficos PNG serão gravados.
PASTA_GRAFICOS = "graficos"

# Cabeçalho obrigatório do Tema 1, exatamente como no enunciado.
COLUNAS_OBRIGATORIAS = ["produto", "categoria", "mes", "quantidade", "preco_unitario", "regiao"]

# Mínimo de linhas de dados exigido (sem contar o cabeçalho).
MINIMO_LINHAS = 30

# Ordem cronológica dos meses. Sem isso, o Pandas ordenaria em ordem
# ALFABÉTICA (Abril, Fevereiro, Janeiro...), o que quebraria o gráfico de linha.
ORDEM_MESES = [
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro",
]

# Paleta de cores usada nos gráficos (códigos hexadecimais, como na aula).
COR_PRINCIPAL = "#2E86AB"
COR_LINHA = "#B23A48"
COR_REGIAO = "#3A7D44"
COR_PRODUTO = "#6A4C93"
COR_DESTAQUE = "#F18F01"


# ---------------------------------------------------------------------------
# FUNÇÕES AUXILIARES DE FORMATAÇÃO
# ---------------------------------------------------------------------------

def formatar_reais(valor):
    """Converte um número para o padrão brasileiro de moeda.

    Exemplo: 844061.2 -> "R$ 844.061,20"
    O Python usa ponto como separador decimal; aqui trocamos os sinais.
    """
    # :,.2f -> separador de milhar com vírgula e 2 casas decimais (padrão EUA).
    texto = f"{valor:,.2f}"
    # Troca vírgula por ponto e ponto por vírgula usando um caractere temporário.
    texto = texto.replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {texto}"


def formatar_eixo_reais(valor, _posicao):
    """Formata os números do eixo dos gráficos em milhares de reais.

    O Matplotlib chama esta função para cada marca do eixo.
    O segundo parâmetro (_posicao) é exigido pelo FuncFormatter, mas não é usado.
    Exemplo: 150000 -> "R$ 150 mil"
    """
    return f"R$ {valor / 1000:,.0f} mil".replace(",", ".")


def titulo_secao(texto):
    """Imprime um título destacado no terminal, para organizar a saída."""
    print("\n" + "=" * 70)
    print(texto.upper())
    print("=" * 70)


# ---------------------------------------------------------------------------
# ETAPA 1 - LEITURA
# ---------------------------------------------------------------------------

def carregar_dados(nome_arquivo):
    """Lê a planilha .xlsx e devolve um DataFrame.

    pd.read_excel é a função OBRIGATÓRIA da atividade para ler o xlsx.
    engine="openpyxl" informa a biblioteca que interpreta arquivos .xlsx.
    """
    try:
        dados = pd.read_excel(nome_arquivo, engine="openpyxl")
    except FileNotFoundError:
        # Erro comum: o .xlsx não está na mesma pasta em que o programa roda.
        print(f"ERRO: o arquivo '{nome_arquivo}' não foi encontrado.")
        print("Confira se o terminal está aberto na pasta do projeto.")
        sys.exit(1)
    return dados


# ---------------------------------------------------------------------------
# ETAPA 2 - VALIDAÇÃO
# ---------------------------------------------------------------------------

def validar_estrutura(df):
    """Confere se a planilha cumpre as regras do enunciado.

    - O cabeçalho deve ser exatamente o do Tema 1.
    - Deve haver no mínimo 30 linhas de dados.
    Se algo estiver errado, o programa avisa e para.
    """
    titulo_secao("Validação da planilha")

    # list(df.columns) devolve os nomes das colunas na ordem do arquivo.
    colunas_encontradas = list(df.columns)
    if colunas_encontradas != COLUNAS_OBRIGATORIAS:
        print("ERRO: o cabeçalho não segue o padrão do Tema 1.")
        print(f"Esperado:   {COLUNAS_OBRIGATORIAS}")
        print(f"Encontrado: {colunas_encontradas}")
        sys.exit(1)
    print(f"Cabeçalho correto: {', '.join(colunas_encontradas)}")

    # len(df) conta as linhas de dados (o cabeçalho não entra na conta).
    total_linhas = len(df)
    if total_linhas < MINIMO_LINHAS:
        print(f"ERRO: a planilha tem {total_linhas} linhas; o mínimo é {MINIMO_LINHAS}.")
        sys.exit(1)
    print(f"Quantidade de linhas: {total_linhas} (mínimo exigido: {MINIMO_LINHAS}) -> OK")


# ---------------------------------------------------------------------------
# ETAPA 3 - VISÃO GERAL DOS DADOS
# ---------------------------------------------------------------------------

def mostrar_visao_geral(df):
    """Apresenta a planilha: primeiras linhas, tamanho e tipos de dados."""
    titulo_secao("Visão geral dos dados")

    # Configura o Pandas para mostrar todas as colunas lado a lado no terminal.
    pd.set_option("display.max_columns", None)
    pd.set_option("display.width", 120)

    # head() mostra as 5 primeiras linhas (uso OBRIGATÓRIO na atividade).
    print("Primeiras 5 linhas (df.head()):")
    print(df.head())

    # shape devolve uma tupla (linhas, colunas).
    linhas, colunas = df.shape
    print(f"\nTamanho da base: {linhas} linhas x {colunas} colunas")

    # dtypes mostra o tipo de cada coluna (texto, inteiro, decimal...).
    print("\nTipos de dados de cada coluna (df.dtypes):")
    print(df.dtypes)


# ---------------------------------------------------------------------------
# ETAPA 4 - PREPARAÇÃO E COLUNA CALCULADA
# ---------------------------------------------------------------------------

def preparar_dados(df):
    """Limpa os dados e cria a coluna calculada faturamento.

    Passos:
    1. Faz uma cópia para não alterar o DataFrame original.
    2. Remove espaços extras dos textos (ex.: " Sul " vira "Sul").
    3. Remove linhas sem quantidade ou sem preço (não dá para calcular).
    4. Garante que quantidade é inteiro e preço é decimal.
    5. Transforma 'mes' em categoria ORDENADA (ordem do calendário).
    6. Cria faturamento = quantidade * preco_unitario  (coluna OBRIGATÓRIA).
    """
    dados = df.copy()

    # Remove espaços no início e no fim de cada coluna de texto.
    for coluna in ["produto", "categoria", "mes", "regiao"]:
        dados[coluna] = dados[coluna].astype(str).str.strip()

    # Padroniza o mês com a primeira letra maiúscula (ex.: "maio" -> "Maio").
    dados["mes"] = dados["mes"].str.capitalize()

    # dropna remove linhas vazias apenas nas colunas usadas no cálculo.
    dados = dados.dropna(subset=["quantidade", "preco_unitario"])

    # astype converte o tipo da coluna.
    dados["quantidade"] = dados["quantidade"].astype(int)
    dados["preco_unitario"] = dados["preco_unitario"].astype(float)

    # pd.Categorical com ordered=True faz o Pandas respeitar a ordem
    # cronológica dos meses nos agrupamentos e nos gráficos.
    meses_presentes = [m for m in ORDEM_MESES if m in dados["mes"].unique()]
    dados["mes"] = pd.Categorical(dados["mes"], categories=meses_presentes, ordered=True)

    # COLUNA CALCULADA OBRIGATÓRIA DO TEMA 1.
    # A multiplicação é feita linha a linha (operação vetorizada do Pandas).
    # round(2) arredonda para centavos.
    dados["faturamento"] = (dados["quantidade"] * dados["preco_unitario"]).round(2)

    titulo_secao("Coluna calculada criada")
    print("faturamento = quantidade * preco_unitario")
    print(dados[["produto", "mes", "quantidade", "preco_unitario", "faturamento"]].head())

    return dados


# ---------------------------------------------------------------------------
# ETAPA 5 - ANÁLISES OBRIGATÓRIAS (groupby)
# ---------------------------------------------------------------------------

def calcular_indicadores(df):
    """Calcula as cinco análises obrigatórias do Tema 1.

    Devolve um dicionário com os resultados, para reaproveitar nos
    gráficos e nas conclusões sem repetir contas.
    """
    # 1) Faturamento total: soma de toda a coluna.
    faturamento_total = df["faturamento"].sum()

    # groupby agrupa as linhas que têm o mesmo valor na coluna informada;
    # ["faturamento"].sum() soma o faturamento de cada grupo;
    # sort_values(ascending=False) ordena do maior para o menor.

    # 2) Faturamento por produto -> o primeiro é o produto campeão.
    por_produto = df.groupby("produto")["faturamento"].sum().sort_values(ascending=False)

    # 3) Faturamento por categoria -> a primeira é a categoria campeã.
    por_categoria = df.groupby("categoria")["faturamento"].sum().sort_values(ascending=False)

    # 4) Faturamento por mês. observed=True considera só os meses existentes.
    #    Aqui NÃO ordenamos por valor: mantemos a ordem do calendário.
    por_mes = df.groupby("mes", observed=True)["faturamento"].sum()

    # 5) Faturamento por região.
    por_regiao = df.groupby("regiao")["faturamento"].sum().sort_values(ascending=False)

    # Agrupamento extra com agg(): várias estatísticas de uma vez por categoria.
    resumo_categoria = (
        df.groupby("categoria")
        .agg(
            unidades=("quantidade", "sum"),       # total de unidades vendidas
            faturamento=("faturamento", "sum"),   # total faturado
            registros=("produto", "count"),       # número de linhas de venda
        )
        .sort_values("faturamento", ascending=False)
    )

    # Unidades vendidas por produto (para comparar volume x faturamento).
    unidades_produto = df.groupby("produto")["quantidade"].sum().sort_values(ascending=False)

    return {
        "faturamento_total": faturamento_total,
        "por_produto": por_produto,
        "por_categoria": por_categoria,
        "por_mes": por_mes,
        "por_regiao": por_regiao,
        "resumo_categoria": resumo_categoria,
        "unidades_produto": unidades_produto,
    }


def exibir_indicadores(ind):
    """Mostra no terminal os resultados das análises obrigatórias."""
    titulo_secao("Análises obrigatórias")

    print(f"1) Faturamento total: {formatar_reais(ind['faturamento_total'])}")

    # idxmax devolve o RÓTULO (nome) do maior valor; max devolve o valor.
    produto_top = ind["por_produto"].idxmax()
    print(f"2) Produto com maior faturamento: {produto_top} "
          f"({formatar_reais(ind['por_produto'].max())})")

    categoria_top = ind["por_categoria"].idxmax()
    print(f"3) Categoria com maior faturamento: {categoria_top} "
          f"({formatar_reais(ind['por_categoria'].max())})")

    print("4) Faturamento por mês:")
    for mes, valor in ind["por_mes"].items():
        print(f"   {mes:<10} {formatar_reais(valor):>16}")

    print("5) Faturamento por região:")
    for regiao, valor in ind["por_regiao"].items():
        print(f"   {regiao:<13} {formatar_reais(valor):>16}")

    print("\nResumo por categoria (groupby + agg):")
    print(ind["resumo_categoria"])


# ---------------------------------------------------------------------------
# ETAPA 6 - FILTROS
# ---------------------------------------------------------------------------

def filtrar_acima_da_media(df):
    """FILTRO 1: vendas com faturamento acima da média das vendas.

    Pergunta respondida: quais vendas "puxam" o faturamento da loja?
    A expressão df["faturamento"] > media gera uma lista de True/False;
    df[...] mantém apenas as linhas marcadas como True.
    """
    titulo_secao("Filtro 1 - Vendas acima da média")

    media = df["faturamento"].mean()
    acima = df[df["faturamento"] > media]

    # Participação dessas vendas no faturamento total da loja.
    participacao = acima["faturamento"].sum() / df["faturamento"].sum() * 100

    print(f"Faturamento médio por venda: {formatar_reais(media)}")
    print(f"Vendas acima da média: {len(acima)} de {len(df)} "
          f"({len(acima) / len(df) * 100:.1f}% dos registros)")
    print(f"Essas vendas somam {participacao:.1f}% do faturamento total.")
    print("Produtos que aparecem nesse grupo:")
    # value_counts conta quantas vezes cada produto aparece no filtro.
    print(acima["produto"].value_counts())

    return acima, media, participacao


def filtrar_por_regiao(df, regiao):
    """FILTRO 2: mantém apenas as vendas da região escolhida.

    A comparação é feita em minúsculas para aceitar "sul", "SUL" ou "Sul".
    Retorna um DataFrame vazio se a região não existir.
    """
    return df[df["regiao"].str.lower() == regiao.strip().lower()]


def filtrar_categoria_no_mes(df, categoria, mes):
    """FILTRO 3 (combinado): vendas de uma categoria em um mês específico.

    O operador & significa "E": as duas condições precisam ser verdadeiras.
    Cada condição fica entre parênteses, como exige o Pandas.
    """
    return df[(df["categoria"] == categoria) & (df["mes"] == mes)]


# ---------------------------------------------------------------------------
# ETAPA 7 - GRÁFICOS
# ---------------------------------------------------------------------------

def salvar_e_exibir(fig, nome_arquivo, mostrar):
    """Salva a figura em PNG (fig.savefig) e, se pedido, abre a janela.

    - os.makedirs cria a pasta 'graficos' se ela ainda não existir.
    - dpi=300 gera imagem com boa resolução (adequada para impressão).
    - bbox_inches="tight" evita que títulos e rótulos fiquem cortados.
    - plt.close libera a memória quando a janela não será exibida.
    """
    os.makedirs(PASTA_GRAFICOS, exist_ok=True)
    caminho = os.path.join(PASTA_GRAFICOS, nome_arquivo)

    plt.tight_layout()
    fig.savefig(caminho, dpi=300, bbox_inches="tight")
    print(f"Gráfico salvo em: {caminho}")

    if mostrar:
        plt.show()
    else:
        plt.close(fig)


def escrever_valores_nas_barras(ax, valores, horizontal=False):
    """Escreve o valor de cada barra logo acima (ou ao lado) dela.

    enumerate fornece a posição e o valor de cada barra ao mesmo tempo.
    """
    for posicao, valor in enumerate(valores):
        texto = formatar_reais(valor).replace(",00", "")
        if horizontal:
            # Em barras horizontais o texto vai à direita da barra.
            ax.text(valor, posicao, " " + texto, va="center", fontsize=8)
        else:
            # ha="center" centraliza; va="bottom" coloca o texto acima da barra.
            ax.text(posicao, valor, texto, ha="center", va="bottom", fontsize=9)


def gerar_barras_categoria(ind, mostrar=True):
    """GRÁFICO 1 (obrigatório) - BARRAS: faturamento por categoria.

    Por que barras? O objetivo é COMPARAR grupos diferentes (categorias).
    A altura da barra torna a comparação imediata.
    """
    dados = ind["por_categoria"]

    fig, ax = plt.subplots(figsize=(9, 5))
    ax.bar(dados.index, dados.values, color=COR_PRINCIPAL)
    escrever_valores_nas_barras(ax, dados.values)

    ax.set_title(f"Faturamento por categoria - {NOME_LOJA} (jan a jun)", fontsize=13, pad=12)
    ax.set_xlabel("Categoria")
    ax.set_ylabel("Faturamento (R$)")
    # Formata o eixo Y como "R$ 100 mil" em vez de "100000".
    ax.yaxis.set_major_formatter(FuncFormatter(formatar_eixo_reais))
    # margins dá um respiro no topo para o texto da maior barra não cortar.
    ax.margins(y=0.12)
    ax.grid(axis="y", linestyle="--", alpha=0.4)
    # set_axisbelow coloca a grade atrás das barras.
    ax.set_axisbelow(True)

    salvar_e_exibir(fig, "01_faturamento_por_categoria.png", mostrar)


def gerar_linha_mensal(ind, mostrar=True):
    """GRÁFICO 2 (obrigatório) - LINHA: faturamento por mês.

    Por que linha? Os meses seguem uma ORDEM NO TEMPO. A linha mostra
    tendência: crescimento, queda e picos.
    """
    dados = ind["por_mes"]
    # Converte os rótulos (categorias) para texto, para usar no eixo X.
    meses = [str(m) for m in dados.index]

    fig, ax = plt.subplots(figsize=(9, 5))
    ax.plot(meses, dados.values, marker="o", color=COR_LINHA, linewidth=2)

    # Destaca o mês de pico com um marcador maior e uma anotação.
    # np.argmax devolve a POSIÇÃO do maior valor.
    pos_pico = int(np.argmax(dados.values))
    ax.scatter(meses[pos_pico], dados.values[pos_pico], s=150, color=COR_DESTAQUE, zorder=3)
    ax.annotate(
        f"Pico: {formatar_reais(dados.values[pos_pico])}",
        xy=(pos_pico, dados.values[pos_pico]),
        xytext=(0, 12),
        textcoords="offset points",
        ha="center",
        fontsize=9,
    )

    # Linha tracejada com a média mensal, como referência de comparação.
    media_mensal = dados.mean()
    ax.axhline(media_mensal, color="gray", linestyle=":", linewidth=1.5,
               label=f"Média mensal: {formatar_reais(media_mensal)}")

    ax.set_title(f"Evolução do faturamento mensal - {NOME_LOJA}", fontsize=13, pad=12)
    ax.set_xlabel("Mês")
    ax.set_ylabel("Faturamento (R$)")
    ax.yaxis.set_major_formatter(FuncFormatter(formatar_eixo_reais))
    ax.margins(y=0.15)
    ax.grid(True, linestyle="--", alpha=0.4)
    ax.legend(loc="lower right")

    salvar_e_exibir(fig, "02_faturamento_por_mes.png", mostrar)


def gerar_barras_regiao(ind, mostrar=True):
    """GRÁFICO 3 (obrigatório) - BARRAS: faturamento por região.

    O enunciado permite pizza OU barras. Escolhemos BARRAS porque:
    - são 5 regiões, e algumas têm valores próximos (difícil comparar fatias);
    - barras permitem título e nome nos dois eixos (critério de avaliação);
    - o percentual de participação é escrito acima de cada barra, então a
      informação de "parte do total" que a pizza daria não se perde.
    """
    dados = ind["por_regiao"]
    total = ind["faturamento_total"]

    fig, ax = plt.subplots(figsize=(9, 5))
    ax.bar(dados.index, dados.values, color=COR_REGIAO)

    # Escreve valor e participação percentual acima de cada barra.
    for posicao, valor in enumerate(dados.values):
        ax.text(posicao, valor,
                f"{formatar_reais(valor).replace(',00', '')}\n({valor / total * 100:.1f}%)",
                ha="center", va="bottom", fontsize=9)

    ax.set_title(f"Faturamento por região - {NOME_LOJA}", fontsize=13, pad=12)
    ax.set_xlabel("Região")
    ax.set_ylabel("Faturamento (R$)")
    ax.yaxis.set_major_formatter(FuncFormatter(formatar_eixo_reais))
    ax.margins(y=0.18)
    ax.grid(axis="y", linestyle="--", alpha=0.4)
    ax.set_axisbelow(True)

    salvar_e_exibir(fig, "03_faturamento_por_regiao.png", mostrar)


def gerar_barras_produto(ind, mostrar=True):
    """GRÁFICO 4 (extra) - BARRAS HORIZONTAIS: faturamento por produto.

    Por que horizontais? Os nomes dos produtos são longos; deitados, eles
    ficam legíveis sem precisar girar o texto.
    """
    # sort_values() crescente: o maior fica no topo do gráfico horizontal.
    dados = ind["por_produto"].sort_values()

    fig, ax = plt.subplots(figsize=(10, 6))
    # Pinta o produto campeão com a cor de destaque.
    cores = [COR_DESTAQUE if p == dados.idxmax() else COR_PRODUTO for p in dados.index]
    ax.barh(dados.index, dados.values, color=cores)
    escrever_valores_nas_barras(ax, dados.values, horizontal=True)

    ax.set_title(f"Faturamento por produto - {NOME_LOJA}", fontsize=13, pad=12)
    ax.set_xlabel("Faturamento (R$)")
    ax.set_ylabel("Produto")
    ax.xaxis.set_major_formatter(FuncFormatter(formatar_eixo_reais))
    ax.margins(x=0.18)
    ax.grid(axis="x", linestyle="--", alpha=0.4)
    ax.set_axisbelow(True)

    salvar_e_exibir(fig, "04_faturamento_por_produto.png", mostrar)


def gerar_barras_regiao_escolhida(df, regiao, mostrar=True):
    """GRÁFICO 5 (extra, usa FILTRO 2): categorias dentro de uma região.

    Ordem correta: primeiro FILTRA a região, depois AGRUPA por categoria.
    Se agrupássemos antes, perderíamos as linhas individuais do filtro.
    """
    dados_regiao = filtrar_por_regiao(df, regiao)

    # empty é True quando nenhuma linha passou pelo filtro.
    if dados_regiao.empty:
        regioes = ", ".join(sorted(df["regiao"].unique()))
        print(f"Região '{regiao}' não encontrada. Opções: {regioes}")
        return

    # Usa o nome como está na planilha (com maiúsculas corretas).
    nome_regiao = dados_regiao["regiao"].iloc[0]
    totais = dados_regiao.groupby("categoria")["faturamento"].sum().sort_values(ascending=False)

    titulo_secao(f"Filtro 2 - Região {nome_regiao}")
    print(f"Registros encontrados: {len(dados_regiao)}")
    print(f"Faturamento da região: {formatar_reais(dados_regiao['faturamento'].sum())}")
    print(totais)

    fig, ax = plt.subplots(figsize=(9, 5))
    ax.bar(totais.index, totais.values, color=COR_PRINCIPAL)
    escrever_valores_nas_barras(ax, totais.values)

    ax.set_title(f"Faturamento por categoria na região {nome_regiao}", fontsize=13, pad=12)
    ax.set_xlabel("Categoria")
    ax.set_ylabel("Faturamento (R$)")
    ax.yaxis.set_major_formatter(FuncFormatter(formatar_eixo_reais))
    ax.margins(y=0.12)
    ax.grid(axis="y", linestyle="--", alpha=0.4)
    ax.set_axisbelow(True)

    # replace troca espaços e hífen para gerar um nome de arquivo seguro.
    nome_arquivo = "05_regiao_" + nome_regiao.lower().replace(" ", "_").replace("-", "_") + ".png"
    salvar_e_exibir(fig, nome_arquivo, mostrar)


# ---------------------------------------------------------------------------
# ETAPA 8 - CONCLUSÕES INTERPRETATIVAS
# ---------------------------------------------------------------------------

def gerar_conclusoes(df, ind):
    """Escreve conclusões calculadas a partir dos dados.

    As frases são montadas com os números reais: se a planilha mudar,
    as conclusões se atualizam sozinhas (nada está "chumbado" no texto).
    """
    titulo_secao("Conclusões")

    total = ind["faturamento_total"]

    # Conclusão 1 - concentração por categoria.
    cat = ind["por_categoria"]
    duas_maiores = cat.iloc[:2].sum() / total * 100
    print(f"1. {cat.index[0]} lidera com {cat.iloc[0] / total * 100:.1f}% do faturamento. "
          f"Somadas, {cat.index[0]} e {cat.index[1]} respondem por {duas_maiores:.1f}% "
          f"do total: a loja depende de produtos de alto valor.")

    # Conclusão 2 - volume não é faturamento.
    mais_unidades = ind["unidades_produto"].idxmax()
    unidades = ind["unidades_produto"].max()
    fat_mais_unidades = ind["por_produto"][mais_unidades]
    print(f"2. O produto mais vendido em unidades é {mais_unidades} ({unidades} un.), "
          f"mas ele faturou só {formatar_reais(fat_mais_unidades)}. Já {ind['por_produto'].idxmax()} "
          f"faturou {formatar_reais(ind['por_produto'].max())} com bem menos unidades: "
          f"vender mais peças não significa faturar mais.")

    # Conclusão 3 - sazonalidade mensal.
    mes = ind["por_mes"]
    mes_pico, mes_baixo = mes.idxmax(), mes.idxmin()
    diferenca = (mes.max() - mes.min()) / mes.min() * 100
    print(f"3. {mes_pico} foi o melhor mês ({formatar_reais(mes.max())}) e {mes_baixo} o pior "
          f"({formatar_reais(mes.min())}), uma diferença de {diferenca:.1f}%. O pico coincide "
          f"com datas comemorativas (Dia das Mães), bom momento para reforçar estoque.")

    # Conclusão 4 - regiões (com explicação usando um groupby de contagem).
    reg = ind["por_regiao"]
    vendas_por_regiao = df.groupby("regiao")["produto"].count()
    print(f"4. A região {reg.index[0]} liderou ({reg.iloc[0] / total * 100:.1f}% do total), "
          f"com {vendas_por_regiao[reg.index[0]]} registros de venda; {reg.index[-1]} ficou por "
          f"último ({reg.iloc[-1] / total * 100:.1f}%), com {vendas_por_regiao[reg.index[-1]]} registros. "
          f"A diferença vem mais do número de vendas e dos produtos caros vendidos em cada região.")


# ---------------------------------------------------------------------------
# FLUXOS DE EXECUÇÃO
# ---------------------------------------------------------------------------

def executar_relatorio_completo(df, ind, mostrar=True, regiao_filtro="Centro-Oeste"):
    """Executa todas as etapas da análise em sequência (modo apresentação)."""
    exibir_indicadores(ind)
    filtrar_acima_da_media(df)

    # Exemplo do filtro combinado: Celulares em Maio (mês de pico).
    titulo_secao("Filtro 3 - Celulares em Maio (filtro combinado com &)")
    print(filtrar_categoria_no_mes(df, "Celulares", "Maio")
          [["produto", "quantidade", "preco_unitario", "faturamento", "regiao"]])

    titulo_secao("Gráficos")
    gerar_barras_categoria(ind, mostrar)
    gerar_linha_mensal(ind, mostrar)
    gerar_barras_regiao(ind, mostrar)
    gerar_barras_produto(ind, mostrar)
    gerar_barras_regiao_escolhida(df, regiao_filtro, mostrar)

    gerar_conclusoes(df, ind)


def exibir_menu(df, ind):
    """Menu interativo: útil para mostrar cada parte no vídeo, uma por vez."""
    while True:
        print(f"\n----------------- MENU - {NOME_LOJA.upper()} -----------------")
        print("1 - Análises obrigatórias (totais, produto, categoria, mês, região)")
        print("2 - Filtro: vendas acima da média")
        print("3 - Gráfico de barras: faturamento por categoria")
        print("4 - Gráfico de linha: faturamento por mês")
        print("5 - Gráfico de barras: faturamento por região")
        print("6 - Gráfico extra: faturamento por produto")
        print("7 - Filtro por região (você digita) + gráfico")
        print("8 - Conclusões")
        print("9 - Rodar tudo em sequência")
        print("0 - Sair")

        # input devolve texto; por isso as comparações usam aspas.
        opcao = input("Escolha uma opção: ").strip()

        if opcao == "1":
            exibir_indicadores(ind)
        elif opcao == "2":
            filtrar_acima_da_media(df)
        elif opcao == "3":
            gerar_barras_categoria(ind)
        elif opcao == "4":
            gerar_linha_mensal(ind)
        elif opcao == "5":
            gerar_barras_regiao(ind)
        elif opcao == "6":
            gerar_barras_produto(ind)
        elif opcao == "7":
            regiao = input("Digite a região (ex.: Sul, Nordeste, Centro-Oeste): ")
            gerar_barras_regiao_escolhida(df, regiao)
        elif opcao == "8":
            gerar_conclusoes(df, ind)
        elif opcao == "9":
            executar_relatorio_completo(df, ind)
        elif opcao == "0":
            print("Encerrando o programa.")
            break
        else:
            print("Opção inválida. Digite um número de 0 a 9.")


def main():
    """Ponto de entrada do programa: lê, valida, prepara e analisa."""
    print("Atividade Avaliativa 2BI - Tema 1: Vendas de uma loja")
    print("Aluno: Bruno Cesar Sampaio Ribeiro de Assis")
    print(f"Loja analisada: {NOME_LOJA}")
    print(f"Arquivo analisado: {ARQUIVO_XLSX}")

    vendas = carregar_dados(ARQUIVO_XLSX)
    validar_estrutura(vendas)
    mostrar_visao_geral(vendas)
    vendas = preparar_dados(vendas)
    indicadores = calcular_indicadores(vendas)

    # sys.argv é a lista de palavras digitadas no terminal após "python".
    if "--auto" in sys.argv:
        # Modo automático: gera tudo e só salva os PNGs (sem abrir janelas).
        executar_relatorio_completo(vendas, indicadores, mostrar=False)
    else:
        exibir_menu(vendas, indicadores)


# Só executa main() quando o arquivo é rodado diretamente.
if __name__ == "__main__":
    main()
