"""
gerar_planilha.py
=================
Atividade Avaliativa 2º Bimestre - LTP (Python) - SENAC/DF
Aluno: Bruno Cesar Sampaio Ribeiro de Assis
Tema 1 - Vendas de uma loja

Finalidade deste script
-----------------------
Criar a base de dados própria do aluno, no formato .xlsx, seguindo
EXATAMENTE o cabeçalho exigido pelo Tema 1:

    produto, categoria, mes, quantidade, preco_unitario, regiao

Contexto inventado para a base: a "BS Tech Store", uma loja de
informática e eletrônicos com vendas em todas as regiões do Brasil,
de janeiro a junho.

Como os dados são montados (para que sejam coerentes):
- Cada produto tem um preço fixo, compatível com o mercado.
- Cada produto tem uma "quantidade base" mensal (ex.: cabos vendem
  muito mais unidades que notebooks).
- Cada mês tem um fator de sazonalidade (ex.: maio é mais forte por
  causa do Dia das Mães, fevereiro por causa da volta às aulas).
- Cada região tem um peso (a região Sudeste tem mais clientes).
- O NumPy sorteia a região de cada venda e aplica uma pequena variação
  aleatória na quantidade, para os números não ficarem "certinhos".

A semente aleatória (SEMENTE) é fixa: rodar o script de novo gera
sempre a MESMA planilha. Se mudar a semente, a planilha muda.

Como executar:
    python gerar_planilha.py
"""

# NumPy: usado aqui para gerar números aleatórios de forma controlada.
import numpy as np

# Pandas: usado para montar a tabela (DataFrame) e salvá-la em .xlsx.
import pandas as pd

# ---------------------------------------------------------------------------
# CONFIGURAÇÕES
# ---------------------------------------------------------------------------

# Nome do arquivo exigido pelo professor: nome do aluno em camelCase,
# sem acentos e sem espaços.
NOME_ARQUIVO = "brunoCesarSampaioRibeiroDeAssis.xlsx"

# Semente do gerador aleatório. Mantê-la fixa garante que a planilha
# seja reproduzível (sempre os mesmos números).
SEMENTE = 2026

# Cabeçalho obrigatório do Tema 1, na ordem exata do enunciado.
COLUNAS = ["produto", "categoria", "mes", "quantidade", "preco_unitario", "regiao"]

# Catálogo da loja: produto -> (categoria, preço unitário, quantidade base/mês).
# A quantidade base é a média de unidades que o produto vende em um mês "normal".
CATALOGO = {
    "Notebook": ("Informática", 3499.90, 12),
    "Monitor 24 polegadas": ("Informática", 899.90, 15),
    "Impressora Multifuncional": ("Informática", 749.90, 8),
    "Teclado Mecânico": ("Periféricos", 299.90, 25),
    "Mouse Gamer": ("Periféricos", 159.90, 35),
    "Headset": ("Periféricos", 249.90, 22),
    "Smartphone": ("Celulares", 1899.90, 18),
    "Smartwatch": ("Celulares", 699.90, 14),
    "Carregador Turbo": ("Acessórios", 89.90, 50),
    "Cabo USB-C": ("Acessórios", 39.90, 80),
}

# Fator de sazonalidade de cada mês (1.0 = mês normal).
SAZONALIDADE = {
    "Janeiro": 1.00,   # mês de referência
    "Fevereiro": 1.15, # volta às aulas: mais notebooks e periféricos
    "Março": 0.90,     # mês mais fraco após a volta às aulas
    "Abril": 0.95,
    "Maio": 1.30,      # Dia das Mães: pico de vendas
    "Junho": 1.10,     # Dia dos Namorados
}

# Regiões e a chance de cada uma aparecer em uma venda (somam 1.0).
PESO_REGIOES = {
    "Sudeste": 0.30,
    "Nordeste": 0.22,
    "Sul": 0.18,
    "Centro-Oeste": 0.18,
    "Norte": 0.12,
}


def gerar_registros():
    """Monta a lista de vendas (uma linha por produto em cada mês).

    São 10 produtos x 6 meses = 60 linhas, o dobro do mínimo exigido (30).
    Retorna uma lista de dicionários, um para cada linha da planilha.
    """
    # Cria o gerador de números aleatórios do NumPy com a semente fixa.
    gerador = np.random.default_rng(SEMENTE)

    # Listas com os nomes e as probabilidades das regiões (mesma ordem).
    regioes = list(PESO_REGIOES.keys())
    probabilidades = list(PESO_REGIOES.values())

    registros = []

    # Percorre cada mês e, dentro dele, cada produto do catálogo.
    for mes, fator_mes in SAZONALIDADE.items():
        for produto, (categoria, preco, quantidade_base) in CATALOGO.items():
            # Sorteia a região da venda respeitando os pesos definidos.
            regiao = gerador.choice(regioes, p=probabilidades)

            # Variação aleatória entre -20% e +20% para dar realismo.
            variacao = gerador.uniform(0.80, 1.20)

            # Quantidade = base x sazonalidade x variação, arredondada.
            # max(1, ...) garante que nenhuma venda tenha quantidade zero.
            quantidade = max(1, int(round(quantidade_base * fator_mes * variacao)))

            registros.append(
                {
                    "produto": produto,
                    "categoria": categoria,
                    "mes": mes,
                    "quantidade": quantidade,
                    "preco_unitario": preco,
                    "regiao": str(regiao),
                }
            )

    return registros


def salvar_planilha(registros):
    """Converte os registros em DataFrame e grava o arquivo .xlsx."""
    # Cria a tabela garantindo a ordem exata das colunas do enunciado.
    df = pd.DataFrame(registros, columns=COLUNAS)

    # Grava em .xlsx (o Pandas usa o openpyxl por baixo dos panos).
    # index=False evita criar uma coluna extra com o número da linha.
    # sheet_name define o nome da aba dentro da planilha.
    df.to_excel(NOME_ARQUIVO, index=False, sheet_name="vendas", engine="openpyxl")

    print(f"Planilha criada: {NOME_ARQUIVO}")
    print(f"Linhas de dados (sem o cabeçalho): {len(df)}")
    print(f"Colunas: {', '.join(df.columns)}")


# Este bloco só roda quando o arquivo é executado diretamente
# (python gerar_planilha.py), e não quando é importado por outro script.
if __name__ == "__main__":
    salvar_planilha(gerar_registros())
