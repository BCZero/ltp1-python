# Roteiro de Apresentação: Vídeo da Atividade 2BI

**Aluno:** Bruno Cesar Sampaio Ribeiro de Assis · **Tema 1:** Vendas de uma loja
**Duração-alvo:** 6 minutos (o professor sugere de 4 a 7)
**Regra de ouro:** fale com **suas palavras e sua voz**. O professor tira ponto de quem só mostra a tela. Para cada coisa que aparecer, diga **o que é, por que fez e o que significa**.

> As falas abaixo são um **guia**, não um texto para ler. Use os tópicos e explique do seu jeito. Leitura robótica soa como texto pronto e pesa contra no critério de autoria.

---

## Parte 0: Preparação (antes de apertar REC)

### Checklist técnico
- [ ] Rodou `python analise_vendas.py --auto` hoje e não deu erro.
- [ ] A pasta `graficos/` tem os 5 PNGs.
- [ ] Abertos e prontos: **VS Code** (pasta do projeto) e **Excel** (o `.xlsx`).
- [ ] No VS Code: terminal limpo (`cls` ou `clear`) e fonte grande (`Ctrl` + `+` duas ou três vezes).
- [ ] Notificações desligadas (WhatsApp, e-mail, Teams).
- [ ] Microfone testado. Grave 10 segundos e escute.
- [ ] Gravador escolhido: **OBS Studio** (grátis), **Xbox Game Bar** (`Win + Alt + R`, grava uma janela) ou **Clipchamp/Ferramenta de Captura** do Windows 11.

### Se você mudou a planilha
Se você editou valores no Excel (recomendado, para a base ser sua), rode de novo o modo `--auto` e **atualize os números deste roteiro** com a saída nova. As conclusões do programa se ajustam sozinhas; as falas deste roteiro não.

---

## Roteiro cronometrado

Os itens seguem a ordem do **roteiro obrigatório do enunciado** (são 11 itens; todos estão cobertos abaixo).

### ⏱ 0:00 – 0:20 · Item 1: Apresentação
**Tela:** pasta do projeto no VS Code (o nome do `.xlsx` aparece no explorador).

> "Olá, professor. Meu nome é **Bruno Cesar Sampaio Ribeiro de Assis**. Meu tema é o **Tema 1, Vendas de uma loja**, e o arquivo é o **brunoCesarSampaioRibeiroDeAssis.xlsx**. Criei a base de uma loja fictícia de informática e eletrônicos, a BS Tech Store, com vendas de janeiro a junho em todas as regiões do Brasil."

---

### ⏱ 0:20 – 1:05 · Item 2: Colunas da planilha
**Tela:** Excel com a planilha aberta.

Explique cada coluna, apontando com o mouse:
- **produto**: o que foi vendido, por exemplo Notebook e Cabo USB-C (são 10 produtos).
- **categoria**: o grupo do produto (Informática, Celulares, Periféricos e Acessórios).
- **mes**: o mês da venda, de Janeiro a Junho, guardado como texto.
- **quantidade**: quantas unidades foram vendidas (número inteiro).
- **preco_unitario**: o preço de uma unidade, em reais (decimal).
- **regiao**: a região do Brasil onde a venda aconteceu.

> Explique a lógica da base: "Montei os dados para serem coerentes. O preço de cada produto é fixo e parecido com o de mercado, e produtos baratos como cabos vendem muito mais unidades que notebooks. Coloquei sazonalidade: fevereiro com volta às aulas, maio com Dia das Mães e junho com Dia dos Namorados."

---

### ⏱ 1:05 – 1:20 · Item 3: Mínimo de 30 linhas
**Tela:** Excel. Clique na última linha (`Ctrl + ↓`): ela é a **linha 61**.

> "A planilha tem **60 linhas de dados**, da linha 2 à 61, mais o cabeçalho. É o dobro do mínimo: 10 produtos em 6 meses. O próprio programa também confere isso, como vou mostrar."

---

### ⏱ 1:20 – 2:00 · Item 4: Leitura com Pandas e `head()`
**Tela:** VS Code, função `carregar_dados()`. Depois rode `python analise_vendas.py`.

- Mostre a linha `pd.read_excel(nome_arquivo, engine="openpyxl")`.
> "Aqui o Pandas lê o xlsx com `read_excel`. O `engine="openpyxl"` indica a biblioteca que interpreta o formato .xlsx."
- Rode o programa e mostre a saída **Validação** + **Visão geral**:
> "Antes de analisar, o programa confere se o cabeçalho é exatamente o do Tema 1 e se tem pelo menos 30 linhas; deu **60, OK**. Em seguida o `head()` mostra as 5 primeiras linhas, para eu conferir se a leitura funcionou. O `dtypes` confirma os tipos: quantidade é inteiro e preço é decimal."

---

### ⏱ 2:00 – 2:35 · Item 5: Coluna calculada
**Tela:** função `preparar_dados()`, linha do `faturamento`, e depois a saída "COLUNA CALCULADA CRIADA".

> "A coluna obrigatória é **faturamento = quantidade × preço unitário**. O Pandas faz a conta linha a linha de uma vez só. Por exemplo, na primeira linha foram 13 notebooks a R$ 3.499,90, o que dá R$ 45.498,70."

Mencione também (vale ponto de entendimento):
> "Aqui também transformei o mês em uma **categoria ordenada**. Sem isso, o Pandas ordenaria os meses em ordem alfabética: Abril, Fevereiro, Janeiro..."

---

### ⏱ 2:35 – 3:10 · Item 6: `groupby()`
**Tela:** função `calcular_indicadores()`; depois digite `1` no menu.

> "O `groupby` junta as linhas que têm o mesmo valor. Por exemplo, `groupby("categoria")["faturamento"].sum()` junta todas as vendas de cada categoria e soma o faturamento. O `sort_values` ordena do maior para o menor."

Leia os resultados da opção 1:
- Faturamento total: **R$ 844.061,20**
- Produto campeão: **Notebook, R$ 262.492,50**
- Categoria campeã: **Informática, R$ 379.778,60**
- Melhor mês **Maio** (R$ 180.901,50), pior **Março** (R$ 116.164,50)
- Região líder: **Sul, R$ 256.958,20**

> Se sobrar tempo: "Também usei `agg` para tirar várias estatísticas de uma vez por categoria: unidades, faturamento e número de registros."

---

### ⏱ 3:10 – 3:45 · Item 7: Filtro
**Tela:** digite `2` no menu (vendas acima da média) e depois `7` → `Centro-Oeste`.

> "O filtro `df[df["faturamento"] > media]` mantém só as linhas em que a condição é verdadeira. A venda média foi de R$ 14.067,69. **Só 14 das 60 vendas (23%) ficam acima da média, mas elas somam 61,7% do faturamento**, e são apenas Notebook, Smartphone e Monitor. Poucas vendas de alto valor sustentam a loja."

> No filtro por região: "Aqui eu digito a região. O programa **primeiro filtra e depois agrupa** por categoria. No Centro-Oeste, onde eu moro, Celulares lideram." *(Se digitar errado, mostre que ele avisa.)*

---

### ⏱ 3:45 – 5:15 · Itens 8 e 9: os três gráficos e o porquê de cada um
Mostre cada um pelo menu (`3`, `4`, `5`). Para cada gráfico: **o que mostra → por que esse tipo → o que eu concluo**.

**Gráfico 1: Barras, faturamento por categoria** (`3`)
> "Escolhi **barras** porque o objetivo é **comparar grupos diferentes**, e a altura da barra deixa a diferença clara. Informática está bem à frente, Celulares em segundo; Acessórios ficam bem abaixo."

**Gráfico 2: Linha, faturamento por mês** (`4`)
> "Escolhi **linha** porque os meses têm uma **ordem no tempo**, e a linha mostra a tendência de subida e descida. A linha pontilhada é a média mensal. Fevereiro sobe com a volta às aulas, março é o vale e **maio é o pico**, com o Dia das Mães."

**Gráfico 3: Barras, faturamento por região** (`5`)
> "O enunciado deixava escolher pizza ou barras. Escolhi **barras** porque são 5 regiões e algumas têm valores próximos, então comparar fatias de pizza seria difícil. Para não perder a ideia de 'parte do todo', coloquei o **percentual em cima de cada barra**."

**Bônus (opcional, se o tempo permitir):** gráfico por produto (`6`), em barras horizontais porque os nomes são longos.

Mostre também a pasta `graficos/` com os PNGs:
> "Todos os gráficos são salvos em PNG com `fig.savefig`, com `dpi=300` e `bbox_inches="tight"` para não cortar o texto."

---

### ⏱ 5:15 – 5:55 · Item 10: Conclusões (mínimo 2; apresente 3)
**Tela:** opção `8` do menu.

1. **Concentração:** "Informática e Celulares somam **80% do faturamento**. A loja depende de produtos caros, e qualquer problema de estoque de notebook ou smartphone afeta muito o resultado."
2. **Volume não é faturamento:** "O produto que mais vende em unidades é o **Cabo USB-C (503 unidades)**, mas ele faturou só R$ 20 mil. O Notebook vendeu 75 unidades e faturou R$ 262 mil."
3. **Sazonalidade:** "Maio faturou **55,7% a mais** que março. A loja deveria reforçar estoque e campanhas antes de datas comemorativas."
4. *(Se quiser uma quarta)* **Regiões:** "O Sul liderou, à frente do Sudeste. Olhando os dados, o Sul teve **mais registros de venda (18)** e concentrou vendas grandes de notebook e smartphone. Ou seja, o resultado depende de quais produtos cada região compra, não só do tamanho da região."

---

### ⏱ 5:55 – 6:20 · Item 11: Dificuldade ou aprendizado
Escolha **uma** opção (de preferência uma que você realmente viveu):
- **Ordem dos meses:** "Minha principal dificuldade foi o gráfico de linha. Os meses saíam em ordem alfabética, porque o Pandas trata texto como texto. Resolvi transformando a coluna em categoria ordenada com `pd.Categorical`."
- **Filtrar antes de agrupar:** "Aprendi que a ordem importa: primeiro filtra, depois agrupa. Se agrupar antes, perco as linhas individuais."
- **Ambiente:** "Tive um `ModuleNotFoundError` porque instalei o pandas em outro Python. Aprendi a usar `python -m pip` e a escolher o interpretador no VS Code."

**Encerramento:**
> "O principal aprendizado é que o Pandas organiza os dados e o Matplotlib comunica o resultado. O gráfico só tem valor quando responde a uma pergunta. Obrigado, professor."

---

## Perguntas que o professor pode fazer (prepare as respostas)

| Pergunta | Resposta curta |
|---|---|
| Por que `engine="openpyxl"`? | É a biblioteca que o Pandas usa para ler e gravar o formato .xlsx |
| Diferença entre `fig` e `ax`? | `fig` é a figura inteira (tamanho, salvar); `ax` é a área do gráfico (dados, título, eixos) |
| O que `groupby` faz? | Junta linhas com o mesmo valor em uma coluna e aplica uma conta (soma, média) em cada grupo |
| O que `idxmax` devolve? | O **nome** (rótulo) do maior valor; o `max` devolve o valor |
| Por que `observed=True` no groupby do mês? | Para considerar só os meses que existem na base (o mês é categórico) |
| Por que `dpi=300`? | Resolução boa para impressão e documentos |
| Para que serve `bbox_inches="tight"`? | Remove bordas sobrando e evita que títulos e rótulos sejam cortados |
| Como os dados foram criados? | Com o script `gerar_planilha.py`: defini produtos, preços, sazonalidade e pesos das regiões, e o NumPy sorteou a região e uma variação de ±20% na quantidade, com semente fixa *(explique com suas palavras; se você editou a planilha à mão, conte isso também)* |
| Por que não pizza na região? | 5 fatias com valores próximos são difíceis de comparar; barras comparam melhor e têm eixos nomeados |
| O que acontece se a região não existir? | O filtro gera um DataFrame vazio; o programa testa `.empty` e mostra as regiões válidas |

---

## Ensaio

1. **1º ensaio:** com o roteiro aberto, cronometrando. Anote onde passou do tempo.
2. **2º ensaio:** só com os tópicos (sem ler as falas). Se passar de 7 min, corte o bônus do gráfico por produto e a 4ª conclusão.
3. **Gravação:** se errar, respire e continue. Pequenos tropeços mostram que é você falando.
4. **Revisão:** assista uma vez. Confira se o áudio está audível e se o nome do arquivo aparece.

## Checklist de entrega no Moodle
- [ ] `brunoCesarSampaioRibeiroDeAssis.xlsx`
- [ ] Pelo menos um PNG (sugestão: enviar os 5 da pasta `graficos/`)
- [ ] Vídeo de 4 a 7 minutos, com sua voz
- [ ] (Opcional) `analise_vendas.py`, caso o professor aceite o código junto
