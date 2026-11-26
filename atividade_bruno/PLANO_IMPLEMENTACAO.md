# Plano de Implementação: Atividade Avaliativa 2BI (Tema 1)

**Aluno:** Bruno Cesar Sampaio Ribeiro de Assis
**Tema:** 1, Vendas de uma loja
**Entrega:** `brunoCesarSampaioRibeiroDeAssis.xlsx` + pelo menos 1 PNG + vídeo de 4 a 7 min no Moodle

---

## 1. Entendimento da atividade

| Item | Exigência | Como vamos atender |
|---|---|---|
| Planilha | `.xlsx` próprio, ≥ 30 linhas, cabeçalho exato, nome em camelCase | 60 linhas, cabeçalho `produto,categoria,mes,quantidade,preco_unitario,regiao` |
| Pandas | `read_excel`, `head`, coluna calculada, `groupby`, filtro | Todos presentes, cada um em uma função própria |
| Matplotlib | 3 gráficos com título e eixos; 1 PNG com `fig.savefig` | 3 obrigatórios + 2 extras, todos salvos em PNG |
| Interpretação | Conclusões coerentes | 4 conclusões calculadas a partir dos números |
| Vídeo | Voz do aluno, roteiro de 11 itens | Ver `ROTEIRO_APRESENTACAO.md` |

**Critérios (10 pontos):** xlsx (2) · Pandas (2) · Matplotlib (2) · Interpretação (2) · Vídeo (2).
**Penalidades a evitar:** menos de 30 linhas, nome errado do arquivo, gráfico sem título ou sem eixos, vídeo sem voz, mostrar a tela sem explicar, base igual à de colegas.

---

## 2. Fases de execução

| Fase | Atividade | Saída | Status |
|---|---|---|---|
| 1 | Ler o enunciado e a aula de Pandas/Matplotlib | Lista de requisitos (seção 1) | ✅ |
| 2 | Desenhar a base de dados (contexto, produtos, preços, sazonalidade) | Catálogo em `gerar_planilha.py` | ✅ |
| 3 | Gerar e revisar a planilha | `brunoCesarSampaioRibeiroDeAssis.xlsx` | ✅ |
| 4 | Implementar leitura, validação e preparação | `carregar_dados`, `validar_estrutura`, `preparar_dados` | ✅ |
| 5 | Implementar análises (`groupby`) e filtros | `calcular_indicadores`, `filtrar_*` | ✅ |
| 6 | Implementar gráficos e exportação PNG | `gerar_*`, `salvar_e_exibir` | ✅ |
| 7 | Implementar conclusões automáticas e menu | `gerar_conclusoes`, `exibir_menu` | ✅ |
| 8 | Testar (modo `--auto` e menu) | PNGs em `graficos/` | ✅ |
| 9 | Documentar (README, requirements) | `README.md`, `requirements.txt` | ✅ |
| 10 | **Revisar a planilha no Excel e personalizar** | Base com a sua cara | ⏳ Bruno |
| 11 | **Ensaiar e gravar o vídeo** | Vídeo de 4 a 7 min | ⏳ Bruno |
| 12 | **Enviar no Moodle** | xlsx + PNG + vídeo | ⏳ Bruno |

---

## 3. Desenho da base de dados

**Contexto:** "BS Tech Store", loja de informática e eletrônicos que vende em todas as regiões do Brasil.

| Categoria | Produtos (preço) |
|---|---|
| Informática | Notebook (R$ 3.499,90), Monitor 24 polegadas (R$ 899,90), Impressora Multifuncional (R$ 749,90) |
| Celulares | Smartphone (R$ 1.899,90), Smartwatch (R$ 699,90) |
| Periféricos | Teclado Mecânico (R$ 299,90), Mouse Gamer (R$ 159,90), Headset (R$ 249,90) |
| Acessórios | Carregador Turbo (R$ 89,90), Cabo USB-C (R$ 39,90) |

**Regras de coerência:**
- 10 produtos × 6 meses = **60 linhas** (o dobro do mínimo).
- Preço fixo por produto; a quantidade varia.
- Itens baratos vendem mais unidades (cabo > notebook).
- Sazonalidade: fevereiro sobe (volta às aulas), março cai, **maio é o pico** (Dia das Mães), junho sobe de novo (Dia dos Namorados).
- Região sorteada com pesos (Sudeste com mais chance), e variação aleatória de ±20% na quantidade.
- Semente fixa (`2026`): a planilha pode ser recriada igual.

---

## 4. Arquitetura do código (`analise_vendas.py`)

O programa segue o fluxo da aula: **o Pandas prepara e o Matplotlib apresenta**. Cada etapa fica em uma função.

```
main()
 ├─ carregar_dados()             -> pd.read_excel
 ├─ validar_estrutura()          -> cabeçalho exato + ≥ 30 linhas
 ├─ mostrar_visao_geral()        -> head(), shape, dtypes
 ├─ preparar_dados()             -> limpeza + mes ordenado + faturamento
 ├─ calcular_indicadores()       -> 5 groupby + agg (dicionário de resultados)
 └─ exibir_menu()  ou  --auto    -> executar_relatorio_completo()
        ├─ exibir_indicadores()
        ├─ filtrar_acima_da_media()      (filtro 1)
        ├─ filtrar_categoria_no_mes()    (filtro 3, com &)
        ├─ gerar_barras_categoria()      (gráfico 1, obrigatório)
        ├─ gerar_linha_mensal()          (gráfico 2, obrigatório)
        ├─ gerar_barras_regiao()         (gráfico 3, obrigatório)
        ├─ gerar_barras_produto()        (gráfico 4, extra)
        ├─ gerar_barras_regiao_escolhida() -> filtrar_por_regiao() (filtro 2 + gráfico 5)
        └─ gerar_conclusoes()
```

### Decisões técnicas (e por quê)

| Decisão | Motivo |
|---|---|
| `mes` vira `pd.Categorical(..., ordered=True)` | Sem isso o `groupby` ordena os meses em ordem alfabética e o gráfico de linha fica errado |
| Gráfico de região em **barras**, não pizza | 5 regiões com valores próximos são mais fáceis de comparar em barras; barras têm eixos nomeados (critério de avaliação). O percentual aparece em cima de cada barra |
| `calcular_indicadores` devolve um dicionário | Cada conta é feita uma vez e reaproveitada nos gráficos e nas conclusões |
| Conclusões montadas com f-strings | Se a planilha mudar, o texto se atualiza sozinho |
| Função `salvar_e_exibir` única | Todos os gráficos são salvos do mesmo jeito (`dpi=300`, `bbox_inches="tight"`) |
| Menu com `input()` | Permite mostrar cada parte separadamente no vídeo |
| Modo `--auto` | Testa tudo rápido, sem abrir janelas |
| Filtro antes do `groupby` no gráfico por região | Filtrar depois de agrupar perderia as linhas individuais (como explicado na aula) |
| Comparação de texto com `.str.lower()` | O usuário pode digitar "sul", "SUL" ou "Sul" |

---

## 5. Plano de testes

| Teste | Como | Resultado esperado |
|---|---|---|
| Execução completa | `python analise_vendas.py --auto` | 5 PNGs em `graficos/`, sem erros |
| Menu: opção válida | Digitar `3` | Abre o gráfico de categorias |
| Menu: opção inválida | Digitar `abc` | "Opção inválida..." e o menu volta |
| Filtro com minúsculas | Opção `7`, digitar `sul` | Gráfico da região Sul |
| Região inexistente | Opção `7`, digitar `Marte` | Mensagem com as regiões válidas |
| Arquivo ausente | Renomear o xlsx | Mensagem clara, sem traceback |
| Ordem dos meses | Ver o gráfico de linha | Janeiro → Junho |

Todos os testes acima foram executados e passaram.

---

## 6. Próximos passos do Bruno

1. Abrir o `.xlsx` no Excel e conferir as 60 linhas.
2. **(Recomendado para autoria)** Ajustar à mão alguns valores de `quantidade` ou trocar a região de algumas linhas. O código recalcula tudo e as conclusões se atualizam sozinhas. **Depois anote os novos números** para usar no roteiro.
3. Rodar `python analise_vendas.py --auto` e conferir os gráficos.
4. Estudar o `ROTEIRO_APRESENTACAO.md` e ensaiar 2 vezes com cronômetro.
5. Gravar, revisar o áudio e enviar no Moodle: xlsx + `01_faturamento_por_categoria.png` (ou todos os PNGs) + vídeo.
