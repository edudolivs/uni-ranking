# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## O projeto

Projeto de pesquisa da disciplina **Metodologia do Trabalho Científico** (Bacharelado em Ciência de Dados e IA, Centro de Informática, UFPB), orientado pela Profa. Elizabet Maria Spohr de Medeiros. Grupo: Ana Beatriz Mota da Cruz, Eduardo de Oliveira Silva e Nicole Costa e Silva. Todo o conteúdo (documento, dashboard, código e comentários) é em **português do Brasil**.

**Mudança de tema.** A apresentação de 18/09/2026 (`pesquisa_aplicada_comprimido.pdf`) tratava dos *determinantes do desempenho de universidades em rankings internacionais*. O projeto escrito passou a **comparar os critérios de avaliação de rankings brasileiros e internacionais**:

- **Pergunta:** quais características das universidades mais pesam na posição em rankings brasileiros e em internacionais, e como esses pesos diferem ao longo de, no mínimo, cinco anos?
- **Método:** análise exploratória de dados (EDA) longitudinal, com correlação de Spearman por ranking e por ano e comparação entre os dois grupos. Os resultados são lidos como associações, não como causas.
- **Rankings internacionais:** QS, ARWU, Leiden e SCImago. **Brasileiros:** RUF (Folha), IGC/CPC (Inep) e rankings de pesquisa com recorte brasileiro.
- **Hipóteses:** (a) nos internacionais, pesquisa e citações se associam mais à posição; (b) nos brasileiros, pesam mais ensino, opinião de docentes e empregadores e condições de oferta; (c) os padrões são estáveis no tempo, mas diferentes entre os grupos.
- **Amostra principal:** universidades brasileiras presentes em pelo menos um ranking de cada grupo.

## Arquivos principais

| Arquivo | O que é |
|---|---|
| `pesquisa_aplicada_comprimido.pdf` | Slides da apresentação original (tema antigo). Traz o dicionário de 15 variáveis que o grupo planejou. |
| `grupo-2-versao-1-projeto-de-pesquisa.docx` | Modelo oficial da UFPB. **Não alterar.** |
| `projeto-de-pesquisa.docx` | Projeto escrito no formato exato do modelo. Este é o entregável. |
| `dashboard_rankings.html` | Dashboard gerado (não editar à mão; editar o template e regerar). |
| `scripts/` | Pipeline de dados e do dashboard. |
| `dados/` | Bases brutas por fonte e `dados/processado/` com os derivados. |

## Comandos

Não há `requirements.txt`, testes nem linter. Os scripts precisam de `openpyxl`, que o Python do sistema não tem. Editar o docx exige `python-docx`. Use um venv:

```bash
python3 -m venv .venv && .venv/bin/pip install openpyxl python-docx

# 1. Leiden: extrai só o Brasil das planilhas grandes (uma edição por chamada; segundos cada)
.venv/bin/python scripts/extrair_leiden.py "dados/leiden/CWTS Leiden Ranking 2024.xlsx" 2024
# 2. QS, SCImago, RUF: extraem dados/qs|scimago|ruf/ para dados/processado/
.venv/bin/python scripts/extrair_qs.py
.venv/bin/python scripts/extrair_scimago.py
.venv/bin/python scripts/extrair_ruf.py
# 3. Agrega ARWU + IGC + CPC + Enade + Leiden + QS + SCImago + RUF em dados/processado/dados_dashboard.json
#    (lê dados/cpc/*.xlsx direto, sem extrator próprio -- mesma granularidade do Enade)
.venv/bin/python scripts/preparar_dados.py
# 4. Embute o JSON no template e grava dashboard_rankings.html
python3 scripts/gerar_dashboard.py
```

`preparar_dados.py` imprime contagens de controle. Com os dados atuais, os valores esperados são:
- **ARWU**, brasileiras por ano: 21, 21, 18, 18, 18, 14 (2021–2026).
- **Leiden**, brasileiras por edição: ~35 a 39 (2021–2025; 2021 entrou depois, confira o log antes de assumir o número).
- **IGC**, linhas por ano: 2.090 (2019), 2.012 (2021), 2.023 (2022), 2.101 (2023).
- **CPC**, cursos por ano: 8.368 (2019), 7.997 (2021), 9.891 (2022), 9.812 (2023) -- próximo do Enade, mas não idêntico (escopos levemente diferentes).
- **Enade**, cursos por ano: 7.997, 9.896, 9.812 (2021–2023).
- **QS**, brasileiras por ano: 22, 22, 22, 20, 22, 27 (2017–2022), depois 35, 35, 35, 24, 22 (2023–2027) -- a queda em 2026/2027 é real na fonte, não é bug.
- **SCImago**, universidades (setor "Universities") por ano: 151, 170, 166, 165, 171 (2022–2026).
- **RUF**: 192 (2019, só sub-ranking de Mercado -- ver nota abaixo), ~203-204 por ano (2023–2025).

## Arquitetura do pipeline

`extrair_leiden.py` / `extrair_qs.py` / `extrair_scimago.py` / `extrair_ruf.py` → `dados/processado/*.csv` → `preparar_dados.py` (lê também `dados/cpc/*.xlsx` direto) → `dados_dashboard.json` → `gerar_dashboard.py` + `dashboard_template.html` → `dashboard_rankings.html`

- **Leiden:** cada xlsx tem 170 a 262 MB e a aba `Results` descompacta para cerca de 1 GB de XML. Leia sempre com `openpyxl` em `read_only` e em fluxo, nunca com pandas inteiro. O filtro mantém Brasil, `Field == "All sciences"`, `Frac_counting == 0` e todos os períodos. **Os nomes das colunas mudam entre edições** (`P_top10` em 2022 e `P_top_10` em 2025, `impact_P` e `P`), por isso `extrair_leiden.py` normaliza removendo `_` e usando minúsculas. `PP_top10` é fração (0,09 = 9%). Os `.xlsx` brutos não são versionados (ver `## Git`); só os CSVs já filtrados em `dados/processado/`.
- **ARWU** (`dados/arwu/arwu_<ano>.json`): JSON da API pública `https://www.shanghairanking.com/api/pub/v1/arwu/rank?version=<ano>`, com 1.000 universidades por ano. A posição só é exata até a 100ª; depois vem em faixas ("801-900"), e `ranking_min` guarda o limite inferior para ordenar. `indData` traz códigos de indicador (165 a 170) **sem rótulo**: não atribua significado sem confirmar na fonte.
- **QS** (`dados/qs/*.csv`, `extrair_qs.py`): **6 arquivos, 6 esquemas de coluna diferentes** (2017-2022 é um único CSV longo multi-ano sem indicadores por categoria; 2023 a 2027 são wide, cada um com nomenclatura própria para reputação acadêmica/empregador, proporção docente-aluno, citações por docente, docentes/alunos internacionais, rede internacional de pesquisa, empregabilidade -- `ger score` em 2023 -- e sustentabilidade, só desde 2024). Encoding varia (`cp1252`/`utf-8-sig`) -- cada arquivo tem o encoding certo hardcoded no script, não adivinhe. Indicador ausente na edição fica nulo.
- **SCImago** (`dados/scimago/*.csv`, `extrair_scimago.py`): já vem filtrado por Brasil na origem; o script filtra `Sector == "Universities"` (o SIR também cobre hospitais, governo, ONGs e empresas). Separador `;`.
- **RUF** (`dados/ruf/*.csv`, `extrair_ruf.py`): **o arquivo de 2019 NÃO é a posição geral do RUF** -- é a sub-classificação de Mercado (a coluna "Ranking" do CSV original está ordenada exatamente pela nota de mercado, empates inclusive). `extrair_ruf.py` já corrige isso: `ranking` fica nulo em 2019, só `pos_mercado`/`nota_mercado` são preenchidos. 2023-2025 trazem os 5 indicadores completos (Ensino, Pesquisa, Mercado, Inovação, Internacionalização) e a posição geral.
- **IGC** (`dados/igc/`, Inep): a primeira aba muda de nome por ano; os cabeçalhos têm espaços e `*` (limpos no código). Há **linhas de rodapé com notas**, filtradas por `ano_valido()`. Não existe edição 2020. A URL é `https://download.inep.gov.br/educacao_superior/indicadores/resultados/<ano>/IGC_<ano>.xlsx` (em 2022 é `igc_2022.xlsx`).
- **CPC** (`dados/cpc/*.xlsx`, Inep): Conceito Preliminar de Curso, um dos componentes do próprio IGC -- por isso tende a correlacionar fortemente com ele, não é uma fonte independente. Mesma estrutura do Enade (lido por `ler_cpc()` em `preparar_dados.py`, que espelha `ler_enade()`); nomes de aba e ordem de coluna variam por ano, tratados por dicionário de índice (não por posição fixa).
- **Enade** (`dados/enade/conceito_enade_<ano>.xlsx`): uma linha por curso; a aba se chama `PLANILHA_ENADE` ou `Planilha1` conforme o ano. Cada ano avalia áreas diferentes, então totais entre anos não indicam crescimento.
- **Cruzamento entre bases:** só por nome, via `chave()` em `preparar_dados.py` (sem acentos, sem palavras como "universidade/university/de", `TRAD` traduz state/estadual etc.). Equivalências que os tokens não resolvem vão em `ALIAS`; hoje só Unesp. **USP, Unicamp, Unesp e UERJ não aparecem no IGC/CPC/Enade.** A hipótese, não verificada, é que pertencem a sistemas estaduais. `cruzar()` é genérico para N fontes (hoje 8); cada fonte entra só com a edição/ano mais recente.
- **Correlação de Spearman** (`correlacionar()` em `preparar_dados.py`): implementada em Python puro (rank médio em empate + Pearson nos postos), sem dependência nova. Cobre estabilidade temporal (posição entre edições consecutivas), indicador de pesquisa×posição nos internacionais, nota média no IGC/CPC/Enade × posição no ARWU/QS/SCImago no mesmo ano (`inep_x_internacional`), e cada coluna de indicador do QS e do RUF × posição geral (`colunas_qs_ruf`; o dashboard mostra a edição mais recente, com sinal invertido para leitura, destacando a maior e a menor). `n` sempre reportado ao lado de `rho`; sem p-value forçado em amostra pequena. Mínimo de 5 universidades em comum para calcular.
- **Dashboard:** o template usa Chart.js por CDN (precisa de internet ao abrir) e cores da skill `dataviz`, já validadas para daltonismo, nos temas claro e escuro. **Atenção:** `--ord3` é idêntico a `--s1` (mesmo azul) -- não use `ord[2]` num gráfico que já usa `T.s1`, ou as séries ficam indistinguíveis (já aconteceu no gráfico de cobertura; corrigido usando `ord[0]`). A UFPB é destacada em laranja (`ehUfpb`). Cada gráfico tem uma tabela-espelho em `<details>`. O gráfico de cobertura (Visão geral) mostra só ARWU, Leiden e QS, nos anos com ARWU ou Leiden (SCImago fica fora por não ser um recorte de topo). 6 painéis por botão: Visão geral, Pesquisa e impacto, Ensino e avaliação, Comparar universidades, Correlações, Fontes e limites.

## O documento (`projeto-de-pesquisa.docx`)

- Segue o modelo da UFPB: Arial 12, margens 3/2/3/2 cm, capa, folha de rosto, aprovação, resumo (150–500 palavras, 3 palavras-chave), sumário e as seções 1–7 (Tema e problema, Hipóteses, Objetivos, Justificativa, Referencial teórico, Metodologia, Cronograma), mais as Referências (ABNT NBR 6023, ordem alfabética, linha em branco entre elas).
- **O script que gerou o docx não existe mais.** Ele ficava num diretório temporário da sessão e se perdeu. Para alterar o documento, edite o próprio `.docx` com `python-docx`, trocando o texto dos *runs* e clonando (`deepcopy`) parágrafos do mesmo tipo para manter a formatação.
- **O sumário é texto estático**, com pontilhado e números de página **estimados**, que precisam ser atualizados à mão. O cronograma é a tabela do modelo (colunas Mar–Jul) e alimenta as 11 atividades numeradas da Metodologia.
- **Referências:** são 10. Toda citação no texto precisa ter referência e vice-versa. Não invente referências nem dados bibliográficos; confirme DOI, volume e páginas (por exemplo, via `api.crossref.org/works/<doi>`).
- Edite o documento só quando pedido e não converta para PDF (preferência do usuário).

## Estado atual

**Feito**
- Projeto escrito no formato do modelo, já enxugado (justificativa, referencial teórico e metodologia mais curtos; 15 → 10 referências).
- Bases baixadas: ARWU 2021–2026, Leiden 2021–2025, QS 2017–2027, SCImago 2022–2026, RUF 2019+2023–2025, IGC 2019 e 2021–2023, CPC 2019 e 2021–2023, Enade 2021–2023 -- as 8 fontes que a pergunta de pesquisa original previa (QS/ARWU/Leiden/SCImago internacionais; RUF/IGC-CPC/Enade brasileiros).
- Pipeline completo (`extrair_leiden.py`, `extrair_qs.py`, `extrair_scimago.py`, `extrair_ruf.py`, `preparar_dados.py`, `gerar_dashboard.py`) e dashboard (`dashboard_rankings.html`), 6 painéis por botão (estilo Power BI, sem rolagem). Testado com Chromium headless (Playwright): clique nos 6 painéis, zero erro de console, screenshots conferidos nos temas claro e escuro.
- **Correlação de Spearman implementada** (`correlacionar()` em `preparar_dados.py`, painel "Correlações" no dashboard): estabilidade temporal por ranking, indicador×posição nos internacionais, índices do Inep × posição internacional, maior e menor correlação entre as colunas do QS e do RUF. Python puro, sem dependência nova.
- Confirmado: Selten et al. (2020) de fato usa correlação de Spearman -- a citação no texto do `.docx` está correta.

**Pendente / em aberto**
- **Bases fechadas (decisão do grupo, 16/10/2026):** as edições já baixadas são todas as que serão usadas. Leiden 2019–2020, RUF 2020–2022 e IGC/CPC 2024–2025 **não** entram; o dashboard não mostra mais aviso de falta de dados.
- **RUF 2019 não é a posição geral**, é só a sub-classificação de Mercado (ver `## Arquitetura do pipeline`) -- isso ainda conflita com as "cinco edições anuais consecutivas" da Metodologia; confirme e ajuste o texto ou a janela temporal.
- **No documento:**
  - Os meses do cronograma ainda são os do modelo (Mar–Jul); a pesquisa já roda em set/out/2026 pelas datas de acesso das referências citadas -- precisa dos meses reais.
  - O co-orientador está em branco.
  - Os pesos do RUF ainda citam o Jornal da USP; agora há CSV oficial 2025 com os 5 indicadores, dá pra trocar por fonte primária.
  - A seção de limitações/resultados esperados ainda não reflete que a correlação existe (dizia "dashboard é só descritivo").

## Git

O `.gitignore` exclui `dados/leiden/`, porque os xlsx têm 170–262 MB e o GitHub recusa arquivos acima de 100 MB. Num clone novo, baixe as planilhas de novo pelo Zenodo antes de rodar `extrair_leiden.py`; os CSVs com o recorte do Brasil já estão versionados em `dados/processado/`. Também ficam fora os logs da extração, `__pycache__/` e `.venv/`. As demais bases (ARWU, QS, SCImago, RUF, IGC, CPC, Enade, poucas dezenas de MB no total) estão no repositório.
