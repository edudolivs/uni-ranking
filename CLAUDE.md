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
# 2. Agrega ARWU + IGC + Enade + CSVs do Leiden em dados/processado/dados_dashboard.json
.venv/bin/python scripts/preparar_dados.py
# 3. Embute o JSON no template e grava dashboard_rankings.html
python3 scripts/gerar_dashboard.py
```

`preparar_dados.py` imprime contagens de controle. Com os dados atuais, os valores esperados são:
- **ARWU**, brasileiras por ano: 21, 21, 18, 18, 18, 14 (2021–2026).
- **Leiden**, brasileiras por edição: 35, 37, 38, 39 (2022–2025).
- **IGC**, linhas por ano: 2.090 (2019), 2.012 (2021), 2.023 (2022), 2.101 (2023).
- **Enade**, cursos por ano: 7.997, 9.896, 9.812 (2021–2023).

## Arquitetura do pipeline

`extrair_leiden.py` → `dados/processado/leiden_br_<ed>.csv` → `preparar_dados.py` → `dados_dashboard.json` → `gerar_dashboard.py` + `dashboard_template.html` → `dashboard_rankings.html`

- **Leiden:** cada xlsx tem 200 a 260 MB e a aba `Results` descompacta para cerca de 1 GB de XML. Leia sempre com `openpyxl` em `read_only` e em fluxo, nunca com pandas inteiro. O filtro mantém Brasil, `Field == "All sciences"`, `Frac_counting == 0` e todos os períodos. **Os nomes das colunas mudam entre edições** (`P_top10` em 2022 e `P_top_10` em 2025, `impact_P` e `P`), por isso `extrair_leiden.py` normaliza removendo `_` e usando minúsculas. `PP_top10` é fração (0,09 = 9%).
- **ARWU** (`dados/arwu/arwu_<ano>.json`): JSON da API pública `https://www.shanghairanking.com/api/pub/v1/arwu/rank?version=<ano>`, com 1.000 universidades por ano. A posição só é exata até a 100ª; depois vem em faixas ("801-900"), e `ranking_min` guarda o limite inferior para ordenar. `indData` traz códigos de indicador (165 a 170) **sem rótulo**: não atribua significado sem confirmar na fonte.
- **IGC** (`dados/igc/`, Inep): a primeira aba muda de nome por ano; os cabeçalhos têm espaços e `*` (limpos no código). Há **linhas de rodapé com notas**, filtradas por `ano_valido()`. Não existe edição 2020. A URL é `https://download.inep.gov.br/educacao_superior/indicadores/resultados/<ano>/IGC_<ano>.xlsx` (em 2022 é `igc_2022.xlsx`).
- **Enade** (`dados/enade/conceito_enade_<ano>.xlsx`): uma linha por curso; a aba se chama `PLANILHA_ENADE` ou `Planilha1` conforme o ano. Cada ano avalia áreas diferentes, então totais entre anos não indicam crescimento.
- **Cruzamento entre bases:** só por nome, via `chave()` em `preparar_dados.py` (sem acentos, sem palavras como "universidade/university/de", `TRAD` traduz state/estadual etc.). Equivalências que os tokens não resolvem vão em `ALIAS`; hoje só Unesp. **USP, Unicamp, Unesp e UERJ não aparecem no IGC/Enade.** A hipótese, não verificada, é que pertencem a sistemas estaduais.
- **Dashboard:** o template usa Chart.js por CDN (precisa de internet ao abrir) e cores da skill `dataviz`, já validadas para daltonismo, nos temas claro e escuro. A UFPB é destacada em laranja (`ehUfpb`). Cada gráfico tem uma tabela-espelho em `<details>`.

## O documento (`projeto-de-pesquisa.docx`)

- Segue o modelo da UFPB: Arial 12, margens 3/2/3/2 cm, capa, folha de rosto, aprovação, resumo (150–500 palavras, 3 palavras-chave), sumário e as seções 1–7 (Tema e problema, Hipóteses, Objetivos, Justificativa, Referencial teórico, Metodologia, Cronograma), mais as Referências (ABNT NBR 6023, ordem alfabética, linha em branco entre elas).
- **O script que gerou o docx não existe mais.** Ele ficava num diretório temporário da sessão e se perdeu. Para alterar o documento, edite o próprio `.docx` com `python-docx`, trocando o texto dos *runs* e clonando (`deepcopy`) parágrafos do mesmo tipo para manter a formatação.
- **O sumário é texto estático**, com pontilhado e números de página **estimados**, que precisam ser atualizados à mão. O cronograma é a tabela do modelo (colunas Mar–Jul) e alimenta as 11 atividades numeradas da Metodologia.
- **Referências:** são 10. Toda citação no texto precisa ter referência e vice-versa. Não invente referências nem dados bibliográficos; confirme DOI, volume e páginas (por exemplo, via `api.crossref.org/works/<doi>`).
- Edite o documento só quando pedido e não converta para PDF (preferência do usuário).

## Estado atual

**Feito**
- Projeto escrito no formato do modelo, já enxugado (justificativa, referencial teórico e metodologia mais curtos; 15 → 10 referências).
- Bases baixadas: ARWU 2021–2026, Leiden 2022–2025, IGC 2019 e 2021–2023, Enade 2021–2023.
- Pipeline e dashboard inicial (`dashboard_rankings.html`). O JavaScript foi testado sem erros num ambiente simulado, sem navegador, e **ainda não foi conferido visualmente**.

**Pendente / em aberto**
- **Faltam QS, SCImago e RUF.** O QS e o SCImago bloqueiam download automático (403); o RUF não tem dados abertos. Download manual: SCImago pelo botão "Download data" em scimagoir.com; QS em topuniversities.com. Não contorne bloqueios.
- **O RUF parece não ter edições de 2020 a 2022.** Isso conflita com as "cinco edições anuais consecutivas (ex.: 2021–2025)" da Metodologia; confirme e ajuste o texto ou a janela temporal.
- **Leiden 2019–2021** não foi baixado (os DOIs no Zenodo estão listados em traditional.leidenranking.com/downloads). **IGC 2024/2025:** o Inep mostra uma aba de 2025, mas o link do arquivo não foi encontrado.
- **No documento:**
  - Os meses do cronograma ainda são os do modelo.
  - O co-orientador está em branco.
  - Falta confirmar que Selten et al. (2020) usam a correlação de Spearman.
  - Os pesos do RUF vêm do Jornal da USP, não do site da Folha.
- **Análise:** ainda não há correlações nem comparação entre rankings; o dashboard é só descritivo.

## Git

O `.gitignore` exclui `dados/leiden/`, porque os xlsx têm 200–260 MB e o GitHub recusa arquivos acima de 100 MB. Num clone novo, baixe as planilhas de novo pelo Zenodo antes de rodar `extrair_leiden.py`; os CSVs com o recorte do Brasil já estão versionados em `dados/processado/`. Também ficam fora os logs da extração, `__pycache__/` e `.venv/`. As demais bases (ARWU, IGC, Enade, cerca de 12 MB) estão no repositório.
