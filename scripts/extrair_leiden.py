"""Extrai das planilhas do Leiden Ranking só as universidades do Brasil.

Uso: python extrair_leiden.py "<arquivo.xlsx>" <edicao>

A aba Results tem uma linha por universidade x área x período x método de contagem
(cerca de 190 mil linhas e 85 colunas), então a leitura é feita em fluxo.
Mantém: país Brasil, área "All sciences", contagem inteira (Frac_counting == 0),
todos os períodos. Os nomes das colunas variam entre edições (P_top10 / P_top_10,
impact_P / P), por isso são normalizados.
"""
import csv
import os
import re
import sys

import openpyxl

ARQ, EDICAO = sys.argv[1], sys.argv[2]
SAIDA = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "..", "dados", "processado", f"leiden_br_{EDICAO}.csv"
)

# nome normalizado (minúsculo, sem "_") -> nome de saída
COLUNAS = {
    "p": "P",
    "impactp": "P",
    "ptop10": "P_top10",
    "pptop10": "PP_top10",
    "pptop1": "PP_top1",
    "mncs": "MNCS",
    "mcs": "MCS",
    "ppcollab": "PP_collab",
    "ppintcollab": "PP_int_collab",
    "ppindustrycollab": "PP_industry_collab",
    "ppoa": "PP_OA",
    "tcs": "TCS",
    "tncs": "TNCS",
}


def norm(nome):
    return re.sub(r"[_\s]", "", str(nome)).lower()


os.makedirs(os.path.dirname(SAIDA), exist_ok=True)
wb = openpyxl.load_workbook(ARQ, read_only=True)
ws = wb["Results"]
linhas = ws.iter_rows(values_only=True)
cab = [str(c) for c in next(linhas)]
idx = {c: i for i, c in enumerate(cab)}
for obrig in ("University", "Country", "Field", "Period", "Frac_counting"):
    assert obrig in idx, f"coluna {obrig} ausente em {ARQ}"

# índice de saída -> posição na planilha (a primeira coluna que casar vence)
alvo = {}
for i, c in enumerate(cab):
    saida = COLUNAS.get(norm(c))
    if saida and saida not in alvo:
        alvo[saida] = i
ordem = ["University", "Period", "Edicao"] + list(alvo.keys())

n_lidas = n_ok = 0
with open(SAIDA, "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(ordem)
    for r in linhas:
        n_lidas += 1
        if r[idx["Country"]] != "Brazil" or r[idx["Field"]] != "All sciences":
            continue
        if r[idx["Frac_counting"]] not in (0, "0"):
            continue
        w.writerow([r[idx["University"]], r[idx["Period"]], EDICAO] + [r[i] for i in alvo.values()])
        n_ok += 1
print(f"{EDICAO}: {n_lidas} linhas lidas, {n_ok} mantidas -> {os.path.abspath(SAIDA)}")
print("colunas:", ordem)
