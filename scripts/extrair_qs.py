"""Extrai das planilhas do QS World University Rankings só as universidades do Brasil.

Entrada: dados/qs/*.csv (um arquivo por edição, exceto o combinado 2017-2022).
Saída: dados/processado/qs_br.csv

O QS mudou o nome e a presença das colunas de indicador em quase toda edição
(ex. "Citations per Faculty Score" em 2024, "CPF SCORE" em 2026, ausente em
2017-2022). Por isso, ao contrário do extrair_leiden.py (que normaliza nomes
de coluna automaticamente), aqui cada arquivo tem um mapa de colunas explícito
-- os esquemas são estruturalmente diferentes demais pra normalizar por regra.

Indicadores mantidos quando existem no arquivo: Academic Reputation (AR),
Employer Reputation (ER), Faculty/Student Ratio (FSR), Citations per Faculty
(CPF). O arquivo 2017-2022 (formato longo, uma linha por universidade-ano) não
tem esse detalhe por indicador; fica só com a posição e a nota geral.

Nenhum valor é inventado: indicador ausente no esquema do ano fica nulo.
"""
import csv
import os
import re

RAIZ = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "dados", "qs")
SAIDA = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "dados", "processado", "qs_br.csv")

CAMPOS = ["ano", "nome", "rank_raw", "rank_min", "score_geral",
          "ar_score", "ar_rank", "er_score", "er_rank", "fsr_score", "fsr_rank", "cpf_score", "cpf_rank"]


def num(x):
    if x is None:
        return None
    s = str(x).strip()
    if s in ("", "-", "n/a", "N/A"):
        return None
    try:
        return float(s.replace(",", "."))
    except ValueError:
        return None


def rank_min(x):
    """'92' -> 92; '611-620' -> 611; vazio/'-' -> None."""
    if x is None:
        return None
    m = re.match(r"\d+", str(x).strip())
    return int(m.group()) if m else None


def abrir(nome_arquivo, encoding):
    caminho = os.path.join(RAIZ, nome_arquivo)
    return open(caminho, encoding=encoding, newline="")


# cada entrada: arquivo, encoding, ano (None = multi-ano, usa coluna 'ano_col'),
# e o mapa campo-de-saida -> nome-da-coluna-no-csv (None = não existe nesse ano)
EDICOES = [
    {
        "arquivo": "qs-world-university-rankings-2017-to-2022-V2.csv", "encoding": "utf-8-sig",
        "multi_ano": True, "col_ano": "year", "col_nome": "university", "col_pais": "country",
        "col_rank": "rank_display", "col_score": "score",
        "col_ar_s": None, "col_ar_r": None, "col_er_s": None, "col_er_r": None,
        "col_fsr_s": None, "col_fsr_r": None, "col_cpf_s": None, "col_cpf_r": None,
    },
    {
        "arquivo": "2023 QS World University Rankings.csv", "encoding": "utf-8-sig", "ano": 2023,
        "col_nome": "institution", "col_pais": "location", "col_rank": "Rank", "col_score": "score scaled",
        "col_ar_s": "ar score", "col_ar_r": "ar rank", "col_er_s": "er score", "col_er_r": "er rank",
        "col_fsr_s": "fsr score", "col_fsr_r": "fsr rank", "col_cpf_s": "cpf score", "col_cpf_r": "cpf rank",
    },
    {
        "arquivo": "2024 QS World University Rankings 1.1 (For qs.com).csv", "encoding": "utf-8-sig", "ano": 2024,
        "col_nome": "Institution Name", "col_pais": "Country", "col_rank": "2024 RANK", "col_score": "Overall SCORE",
        "col_ar_s": "Academic Reputation Score", "col_ar_r": "Academic Reputation Rank",
        "col_er_s": "Employer Reputation Score", "col_er_r": "Employer Reputation Rank",
        "col_fsr_s": "Faculty Student Score", "col_fsr_r": "Faculty Student Rank",
        "col_cpf_s": "Citations per Faculty Score", "col_cpf_r": "Citations per Faculty Rank",
    },
    {
        "arquivo": "QS World University Rankings 2025 (Top global universities).csv", "encoding": "cp1252", "ano": 2025,
        "col_nome": "Institution_Name", "col_pais": "Location", "col_rank": "RANK_2025", "col_score": "Overall_Score",
        "col_ar_s": "Academic_Reputation_Score", "col_ar_r": "Academic_Reputation_Rank",
        "col_er_s": "Employer_Reputation_Score", "col_er_r": "Employer_Reputation_Rank",
        "col_fsr_s": "Faculty_Student_Score", "col_fsr_r": "Faculty_Student_Rank",
        "col_cpf_s": "Citations_per_Faculty_Score", "col_cpf_r": "Citations_per_Faculty_Rank",
    },
    {
        "arquivo": "2026 QS World University Rankings.csv", "encoding": "utf-8-sig", "ano": 2026,
        "col_nome": "Institution Name", "col_pais": "Country/Territory", "col_rank": "2026 Rank", "col_score": "Overall SCORE",
        "col_ar_s": "AR SCORE", "col_ar_r": "AR RANK", "col_er_s": "ER SCORE", "col_er_r": "ER RANK",
        "col_fsr_s": "FSR SCORE", "col_fsr_r": "FSR RANK", "col_cpf_s": "CPF SCORE", "col_cpf_r": "CPF RANK",
    },
    {
        "arquivo": "2027 QS World University Rankings.csv", "encoding": "utf-8-sig", "ano": 2027,
        "col_nome": "Name", "col_pais": "Country/Territory", "col_rank": "Rank", "col_score": "Overall SCORE",
        "col_ar_s": "AR SCORE", "col_ar_r": "AR RANK", "col_er_s": "ER SCORE", "col_er_r": "ER RANK",
        "col_fsr_s": "FSR SCORE", "col_fsr_r": "FSR RANK", "col_cpf_s": "CPF SCORE", "col_cpf_r": "CPF RANK",
    },
]


def ler_arquivo(cfg):
    linhas_lidas = 0
    saida = []
    with abrir(cfg["arquivo"], cfg["encoding"]) as f:
        leitor = csv.DictReader(f)
        # remove BOM residual do nome da 1a coluna, se sobrou
        leitor.fieldnames = [c.lstrip("﻿") for c in leitor.fieldnames]
        for r in leitor:
            linhas_lidas += 1
            pais = (r.get(cfg["col_pais"]) or "").strip()
            if pais != "Brazil":
                continue
            ano = int(r[cfg["col_ano"]]) if cfg.get("multi_ano") else cfg["ano"]
            rk_raw = r.get(cfg["col_rank"])
            saida.append({
                "ano": ano, "nome": (r.get(cfg["col_nome"]) or "").strip(),
                "rank_raw": rk_raw, "rank_min": rank_min(rk_raw), "score_geral": num(r.get(cfg["col_score"])),
                "ar_score": num(r.get(cfg["col_ar_s"])) if cfg["col_ar_s"] else None,
                "ar_rank": num(r.get(cfg["col_ar_r"])) if cfg["col_ar_r"] else None,
                "er_score": num(r.get(cfg["col_er_s"])) if cfg["col_er_s"] else None,
                "er_rank": num(r.get(cfg["col_er_r"])) if cfg["col_er_r"] else None,
                "fsr_score": num(r.get(cfg["col_fsr_s"])) if cfg["col_fsr_s"] else None,
                "fsr_rank": num(r.get(cfg["col_fsr_r"])) if cfg["col_fsr_r"] else None,
                "cpf_score": num(r.get(cfg["col_cpf_s"])) if cfg["col_cpf_s"] else None,
                "cpf_rank": num(r.get(cfg["col_cpf_r"])) if cfg["col_cpf_r"] else None,
            })
    return linhas_lidas, saida


def main():
    todas = []
    for cfg in EDICOES:
        caminho = os.path.join(RAIZ, cfg["arquivo"])
        if not os.path.exists(caminho):
            print(f"AVISO: {cfg['arquivo']} não encontrado, pulando")
            continue
        lidas, br = ler_arquivo(cfg)
        anos = sorted({r["ano"] for r in br})
        print(f"{cfg['arquivo']}: {lidas} linhas lidas, {len(br)} do Brasil, anos {anos}")
        todas.extend(br)

    todas.sort(key=lambda r: (r["ano"], r["rank_min"] if r["rank_min"] is not None else 10**9))
    os.makedirs(os.path.dirname(SAIDA), exist_ok=True)
    with open(SAIDA, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=CAMPOS)
        w.writeheader()
        w.writerows(todas)
    print(f"salvo {os.path.abspath(SAIDA)}, {len(todas)} linhas")
    por_ano = {}
    for r in todas:
        por_ano[r["ano"]] = por_ano.get(r["ano"], 0) + 1
    print("QS brasileiras por ano:", dict(sorted(por_ano.items())))


if __name__ == "__main__":
    main()
