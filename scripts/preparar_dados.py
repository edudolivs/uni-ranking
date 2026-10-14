"""Prepara os dados agregados que alimentam o dashboard inicial.

Entradas (em dados/):
  arwu/arwu_<ano>.json            API pública do ShanghaiRanking
  igc/*.xlsx                      Índice Geral de Cursos (Inep)
  enade/conceito_enade_<ano>.xlsx Conceito Enade por curso (Inep)
  processado/leiden_br_<ed>.csv   gerado por extrair_leiden.py
Saída: dados/processado/dados_dashboard.json

Nenhum valor é inventado: o que não existir na fonte fica nulo.
"""
import csv
import glob
import json
import os
import re
import unicodedata
from collections import Counter, defaultdict

import openpyxl

RAIZ = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "dados")
SAIDA = os.path.join(RAIZ, "processado", "dados_dashboard.json")


def num(x):
    """Converte para float; devolve None para vazio, '-', 'SC' etc."""
    if x is None:
        return None
    try:
        return float(str(x).replace(",", "."))
    except ValueError:
        return None


def ano_valido(x):
    """Ano inteiro entre 2000 e 2100; None para notas de rodapé e linhas vazias."""
    try:
        v = int(float(str(x).strip()))
    except (TypeError, ValueError):
        return None
    return v if 2000 <= v <= 2100 else None


def faixa(x):
    if x is None:
        return "SC"
    s = str(x).strip()
    if s in ("", "-"):
        return "SC"
    try:
        return str(int(float(s)))
    except ValueError:
        return "SC"


# ---------------------------------------------------------------- chave de nome
STOP = {
    "university", "universidade", "universidad", "of", "de", "do", "da", "dos", "das", "e", "the",
    "fundacao", "foundation", "instituto", "institute", "centro", "center", "centre", "faculdade",
}
TRAD = {"state": "estadual", "catholic": "catolica", "pontifical": "pontificia", "pontificial": "pontificia",
        "federal": "federal", "estadual": "estadual", "catolica": "catolica", "pontificia": "pontificia"}


# Mesma instituição com nomes sem equivalência por tokens (inglês x português)
ALIAS = {"estadual paulo sao": "estadual paulista"}  # Sao Paulo State University = Unesp


def titulo(nome):
    """Caixa de título em português: 'UNIVERSIDADE FEDERAL DA PARAÍBA' -> 'Universidade Federal da Paraíba'."""
    miudas = {"de", "da", "do", "das", "dos", "e"}
    palavras = str(nome).lower().split()
    return " ".join(w if (i and w in miudas) else w.capitalize() for i, w in enumerate(palavras))


def chave(nome):
    s = unicodedata.normalize("NFKD", str(nome)).encode("ascii", "ignore").decode().lower()
    toks = [TRAD.get(t, t) for t in re.findall(r"[a-z]+", s)]
    toks = [t for t in toks if t not in STOP]
    k = " ".join(sorted(set(toks)))
    return ALIAS.get(k, k)


# ---------------------------------------------------------------- ARWU
def ler_arwu():
    anos, br = {}, []
    for f in sorted(glob.glob(os.path.join(RAIZ, "arwu", "arwu_*.json"))):
        ano = int(re.search(r"(\d{4})", os.path.basename(f)).group(1))
        lista = json.load(open(f, encoding="utf-8"))["data"]["rankings"]
        brasil = [r for r in lista if r.get("region") == "Brazil"]
        paises = Counter(r.get("region") for r in lista)
        anos[ano] = {"total": len(lista), "brasil": len(brasil), "paises": paises.most_common(8)}
        for r in brasil:
            rk = str(r["ranking"])
            br.append({
                "ano": ano,
                "nome": r["univNameEn"],
                "ranking": rk,
                "ranking_min": int(re.match(r"\d+", rk).group()),
                "regional": r.get("regionRanking"),
                "score": num(r.get("score")),
                "ind": r.get("indData"),
            })
    return anos, br


# ---------------------------------------------------------------- IGC
def ler_igc():
    por_ano, regs = {}, []
    arqs = sorted(glob.glob(os.path.join(RAIZ, "igc", "*.xlsx")))
    for f in arqs:
        wb = openpyxl.load_workbook(f, read_only=True)
        ws = wb[wb.sheetnames[0]]
        it = ws.iter_rows(values_only=True)
        cab = [str(c).strip().replace("*", "").strip() if c else "" for c in next(it)]
        ix = {c: i for i, c in enumerate(cab)}
        cont = Counter()
        ano = None
        for r in it:
            a = ano_valido(r[ix["Ano"]])
            if a is None:
                continue
            ano = a
            fx = faixa(r[ix["IGC (Faixa)"]])
            cat = r[ix["Categoria Administrativa"]] or "n/d"
            org = r[ix["Organização Acadêmica"]] or "n/d"
            cont[(fx, cat)] += 1
            regs.append({
                "ano": ano, "nome": str(r[ix["Nome da IES"]]).strip(), "sigla": r[ix["Sigla da IES"]],
                "uf": r[ix["Sigla da UF"]], "org": org, "cat": cat,
                "igc": num(r[ix["IGC (Contínuo)"]]), "faixa": fx,
                "cursos_cpc": num(r[ix["Nº de Cursos com CPC no triênio"]]),
                "conceito_grad": num(r[ix["Conceito Médio de Graduação"]]),
                "conceito_mest": num(r[ix["Conceito Médio de Mestrado"]]),
                "conceito_dout": num(r[ix["Conceito Médio do doutorado"]]),
            })
        por_ano[ano] = [{"faixa": k[0], "cat": k[1], "n": v} for k, v in sorted(cont.items())]
    return por_ano, regs


# ---------------------------------------------------------------- ENADE
def ler_enade():
    resumo = defaultdict(lambda: {"cursos": 0, "ies": set(), "soma": 0.0, "n": 0})
    por_ies = defaultdict(lambda: {"cursos": 0, "soma": 0.0, "n": 0, "nome": None, "cat": None, "org": None})
    arqs = sorted(glob.glob(os.path.join(RAIZ, "enade", "conceito_enade_*.xlsx")))
    for f in arqs:
        wb = openpyxl.load_workbook(f, read_only=True)
        aba = next(n for n in wb.sheetnames if n.upper().startswith("PLANILHA") and n != "Planilha2")
        it = wb[aba].iter_rows(values_only=True)
        cab = [str(c).strip().replace("*", "").strip() if c else "" for c in next(it)]
        ix = {c: i for i, c in enumerate(cab)}
        for r in it:
            ano = ano_valido(r[ix["Ano"]])
            if ano is None:
                continue
            cat = r[ix["Categoria Administrativa"]] or "n/d"
            cod = r[ix["Código da IES"]]
            v = num(r[ix["Conceito Enade (Contínuo)"]])
            g = resumo[(ano, cat)]
            g["cursos"] += 1
            g["ies"].add(cod)
            if v is not None:
                g["soma"] += v
                g["n"] += 1
            p = por_ies[(ano, cod)]
            p["nome"], p["cat"], p["org"] = str(r[ix["Nome da IES"]]).strip(), cat, r[ix["Organização Acadêmica"]]
            p["cursos"] += 1
            if v is not None:
                p["soma"] += v
                p["n"] += 1
    res = [{"ano": a, "cat": c, "cursos": g["cursos"], "ies": len(g["ies"]),
            "media": round(g["soma"] / g["n"], 3) if g["n"] else None}
           for (a, c), g in sorted(resumo.items())]
    ies = [{"ano": a, "cod": c, "nome": p["nome"], "cat": p["cat"], "org": p["org"], "cursos": p["cursos"],
            "media": round(p["soma"] / p["n"], 3) if p["n"] else None}
           for (a, c), p in sorted(por_ies.items(), key=lambda kv: (kv[0][0], str(kv[0][1])))]
    return res, ies


# ---------------------------------------------------------------- Leiden
def ler_leiden():
    regs = []
    for f in sorted(glob.glob(os.path.join(RAIZ, "processado", "leiden_br_*.csv"))):
        for r in csv.DictReader(open(f, encoding="utf-8")):
            regs.append({
                "edicao": int(r["Edicao"]), "nome": r["University"], "periodo": r["Period"],
                "P": num(r.get("P")), "P_top10": num(r.get("P_top10")), "PP_top10": num(r.get("PP_top10")),
                "MNCS": num(r.get("MNCS")), "PP_int_collab": num(r.get("PP_int_collab")),
                "PP_industry_collab": num(r.get("PP_industry_collab")), "PP_OA": num(r.get("PP_OA")),
            })
    return regs


# ---------------------------------------------------------------- cruzamento
def cruzar(arwu_br, igc, enade_ies, leiden):
    """Une por chave de nome as universidades brasileiras da edição mais recente de cada base."""
    ult = lambda lst, campo: max(r[campo] for r in lst)
    a_ano, i_ano, e_ano, l_ed = (ult(arwu_br, "ano"), ult(igc, "ano"), ult(enade_ies, "ano"),
                                 ult(leiden, "edicao") if leiden else None)
    linhas = defaultdict(dict)
    nomes = {}
    for r in arwu_br:
        if r["ano"] == a_ano:
            k = chave(r["nome"]); linhas[k]["arwu"] = r; nomes.setdefault(k, r["nome"])
    for r in igc:
        if r["ano"] == i_ano and r["org"] == "Universidade":
            k = chave(r["nome"]); linhas[k]["igc"] = r; nomes[k] = titulo(r["nome"])
    for r in enade_ies:
        if r["ano"] == e_ano and r["org"] == "Universidade":
            k = chave(r["nome"]); linhas[k]["enade"] = r; nomes.setdefault(k, titulo(r["nome"]))
    if leiden:
        ult_periodo = max(r["periodo"] for r in leiden if r["edicao"] == l_ed)
        for r in leiden:
            if r["edicao"] == l_ed and r["periodo"] == ult_periodo:
                k = chave(r["nome"]); linhas[k]["leiden"] = r; nomes.setdefault(k, r["nome"])
    saida = []
    for k, d in linhas.items():
        if len(d) < 2:
            continue
        saida.append({
            "nome": nomes[k], "fontes": sorted(d.keys()),
            "arwu": d["arwu"]["ranking"] if "arwu" in d else None,
            "arwu_min": d["arwu"]["ranking_min"] if "arwu" in d else None,
            "leiden_pp": d["leiden"]["PP_top10"] if "leiden" in d else None,
            "leiden_p": d["leiden"]["P"] if "leiden" in d else None,
            "igc": d["igc"]["igc"] if "igc" in d else None,
            "igc_faixa": d["igc"]["faixa"] if "igc" in d else None,
            "enade": d["enade"]["media"] if "enade" in d else None,
        })
    saida.sort(key=lambda r: (-len(r["fontes"]), r["arwu_min"] if r["arwu_min"] else 99999, r["nome"]))
    cobertura = {
        "arwu": sum(1 for d in linhas.values() if "arwu" in d),
        "igc": sum(1 for d in linhas.values() if "igc" in d),
        "enade": sum(1 for d in linhas.values() if "enade" in d),
        "leiden": sum(1 for d in linhas.values() if "leiden" in d),
        "em_2_ou_mais": len(saida),
        "em_todas": sum(1 for r in saida if len(r["fontes"]) == 4),
        "anos": {"arwu": a_ano, "igc": i_ano, "enade": e_ano, "leiden": l_ed},
    }
    return saida, cobertura


def main():
    arwu_anos, arwu_br = ler_arwu()
    igc_ano, igc = ler_igc()
    enade_res, enade_ies = ler_enade()
    leiden = ler_leiden()
    cruz, cob = cruzar(arwu_br, igc, enade_ies, leiden)
    saida = {
        "arwu": {"anos": arwu_anos, "brasil": arwu_br},
        "igc": {"por_ano": igc_ano, "universidades": [r for r in igc if r["org"] == "Universidade"]},
        "enade": {"resumo": enade_res},
        "leiden": leiden,
        "cruzamento": cruz,
        "cobertura": cob,
    }
    with open(SAIDA, "w", encoding="utf-8") as f:
        json.dump(saida, f, ensure_ascii=False, separators=(",", ":"))
    print("salvo", os.path.abspath(SAIDA), round(os.path.getsize(SAIDA) / 1024), "KB")
    print("ARWU brasileiras por ano:", {a: v["brasil"] for a, v in arwu_anos.items()})
    print("IGC linhas por ano:", {a: sum(x["n"] for x in v) for a, v in igc_ano.items()})
    print("ENADE cursos por ano:", {a: sum(r["cursos"] for r in enade_res if r["ano"] == a) for a in sorted({r['ano'] for r in enade_res})})
    print("Leiden linhas:", len(leiden), "| cobertura do cruzamento:", cob)


if __name__ == "__main__":
    main()
