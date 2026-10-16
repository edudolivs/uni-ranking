"""Extrai do SCImago Institutions Rankings as universidades do Brasil.

Entrada: dados/scimago/ScimagoIR <ano> - Overall Rank - BRA.csv (já vem
filtrado por Brasil na origem; filtramos aqui só Sector == "Universities",
porque o SIR cobre também hospitais, governo, ONGs e empresas).
Saída: dados/processado/scimago_br.csv

Separador ';'. 2022-2025 têm 5 colunas (sem "Best Country Quartile"); 2026
acrescenta a 6a. Coluna ausente no ano fica nula.
"""
import csv
import glob
import os
import re

RAIZ = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "dados", "scimago")
SAIDA = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "dados", "processado", "scimago_br.csv")
CAMPOS = ["ano", "nome", "rank_br", "rank_global", "quartil"]


def main():
    todas = []
    arqs = sorted(glob.glob(os.path.join(RAIZ, "ScimagoIR * - Overall Rank - BRA.csv")))
    for caminho in arqs:
        ano = int(re.search(r"ScimagoIR (\d{4})", os.path.basename(caminho)).group(1))
        with open(caminho, encoding="utf-8", newline="") as f:
            leitor = list(csv.DictReader(f, delimiter=";"))
        br = [r for r in leitor if r.get("Sector") == "Universities"]
        print(f"ScimagoIR {ano}: {len(leitor)} linhas (todos os setores), {len(br)} universidades")
        for r in br:
            nome = r["Institution"].replace(" *", "").strip()  # "*" marca nota de rodapé do SIR
            todas.append({
                "ano": ano, "nome": nome, "rank_br": int(r["Rank"]), "rank_global": int(r["Global Rank"]),
                "quartil": r.get("Best Country Quartile") or None,
            })
    todas.sort(key=lambda r: (r["ano"], r["rank_br"]))
    os.makedirs(os.path.dirname(SAIDA), exist_ok=True)
    with open(SAIDA, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=CAMPOS)
        w.writeheader()
        w.writerows(todas)
    print(f"salvo {os.path.abspath(SAIDA)}, {len(todas)} linhas")
    por_ano = {}
    for r in todas:
        por_ano[r["ano"]] = por_ano.get(r["ano"], 0) + 1
    print("SCImago universidades brasileiras por ano:", dict(sorted(por_ano.items())))


if __name__ == "__main__":
    main()
