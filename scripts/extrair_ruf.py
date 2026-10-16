"""Extrai o Ranking Universitário Folha (RUF).

Entrada: dados/ruf/ruf-<ano>.csv (';' como separador).
Saída: dados/processado/ruf.csv

2019 NÃO é o RUF geral: a coluna "Ranking" do arquivo de 2019 está ordenada
exatamente pela "Avaliação do Mercado" (mesma ordem, inclusive os empates) --
ou seja, é a sub-classificação de Mercado, não a posição geral do RUF. Por
isso "ranking" (posição geral) fica nulo em 2019; só pos_mercado/nota_mercado
são preenchidos. 2023-2025 trazem os 5 indicadores completos e a posição geral.
"""
import csv
import glob
import os
import re

RAIZ = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "dados", "ruf")
SAIDA = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "dados", "processado", "ruf.csv")
CAMPOS = ["ano", "nome", "ranking", "categoria", "nota_geral",
          "pos_ensino", "nota_ensino", "pos_pesquisa", "nota_pesquisa",
          "pos_mercado", "nota_mercado", "pos_inovacao", "nota_inovacao",
          "pos_internacionalizacao", "nota_internacionalizacao"]


def num(x):
    if x is None:
        return None
    s = str(x).strip()
    if s in ("", "-"):
        return None
    try:
        return float(s.replace(",", "."))
    except ValueError:
        return None


def ler_2019(caminho):
    with open(caminho, encoding="utf-8-sig", newline="") as f:
        linhas = list(csv.DictReader(f, delimiter=";"))
    saida = []
    for r in linhas:
        saida.append({
            # "Ranking" do arquivo de 2019 é a posição NA SUB-CLASSIFICAÇÃO DE MERCADO
            # (ordenação idêntica à de "Avaliação do Mercado", empates inclusive) --
            # não é a posição geral do RUF, que não está disponível para 2019.
            "ano": 2019, "nome": r["Nome"].strip(), "ranking": None,
            "categoria": r["Pública ou Privada"], "nota_geral": None,
            "pos_ensino": None, "nota_ensino": None, "pos_pesquisa": None, "nota_pesquisa": None,
            "pos_mercado": int(r["Ranking"]), "nota_mercado": num(r.get("Avaliação do Mercado")),
            "pos_inovacao": None, "nota_inovacao": None,
            "pos_internacionalizacao": None, "nota_internacionalizacao": None,
        })
    return linhas, saida


def ler_completo(caminho, ano):
    with open(caminho, encoding="utf-8-sig", newline="") as f:
        linhas = list(csv.DictReader(f, delimiter=";"))
    saida = []
    for r in linhas:
        saida.append({
            "ano": ano, "nome": r["Universidade"].strip(), "ranking": int(r["Ranking"]),
            "categoria": r["Pública ou Privada"], "nota_geral": num(r.get("Nota")),
            "pos_ensino": num(r.get("Posição em Ensino")), "nota_ensino": num(r.get("Nota em Ensino")),
            "pos_pesquisa": num(r.get("Posição em Pesquisa")), "nota_pesquisa": num(r.get("Nota em Pesquisa")),
            "pos_mercado": num(r.get("Posição em Mercado")), "nota_mercado": num(r.get("Nota em Mercado")),
            "pos_inovacao": num(r.get("Posição em Inovação")), "nota_inovacao": num(r.get("Nota em Inovação")),
            "pos_internacionalizacao": num(r.get("Posição em Internacionalização")),
            "nota_internacionalizacao": num(r.get("Nota em Internacionalização")),
        })
    return linhas, saida


def main():
    todas = []
    for caminho in sorted(glob.glob(os.path.join(RAIZ, "ruf-*.csv"))):
        ano = int(re.search(r"ruf-(\d{4})", os.path.basename(caminho)).group(1))
        linhas, saida = ler_2019(caminho) if ano == 2019 else ler_completo(caminho, ano)
        print(f"ruf-{ano}.csv: {len(linhas)} linhas")
        todas.extend(saida)
    todas.sort(key=lambda r: (r["ano"], r["ranking"] if r["ranking"] is not None else r["pos_mercado"] or 10**9))
    os.makedirs(os.path.dirname(SAIDA), exist_ok=True)
    with open(SAIDA, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=CAMPOS)
        w.writeheader()
        w.writerows(todas)
    print(f"salvo {os.path.abspath(SAIDA)}, {len(todas)} linhas")
    por_ano = {}
    for r in todas:
        por_ano[r["ano"]] = por_ano.get(r["ano"], 0) + 1
    print("RUF universidades por ano:", dict(sorted(por_ano.items())))


if __name__ == "__main__":
    main()
