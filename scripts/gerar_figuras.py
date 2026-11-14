"""Gera as figuras e os números do relatório de pesquisa.

Entrada: dados/processado/dados_dashboard.json (gerado por preparar_dados.py)
Saídas:  relatorio/figuras/fig<N>_*.png  (300 dpi, tema claro, para impressão)
         relatorio/resultados.json       (agregados citados no texto do relatório)

Nenhum número do relatório é digitado à mão: o texto lê resultados.json.
Convenção de sinal: nos rankings, posição menor = melhor; nos indicadores, nota maior = melhor.
Por isso o rho de Spearman entre indicador e posição sai negativo quando os dois concordam.
Para leitura, as figuras mostram a "concordância" = -rho (positivo = concordam).
"""
import json
import os
import sys
from collections import defaultdict
from statistics import mean

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.colors import LinearSegmentedColormap

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from preparar_dados import chave, serie_por_ano, spearman  # noqa: E402

RAIZ = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")
ENTRADA = os.path.join(RAIZ, "dados", "processado", "dados_dashboard.json")
DIR_FIG = os.path.join(RAIZ, "relatorio", "figuras")
SAIDA_RES = os.path.join(RAIZ, "relatorio", "resultados.json")

# paleta da skill dataviz (tema claro), a mesma do dashboard
SUP, INK, INK2, MUTED, GRID, AXIS = "#ffffff", "#0b0b0b", "#52514e", "#898781", "#e1e0d9", "#c3c2b7"
CAT = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300", "#6250d6", "#e34948"]
INTL, BR = CAT[0], CAT[1]          # grupo internacional = azul; brasileiro = laranja
ORD = ["#86b6ef", "#5598e7", "#2a78d6", "#1c5cab", "#104281"]  # ordinal (anos)
DIVERG = LinearSegmentedColormap.from_list("div", ["#e34948", "#f0efec", "#2a78d6"])

plt.rcParams.update({
    "font.family": "Arial", "font.size": 9, "axes.edgecolor": AXIS, "axes.labelcolor": INK2,
    "xtick.color": INK2, "ytick.color": INK2, "axes.spines.top": False, "axes.spines.right": False,
    "axes.grid": True, "grid.color": GRID, "grid.linewidth": 0.6, "axes.axisbelow": True,
    "figure.facecolor": SUP, "axes.facecolor": SUP, "savefig.facecolor": SUP, "legend.frameon": False,
})
LARG = 6.3  # polegadas (largura útil da página A4 com margens 3/2 cm)


def eh_ufpb(nome):
    k = chave(nome)
    return k == "federal paraiba"


def salvar(fig, nome):
    os.makedirs(DIR_FIG, exist_ok=True)
    caminho = os.path.join(DIR_FIG, nome)
    fig.savefig(caminho, dpi=300, bbox_inches="tight")
    plt.close(fig)
    print("figura:", os.path.relpath(caminho, RAIZ))


def virgula(x, casas=2):
    return f"{x:.{casas}f}".replace(".", ",")


def eixo_virgula(eixo):
    """Marcas do eixo com vírgula decimal (pt-BR)."""
    from matplotlib.ticker import FuncFormatter
    eixo.set_major_formatter(FuncFormatter(lambda v, _: f"{v:g}".replace(".", ",")))


# ---------------------------------------------------------------- Figura 1: cobertura
def fig_cobertura(d):
    series = {
        "ARWU": {int(a): v["brasil"] for a, v in d["arwu"]["anos"].items()},
        "QS": {},
        "Leiden": {},
        "SCImago": {},
        "RUF": {},
    }
    for nome, regs, campo_ano, campo_val in [("QS", d["qs"], "ano", "rank_min"),
                                             ("SCImago", d["scimago"], "ano", "rank_global"),
                                             ("RUF", d["ruf"], "ano", "ranking")]:
        cont = defaultdict(set)
        for r in regs:
            if r.get(campo_val) is not None:
                cont[r[campo_ano]].add(chave(r["nome"]))
        series[nome] = {a: len(s) for a, s in cont.items()}
    cont = defaultdict(set)
    for r in d["leiden"]:
        cont[r["edicao"]].add(chave(r["nome"]))
    series["Leiden"] = {a: len(s) for a, s in cont.items()}

    fig, ax = plt.subplots(figsize=(LARG, 3.1))
    marcadores = ["o", "s", "^", "D", "v"]
    for i, (nome, s) in enumerate(series.items()):
        anos = sorted(s)
        ys = [s[a] for a in anos]
        ax.plot(anos, ys, color=CAT[i], lw=2, marker=marcadores[i], ms=4.5, label=nome,
                markeredgecolor=SUP, markeredgewidth=0.8)
        ax.annotate(nome, (anos[-1], ys[-1]), xytext=(6, 0), textcoords="offset points",
                    va="center", fontsize=8, color=INK2)
    ax.set_yscale("log")
    ax.set_yticks([10, 20, 50, 100, 200])
    ax.set_yticklabels(["10", "20", "50", "100", "200"])
    ax.set_ylabel("Universidades brasileiras (escala log)")
    ax.set_xticks(range(2017, 2028))
    ax.set_xlim(2016.6, 2028.2)
    ax.legend(ncol=5, loc="upper center", bbox_to_anchor=(0.5, 1.13), fontsize=8)
    salvar(fig, "fig1_cobertura.png")
    return {k: {str(a): v for a, v in sorted(s.items())} for k, s in series.items()}


# ---------------------------------------------------------------- Figura 2: peso por indicador
DIMENSAO = {
    "Citações por docente": "Pesquisa", "Pesquisa": "Pesquisa", "Rede internacional de pesquisa": "Pesquisa",
    "Reputação acadêmica": "Reputação", "Reputação com empregadores": "Mercado/reputação",
    "Empregabilidade": "Mercado/reputação", "Mercado": "Mercado/reputação",
    "Proporção docente/aluno": "Ensino", "Ensino": "Ensino",
    "Docentes internacionais": "Internacionalização", "Alunos internacionais": "Internacionalização",
    "Internacionalização": "Internacionalização", "Inovação": "Inovação", "Sustentabilidade": "Sustentabilidade",
}
ORDEM_DIM = ["Pesquisa", "Ensino", "Mercado/reputação", "Reputação", "Internacionalização", "Inovação",
             "Sustentabilidade"]


def fig_peso_indicador(col):
    ult = {"QS": max(r["ano"] for r in col if r["ranking"] == "QS"),
           "RUF": max(r["ano"] for r in col if r["ranking"] == "RUF")}
    linhas = [r for r in col if r["ano"] == ult[r["ranking"]]]
    linhas.sort(key=lambda r: (ORDEM_DIM.index(DIMENSAO[r["coluna"]]), r["ranking"], r["rho"]))
    fig, ax = plt.subplots(figsize=(LARG, 4.4))
    ys = list(range(len(linhas)))[::-1]
    for y, r in zip(ys, linhas):
        cor = INTL if r["ranking"] == "QS" else BR
        ax.barh(y, -r["rho"], color=cor, height=0.62, edgecolor=SUP, linewidth=1)
        ax.text(-r["rho"] + 0.01, y, f"{virgula(-r['rho'])}  (n={r['n']})", va="center", fontsize=7.5, color=INK2)
    ax.set_yticks(ys)
    ax.set_yticklabels([f"{r['ranking']} {ult[r['ranking']]} · {r['coluna']}" for r in linhas], fontsize=8)
    ax.set_xlim(0, 1.42)
    ax.set_xticks([0, 0.2, 0.4, 0.6, 0.8, 1.0])
    eixo_virgula(ax.xaxis)
    ax.spines["bottom"].set_bounds(0, 1.0)
    ax.set_xlabel("Concordância entre indicador e posição geral (−ρ de Spearman)", x=0.36)
    ax.grid(axis="y", visible=False)
    # separadores de dimensão
    dims = [DIMENSAO[r["coluna"]] for r in linhas]
    for i in range(1, len(dims)):
        if dims[i] != dims[i - 1]:
            ax.axhline(ys[i] + 0.5, color=AXIS, lw=0.6, ls=(0, (2, 2)))
    inicio = 0
    for i in range(1, len(dims) + 1):
        if i == len(dims) or dims[i] != dims[inicio]:
            meio = (ys[inicio] + ys[i - 1]) / 2
            ax.text(1.41, meio, dims[inicio], va="center", ha="right", fontsize=7.5, color=MUTED, style="italic")
            inicio = i
    from matplotlib.patches import Patch
    ax.legend(handles=[Patch(color=INTL, label="QS (internacional)"), Patch(color=BR, label="RUF (brasileiro)")],
              loc="upper center", bbox_to_anchor=(0.45, 1.08), ncol=2, fontsize=8)
    salvar(fig, "fig2_peso_indicador.png")
    return ult


# ---------------------------------------------------------------- Figura 3: variação temporal do rho
def fig_variacao_temporal(col):
    por_ind = defaultdict(dict)
    for r in col:
        por_ind[(r["ranking"], r["coluna"])][r["ano"]] = -r["rho"]
    chaves = sorted(por_ind, key=lambda k: (k[0] != "RUF", -mean(por_ind[k].values())))
    anos = sorted({a for v in por_ind.values() for a in v})
    cor_ano = {a: ORD[i] for i, a in enumerate(anos[-len(ORD):])}
    fig, ax = plt.subplots(figsize=(LARG, 4.6))
    ys = list(range(len(chaves)))[::-1]
    for y, k in zip(ys, chaves):
        v = por_ind[k]
        ax.plot([min(v.values()), max(v.values())], [y, y], color=BR if k[0] == "RUF" else INTL, lw=2.2,
                alpha=0.35, solid_capstyle="round")
        for a, x in sorted(v.items()):
            ax.scatter(x, y, s=34, color=cor_ano[a], edgecolor=SUP, linewidth=0.8, zorder=3)
    ax.set_yticks(ys)
    ax.set_yticklabels([f"{k[0]} · {k[1]}" for k in chaves], fontsize=8)
    ax.set_xlim(0.2, 1.0)
    eixo_virgula(ax.xaxis)
    ax.set_xlabel("Concordância entre indicador e posição geral (−ρ), por edição")
    ax.grid(axis="y", visible=False)
    for a in anos:
        ax.scatter([], [], s=34, color=cor_ano[a], label=str(a), edgecolor=SUP)
    ax.legend(title="Edição", ncol=len(anos), loc="upper center", bbox_to_anchor=(0.5, 1.12), fontsize=8,
              title_fontsize=8)
    salvar(fig, "fig3_variacao_temporal.png")


# ---------------------------------------------------------------- Figura 4: estabilidade
def fig_estabilidade(est):
    por = defaultdict(list)
    for r in est:
        por[r["rotulo"].split(":")[0]].append(r)
    fig, ax = plt.subplots(figsize=(LARG, 2.9))
    marcadores = ["o", "s", "^", "D"]
    for i, nome in enumerate(["ARWU", "QS", "SCImago", "RUF"]):
        rs = sorted(por[nome], key=lambda r: r["para"])
        xs, ys = [r["para"] for r in rs], [r["rho"] for r in rs]
        cor = CAT[[0, 1, 2, 6][i]]
        ax.plot(xs, ys, color=cor, lw=2, marker=marcadores[i], ms=4.5, label=nome,
                markeredgecolor=SUP, markeredgewidth=0.8)
    ax.set_ylim(0.6, 1.01)
    eixo_virgula(ax.yaxis)
    ax.set_ylabel("ρ entre edições consecutivas")
    ax.set_xlabel("Edição (comparada com a anterior)")
    ax.set_xticks(range(2018, 2028))
    ax.set_xlim(2017.6, 2027.4)
    ax.legend(ncol=4, loc="lower left", fontsize=8)
    salvar(fig, "fig4_estabilidade.png")


# ---------------------------------------------------------------- Figura 5: Inep x internacional
def fig_inep_internacional(inep):
    linhas = [(i, r) for i in ["IGC", "CPC", "Enade"] for r in ["ARWU", "QS", "SCImago"]]
    anos = sorted({x["ano"] for x in inep})
    idx = {(x["indice"], x["ranking"], x["ano"]): x for x in inep}
    fig, ax = plt.subplots(figsize=(LARG, 3.4))
    mat = [[(-idx[(i, r, a)]["rho"]) if (i, r, a) in idx else float("nan") for a in anos] for i, r in linhas]
    im = ax.imshow(mat, cmap=DIVERG, vmin=-1, vmax=1, aspect="auto")
    for y, (i, r) in enumerate(linhas):
        for x, a in enumerate(anos):
            it = idx.get((i, r, a))
            if it:
                v = -it["rho"]
                ax.text(x, y, f"{virgula(v)}\nn={it['n']}", ha="center", va="center", fontsize=7,
                        color="#ffffff" if v > 0.6 else INK)
            else:
                ax.text(x, y, "–", ha="center", va="center", fontsize=8, color=MUTED)
    ax.set_xticks(range(len(anos)))
    ax.set_xticklabels(anos)
    ax.set_yticks(range(len(linhas)))
    ax.set_yticklabels([f"{i} × {r}" for i, r in linhas], fontsize=8)
    ax.grid(False)
    for s in ax.spines.values():
        s.set_visible(False)
    cb = fig.colorbar(im, ax=ax, fraction=0.04, pad=0.02)
    cb.set_label("Concordância (−ρ)", fontsize=8)
    eixo_virgula(cb.ax.yaxis)
    cb.outline.set_visible(False)
    salvar(fig, "fig5_inep_internacional.png")


# ---------------------------------------------------------------- Figura 6 e Tabela 2: concordância entre grupos
def concordancia_grupos(d):
    s_ruf = serie_por_ano(d["ruf"], "ano", "nome", "ranking")
    internacionais = {
        "SCImago": serie_por_ano(d["scimago"], "ano", "nome", "rank_global"),
        "QS": serie_por_ano(d["qs"], "ano", "nome", "rank_min"),
        "ARWU": serie_por_ano(d["arwu"]["brasil"], "ano", "nome", "ranking_min"),
    }
    s_igc = serie_por_ano(d["igc"]["universidades"], "ano", "nome", "igc", filtro=lambda r: r["org"] == "Universidade")
    tabela = []
    for nome, s in internacionais.items():
        for ano in sorted(set(s) & set(s_ruf)):
            comuns = set(s[ano]) & set(s_ruf[ano])
            rho, n = spearman([(s_ruf[ano][k], s[ano][k]) for k in comuns])
            if rho is not None:
                tabela.append({"par": f"RUF × {nome}", "ano": ano, "rho": round(rho, 3), "n": n})
    for ano in sorted(set(s_igc) & set(s_ruf)):
        comuns = set(s_igc[ano]) & set(s_ruf[ano])
        rho, n = spearman([(s_ruf[ano][k], s_igc[ano][k]) for k in comuns])
        if rho is not None:
            tabela.append({"par": "RUF × IGC", "ano": ano, "rho": round(rho, 3), "n": n})

    # dispersão RUF x SCImago no ano mais recente em comum
    ano = max(set(internacionais["SCImago"]) & set(s_ruf))
    nomes_ruf = {chave(r["nome"]): r["nome"] for r in d["ruf"] if r["ano"] == ano}
    s_sc = internacionais["SCImago"][ano]
    comuns = sorted(set(s_sc) & set(s_ruf[ano]))
    xs = [s_ruf[ano][k] for k in comuns]
    ys = [s_sc[k] for k in comuns]
    rho, n = spearman(list(zip(xs, ys)))
    fig, ax = plt.subplots(figsize=(LARG, 3.6))
    ax.scatter(xs, ys, s=16, color=INTL, alpha=0.75, edgecolor=SUP, linewidth=0.5,
               label="Universidade brasileira")
    for k, x, y in zip(comuns, xs, ys):
        if eh_ufpb(nomes_ruf.get(k, "")):
            ax.scatter([x], [y], s=60, color=BR, edgecolor=INK, linewidth=0.8, zorder=4, label="UFPB")
            ax.annotate(f"UFPB\nRUF {int(x)}º · SCImago {int(y)}º", (x, y), xytext=(14, -18),
                        textcoords="offset points", fontsize=7.5, color=INK,
                        arrowprops={"arrowstyle": "-", "color": INK2, "lw": 0.6})
            ufpb = {"ruf": int(x), "scimago": int(y)}
    ax.set_yscale("log")
    ax.set_yticks([50, 100, 500, 1000, 5000, 10000])
    ax.set_yticklabels(["50", "100", "500", "1.000", "5.000", "10.000"])
    ax.set_xlabel(f"Posição no RUF {ano} (1 = melhor)")
    ax.set_ylabel(f"Posição global no SCImago {ano} (log)")
    ax.legend(loc="lower right", fontsize=8, title=f"ρ = {virgula(rho)}  ·  n = {n}", title_fontsize=8)
    salvar(fig, "fig6_ruf_scimago.png")
    return tabela, {"ano": ano, "rho": round(rho, 3), "n": n, "ufpb": ufpb}


# ---------------------------------------------------------------- hipóteses
def media_abs(regs):
    return mean(abs(r["rho"]) for r in regs) if regs else None


def hipoteses(c):
    col, pesq = c["colunas_qs_ruf"], c["a_pesquisa_internacional"]
    sel = lambda lista, rot: [r for r in lista if r["rotulo"] == rot]  # noqa: E731

    def bloco(intl_regs, br_regs):
        mi = mean(media_abs(g) for g in intl_regs)
        mb = mean(media_abs(g) for g in br_regs)
        return {"internacional": round(mi, 3), "brasileiro": round(mb, 3),
                "razao_intl_br": round(mi / mb, 3), "razao_br_intl": round(mb / mi, 3)}

    h = {}
    h["pesquisa"] = bloco(
        [sel(pesq, "QS: posição × citações/docente"), sel(pesq, "ARWU × Leiden: posição × publicações"),
         sel(pesq, "ARWU × Leiden: posição × % top 10%")],
        [sel(col, "RUF: Pesquisa")])
    h["ensino"] = bloco([sel(col, "QS: Proporção docente/aluno")], [sel(col, "RUF: Ensino")])
    h["mercado"] = bloco([sel(col, "QS: Reputação com empregadores"), sel(col, "QS: Empregabilidade")],
                         [sel(col, "RUF: Mercado")])
    h["internacionalizacao"] = bloco([sel(col, "QS: Docentes internacionais"), sel(col, "QS: Alunos internacionais")],
                                     [sel(col, "RUF: Internacionalização")])
    h["a_suportada"] = h["pesquisa"]["razao_intl_br"] >= 1.10
    h["b_ensino_suportada"] = h["ensino"]["razao_br_intl"] >= 1.10
    h["b_mercado_suportada"] = h["mercado"]["razao_br_intl"] >= 1.10

    # H(c): amplitude relativa de |rho| ao longo das edições, por indicador
    por_ind = defaultdict(list)
    for r in col:
        por_ind[(r["ranking"], r["coluna"])].append(abs(r["rho"]))
    for r in pesq:
        if r["rotulo"].startswith("ARWU"):
            por_ind[("ARWU×Leiden", r["rotulo"].split("× ")[-1])].append(abs(r["rho"]))
    amp = {}
    for (rk, ind), vals in por_ind.items():
        amp[f"{rk} · {ind}"] = {"ranking": rk, "min": round(min(vals), 3), "max": round(max(vals), 3),
                                "edicoes": len(vals),
                                "amplitude_rel": round((max(vals) - min(vals)) / mean(vals), 3)}
    h["c_amplitude"] = amp
    grupo = lambda rk: "brasileiro" if rk == "RUF" else "internacional"  # noqa: E731
    for g in ("internacional", "brasileiro"):
        vs = [v["amplitude_rel"] for v in amp.values() if grupo(v["ranking"]) == g]
        h[f"c_amp_media_{g}"] = round(mean(vs), 3)
        h[f"c_estaveis_{g}"] = sum(1 for v in vs if v < 0.05)
        h[f"c_total_{g}"] = len(vs)
    difs = {k: round(abs(v["internacional"] - v["brasileiro"]) / max(v["internacional"], v["brasileiro"]), 3)
            for k, v in h.items() if isinstance(v, dict) and "internacional" in v}
    h["c_diferenca_rel_grupos"] = difs
    return h


def main():
    with open(ENTRADA, encoding="utf-8") as f:
        d = json.load(f)
    c = d["correlacoes"]
    res = {"cobertura": fig_cobertura(d)}
    res["ultima_edicao"] = fig_peso_indicador(c["colunas_qs_ruf"])
    fig_variacao_temporal(c["colunas_qs_ruf"])
    fig_estabilidade(c["estabilidade"])
    fig_inep_internacional(c["inep_x_internacional"])
    res["concordancia_grupos"], res["dispersao"] = concordancia_grupos(d)
    res["hipoteses"] = hipoteses(c)
    res["correlacoes"] = c
    os.makedirs(os.path.dirname(SAIDA_RES), exist_ok=True)
    with open(SAIDA_RES, "w", encoding="utf-8") as f:
        json.dump(res, f, ensure_ascii=False, indent=1)
    print("resultados:", os.path.relpath(SAIDA_RES, RAIZ))
    h = res["hipoteses"]
    for k in ("pesquisa", "ensino", "mercado", "internacionalizacao"):
        print(f"  {k}: {h[k]}")
    print("  H(a):", h["a_suportada"], " H(b) ensino:", h["b_ensino_suportada"], " mercado:", h["b_mercado_suportada"])
    print("  H(c) amplitude média:", h["c_amp_media_internacional"], h["c_amp_media_brasileiro"],
          " estáveis(<5%):", h["c_estaveis_internacional"], "/", h["c_total_internacional"], ";",
          h["c_estaveis_brasileiro"], "/", h["c_total_brasileiro"])
    for t in res["concordancia_grupos"]:
        print("  ", t)
    print("  dispersão:", res["dispersao"])


if __name__ == "__main__":
    main()
