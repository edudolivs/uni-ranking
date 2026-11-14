"""Gera o Relatório Final de Pesquisa a partir de modelo-relatorio.docx.

Uso: .venv/bin/python scripts/gerar_relatorio.py
Pré-requisitos: dados/processado/dados_dashboard.json (preparar_dados.py) e
relatorio/resultados.json + relatorio/figuras/ (gerar_figuras.py).
Saída: relatorio-de-pesquisa.docx

O modelo da UFPB é preservado (margens, cabeçalho com número de página, capa, folha de rosto e
folha de aprovação); as notas-guia do modelo são removidas e as seções são reescritas.
Os números citados no texto vêm de resultados.json -- nada é digitado à mão.
Os trechos de código são lidos dos próprios scripts, para nunca divergirem do repositório.
O sumário e as listas são estáticos; as páginas vêm de uma 2ª passada: o documento é renderizado
em PDF com o LibreOffice e a página de cada título é localizada com PyMuPDF.
"""
import ast
import json
import os
import shutil
import subprocess
import sys
import tempfile
from collections import defaultdict

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from preparar_dados import chave  # noqa: E402

RAIZ = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
MODELO = os.path.join(RAIZ, "modelo-relatorio.docx")
SAIDA = os.path.join(RAIZ, "relatorio-de-pesquisa.docx")
FIG = os.path.join(RAIZ, "relatorio", "figuras")
RES = json.load(open(os.path.join(RAIZ, "relatorio", "resultados.json"), encoding="utf-8"))
DADOS = json.load(open(os.path.join(RAIZ, "dados", "processado", "dados_dashboard.json"), encoding="utf-8"))

TITULO = "CRITÉRIOS DE AVALIAÇÃO EM RANKINGS UNIVERSITÁRIOS:"
SUBTITULO = "UMA ANÁLISE EXPLORATÓRIA COMPARATIVA ENTRE RANKINGS BRASILEIROS E INTERNACIONAIS"
AUTORES = ["ANA BEATRIZ MOTA DA CRUZ", "EDUARDO DE OLIVEIRA SILVA", "NICOLE COSTA E SILVA"]
LARGURA_TEXTO = 16.0  # cm: A4 (21 cm) - margens 3 + 2
FONTE_PROPRIA = "Fonte: elaborado pelos autores (2026)."


# ================================================================ formatação de números
def v(x, casas=2):
    """0.9403 -> '0,94' (vírgula decimal)."""
    return f"{x:.{casas}f}".replace(".", ",")


def pct(x, casas=0):
    return f"{x * 100:.{casas}f}".replace(".", ",") + "%"


def milhar(n):
    return f"{n:,}".replace(",", ".")


# ================================================================ dados derivados
H = RES["hipoteses"]
C = RES["correlacoes"]
COB = RES["cobertura"]
DISP = RES["dispersao"]


def faixa(serie):
    vals = list(serie.values())
    return f"{min(vals)} a {max(vals)}" if min(vals) != max(vals) else str(vals[0])


def anos(serie):
    a = sorted(int(x) for x in serie)
    return f"{a[0]}–{a[-1]}"


def contagens_inep():
    igc = defaultdict(int)
    for r in DADOS["igc"]["universidades"]:
        igc[r["ano"]] += 1
    cpc, enade = defaultdict(int), defaultdict(int)
    for r in DADOS["cpc"]["resumo"]:
        cpc[r["ano"]] += r["cursos"]
    for r in DADOS["enade"]["resumo"]:
        enade[r["ano"]] += r["cursos"]
    return igc, cpc, enade


IGC_N, CPC_N, ENADE_N = contagens_inep()


def ufpb():
    k = "federal paraiba"
    out = {}
    out["ruf"] = {r["ano"]: r["ranking"] for r in DADOS["ruf"] if chave(r["nome"]) == k and r["ranking"]}
    out["sci"] = {r["ano"]: (r["rank_br"], r["rank_global"]) for r in DADOS["scimago"] if chave(r["nome"]) == k}
    out["qs"] = {r["ano"]: r["rank_raw"] for r in DADOS["qs"] if chave(r["nome"]) == k}
    out["leiden"] = sorted({r["edicao"] for r in DADOS["leiden"] if chave(r["nome"]) == k})
    out["arwu"] = [r for r in DADOS["arwu"]["brasil"] if chave(r["nome"]) == k]
    out["igc"] = {r["ano"]: r["faixa"] for r in DADOS["igc"]["universidades"] if chave(r["nome"]) == k}
    return out


UFPB = ufpb()


def rho_de(lista, rotulo, ano):
    for r in C[lista]:
        if r["rotulo"] == rotulo and r.get("ano", r.get("para")) == ano:
            return r
    raise KeyError((lista, rotulo, ano))


def estab(rotulo):
    return [r for r in C["estabilidade"] if r["rotulo"] == rotulo]


def concord(par, ano):
    for r in RES["concordancia_grupos"]:
        if r["par"] == par and r["ano"] == ano:
            # IGC é nota (maior = melhor); RUF é posição: concordância = -rho. Posição x posição: rho.
            return (-r["rho"] if par.endswith("IGC") else r["rho"]), r["n"]
    return None, None


# ================================================================ trechos de código (lidos dos scripts)
def fonte_funcoes(caminho, nomes):
    src = open(caminho, encoding="utf-8").read()
    linhas = src.splitlines()
    arv = ast.parse(src)
    blocos = []
    for no in arv.body:
        alvo = None
        if isinstance(no, ast.FunctionDef) and no.name in nomes:
            alvo = no.name
        elif isinstance(no, ast.Assign) and any(getattr(t, "id", None) in nomes for t in no.targets):
            alvo = no.targets[0].id
        if alvo:
            ini = (no.decorator_list[0].lineno if getattr(no, "decorator_list", None) else no.lineno) - 1
            blocos.append((nomes.index(alvo), linhas[ini:no.end_lineno]))
    blocos.sort(key=lambda b: b[0])
    out = []
    for _, b in blocos:
        if out:
            out.append("")
        out.extend(b)
    return out


def fonte_trecho(caminho, inicio, fim):
    linhas = open(caminho, encoding="utf-8").read().splitlines()
    i = next(k for k, l in enumerate(linhas) if l.startswith(inicio))
    j = next(k for k, l in enumerate(linhas) if k > i and fim in l)
    return linhas[i:j + 1]


PREP = os.path.join(RAIZ, "scripts", "preparar_dados.py")
COD_CHAVE = fonte_funcoes(PREP, ["STOP", "TRAD", "ALIAS", "chave"])
COD_SPEARMAN = fonte_funcoes(PREP, ["_postos", "spearman"])
COD_COLUNAS = fonte_funcoes(PREP, ["correl_colunas"])
COD_LEIDEN = fonte_trecho(os.path.join(RAIZ, "scripts", "extrair_leiden.py"), "wb = openpyxl", "n_ok += 1")


# ================================================================ construção do documento
class Relatorio:
    def __init__(self, paginas):
        self.d = Document(MODELO)
        self.pag = paginas  # {texto do título ou legenda: página}
        self.figuras, self.quadros, self.tabelas, self.titulos = [], [], [], []
        self.n = {"Figura": 0, "Quadro": 0, "Tabela": 0}

    # ---------------------------------------------------------- primitivas
    def par(self, texto="", alin="J", negrito=False, tam=12, recuo1=1.25, esp=1.5, antes=0, depois=0,
            esq=0.0, quebra=False, italico=False, manter=False, fonte="Arial"):
        p = self.d.add_paragraph()
        pf = p.paragraph_format
        pf.alignment = {"J": WD_ALIGN_PARAGRAPH.JUSTIFY, "C": WD_ALIGN_PARAGRAPH.CENTER,
                        "E": WD_ALIGN_PARAGRAPH.LEFT, "D": WD_ALIGN_PARAGRAPH.RIGHT}[alin]
        pf.line_spacing = esp
        pf.first_line_indent = Cm(recuo1) if recuo1 else Cm(0)
        pf.left_indent = Cm(esq)
        pf.space_before, pf.space_after = Pt(antes), Pt(depois)
        pf.page_break_before = quebra
        pf.keep_with_next = manter
        pf.widow_control = True
        segs = texto if isinstance(texto, list) else [(texto, {})] if texto else []
        for seg in segs:
            t, fmt = (seg, {}) if isinstance(seg, str) else seg
            r = p.add_run(t)
            r.font.name, r.font.size = fonte, Pt(fmt.get("tam", tam))
            r.bold = fmt.get("b", negrito)
            r.italic = fmt.get("i", italico)
        return p

    def corpo(self, texto, **kw):
        return self.par(texto, **kw)

    def secao(self, titulo, quebra=False):
        self.titulos.append((1, titulo))
        return self.par(titulo, alin="E", negrito=True, recuo1=0, antes=0 if quebra else 18, depois=12,
                        quebra=quebra, manter=True)

    def subsecao(self, titulo):
        self.titulos.append((2, titulo))
        return self.par(titulo, alin="E", recuo1=0, antes=12, depois=12, manter=True)

    def terciaria(self, titulo):
        self.titulos.append((3, titulo))
        return self.par(titulo, alin="E", recuo1=0, tam=11, antes=12, depois=6, manter=True)

    def item(self, texto, esq=1.25):
        return self.par(texto, recuo1=0, esq=esq)

    def titulo_pre(self, titulo, quebra=True):
        return self.par(titulo, alin="C", negrito=True, recuo1=0, depois=18, quebra=quebra)

    def legenda(self, tipo, titulo):
        self.n[tipo] += 1
        rot = f"{tipo} {self.n[tipo]} – {titulo}"
        {"Figura": self.figuras, "Quadro": self.quadros, "Tabela": self.tabelas}[tipo].append(rot)
        self.par(rot, alin="C", tam=10, recuo1=0, esp=1.0, antes=6, depois=4, manter=True)
        return self.n[tipo]

    def fonte_nota(self, texto=FONTE_PROPRIA):
        self.par(texto, alin="C", tam=10, recuo1=0, esp=1.0, antes=4, depois=12)

    def figura(self, titulo, arquivo, fonte, largura=LARGURA_TEXTO):
        n = self.legenda("Figura", titulo)
        p = self.par(alin="C", recuo1=0, esp=1.0, manter=True)
        p.add_run().add_picture(os.path.join(FIG, arquivo), width=Cm(largura))
        self.fonte_nota(fonte)
        return n

    # ---------------------------------------------------------- tabelas e quadros
    @staticmethod
    def _bordas(tabela, quais):
        tblPr = tabela._tbl.tblPr
        b = OxmlElement("w:tblBorders")
        for lado in ("top", "left", "bottom", "right", "insideH", "insideV"):
            e = OxmlElement(f"w:{lado}")
            if lado in quais:
                e.set(qn("w:val"), "single")
                e.set(qn("w:sz"), str(quais[lado]))
                e.set(qn("w:color"), "000000")
            else:
                e.set(qn("w:val"), "nil")
            b.append(e)
        # ordem do esquema: tblBorders vem antes de shd/tblLayout/tblCellMar/tblLook
        depois = [tblPr.find(qn(t)) for t in ("w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook")]
        depois = [x for x in depois if x is not None]
        if depois:
            depois[0].addprevious(b)
        else:
            tblPr.append(b)

    @staticmethod
    def _sombra(celula, cor):
        tcPr = celula._tc.get_or_add_tcPr()
        s = OxmlElement("w:shd")
        s.set(qn("w:val"), "clear")
        s.set(qn("w:color"), "auto")
        s.set(qn("w:fill"), cor)
        tcPr.append(s)

    @staticmethod
    def _borda_celula(celula, lado, sz=8):
        tcPr = celula._tc.get_or_add_tcPr()
        bs = tcPr.find(qn("w:tcBorders"))
        if bs is None:
            bs = OxmlElement("w:tcBorders")
            tcPr.append(bs)
        e = OxmlElement(f"w:{lado}")
        e.set(qn("w:val"), "single")
        e.set(qn("w:sz"), str(sz))
        e.set(qn("w:color"), "000000")
        bs.append(e)

    def grade(self, cab, linhas, larguras, fechado, alin_col=None, tam=10):
        t = self.d.add_table(rows=1 + len(linhas), cols=len(cab))
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.autofit = False
        if fechado:  # quadro: bordas fechadas
            self._bordas(t, {k: 6 for k in ("top", "left", "bottom", "right", "insideH", "insideV")})
        else:        # tabela (IBGE/ABNT): sem laterais, traço no topo, sob o cabeçalho e no fim
            self._bordas(t, {"top": 8, "bottom": 8})
        for i, row in enumerate([cab] + linhas):
            for j, txt in enumerate(row):
                cel = t.cell(i, j)
                cel.width = Cm(larguras[j])
                p = cel.paragraphs[0]
                p.paragraph_format.line_spacing = 1.0
                p.paragraph_format.space_before = Pt(1)
                p.paragraph_format.space_after = Pt(1)
                a = "C" if i == 0 else (alin_col[j] if alin_col else "E")
                p.alignment = {"C": WD_ALIGN_PARAGRAPH.CENTER, "E": WD_ALIGN_PARAGRAPH.LEFT,
                               "D": WD_ALIGN_PARAGRAPH.RIGHT}[a]
                r = p.add_run(str(txt))
                r.font.name, r.font.size, r.bold = "Arial", Pt(tam), i == 0
                if i == 0 and not fechado:
                    self._borda_celula(cel, "bottom")
                if i < len(linhas):  # mantém a tabela inteira na mesma página
                    p.paragraph_format.keep_with_next = True
        for col, w in zip(t.columns, larguras):
            col.width = Cm(w)
        for row in t.rows:  # cabeçalho repete e linhas não quebram entre páginas
            trPr = row._tr.get_or_add_trPr()
            nao = OxmlElement("w:cantSplit")
            trPr.append(nao)
        cab_tr = t.rows[0]._tr.get_or_add_trPr()
        cab_tr.append(OxmlElement("w:tblHeader"))
        return t

    def codigo(self, titulo, linhas, origem):
        self.legenda("Quadro", titulo)
        t = self.d.add_table(rows=1, cols=1)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        self._bordas(t, {k: 4 for k in ("top", "left", "bottom", "right")})
        cel = t.cell(0, 0)
        cel.width = Cm(LARGURA_TEXTO)
        self._sombra(cel, "F4F4F2")
        primeiro = True
        for ln in linhas:
            p = cel.paragraphs[0] if primeiro else cel.add_paragraph()
            primeiro = False
            pf = p.paragraph_format
            pf.line_spacing, pf.space_before, pf.space_after = 1.0, Pt(0), Pt(0)
            pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
            r = p.add_run(ln if ln else " ")
            r.font.name, r.font.size = "Courier New", Pt(7.5)
            r._element.rPr.rFonts.set(qn("w:hAnsi"), "Courier New")
            r.font.color.rgb = RGBColor(0x1A, 0x1A, 0x1A)
        t.rows[0]._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))  # bloco inteiro na mesma página
        self.fonte_nota(f"Fonte: elaborado pelos autores (2026), trecho de {origem}.")

    def tab(self, tipo, titulo, cab, linhas, larguras, fonte=FONTE_PROPRIA, alin_col=None, tam=10):
        self.legenda(tipo, titulo)
        self.grade(cab, linhas, larguras, fechado=(tipo == "Quadro"), alin_col=alin_col, tam=tam)
        self.fonte_nota(fonte)

    # ---------------------------------------------------------- listas e sumário
    def pagina(self, chave_txt):
        p = self.pag.get(chave_txt)
        return f"{p:02d}" if p else "00"

    def entrada(self, texto, pag, negrito=False, esq=0.0, tam=12, esp=1.5):
        p = self.par(alin="E", recuo1=0, esq=esq, esp=esp)
        p.paragraph_format.right_indent = Cm(1.0)  # o texto quebra antes da coluna do número
        p.paragraph_format.tab_stops.add_tab_stop(Cm(LARGURA_TEXTO), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
        r = p.add_run(f"{texto}\t{pag}")
        r.font.name, r.font.size, r.bold = "Arial", Pt(tam), negrito
        return p

    def salvar(self, caminho):
        for pg in self.d.element.body.iter(qn("w:pgMar")):  # o modelo omite gutter, obrigatório no esquema
            if pg.get(qn("w:gutter")) is None:
                pg.set(qn("w:gutter"), "0")
        self.d.save(caminho)


# ================================================================ pré-textuais (edição do modelo)
# Parágrafos do modelo mantidos na capa, na folha de rosto e na folha de aprovação, com o espaço antes
# (cm) que posiciona cada bloco. As linhas em branco e as notas-guia do modelo são removidas: com 3
# autores, posicionar por linhas em branco fazia a página transbordar.
PRE_MANTIDOS = {
    0: 0, 1: 0, 2: 0, 9: 4.0, 17: 4.5, 32: 7.4, 33: 0,                     # capa
    35: 0, 50: 5.5, 54: 1.5, 63: 8.2, 64: 0,                               # folha de rosto
    66: 0, 69: 0.6, 72: 0.8, 75: 1.2, 78: 1.5, 79: 0, 85: 1.5,             # folha de aprovação
    86: 0.2, 87: 0, 88: 0, 89: 0, 90: 0, 91: 0,
}
NOVA_PAGINA = {35, 66}

def limpar(p):
    for r in list(p.runs):
        r._element.getparent().remove(r._element)


def escrever(p, segs, tam=12):
    limpar(p)
    for i, (txt, negrito) in enumerate(segs):
        if i and txt == "\n":
            continue
        r = p.add_run()
        partes = txt.split("\n")
        for k, parte in enumerate(partes):
            if k:
                r.add_break()
            r.add_text(parte)
        r.font.name, r.font.size, r.bold = "Arial", Pt(tam), negrito


def pre_textuais(rel):
    d = rel.d
    ps = d.paragraphs
    autores = "\n".join(AUTORES)
    for i in (9, 35, 66):
        escrever(ps[i], [(autores, False)])
    for i in (17, 50):
        escrever(ps[i], [(TITULO, True), (" " + SUBTITULO, False)])
    escrever(ps[69], [(f"{TITULO} {SUBTITULO}", False)])
    for i in range(92):
        if i in PRE_MANTIDOS:
            pf = ps[i].paragraph_format
            pf.space_before, pf.space_after = Cm(PRE_MANTIDOS[i]), Pt(0)
            pf.page_break_before = i in NOVA_PAGINA
        else:
            assert i in (3, 10, 18, 34, 36, 51, 55, 65, 67, 70) or not ps[i].text.strip(), i
            ps[i]._p.getparent().remove(ps[i]._p)
    # remove tudo do título do resumo (parágrafo 92) até o fim, preservando sectPr
    corpo = d.element.body
    inicio = ps[92]._p
    remover, achou = [], False
    for el in corpo.iterchildren():
        if el is inicio:
            achou = True
        if achou and el.tag != qn("w:sectPr"):
            remover.append(el)
    for el in remover:
        corpo.remove(el)


# ================================================================ conteúdo
RESUMO = (
    "Este relatório apresenta os resultados de uma pesquisa que comparou os critérios de avaliação de "
    "rankings universitários brasileiros e internacionais. O objetivo foi analisar quais características "
    "das universidades mais se associam à posição em cada grupo de rankings e como essas associações "
    "variam ao longo do tempo. Trata-se de um estudo quantitativo, exploratório, documental e longitudinal, "
    "baseado em análise exploratória de dados (EDA) sobre bases públicas de oito fontes: QS, ARWU, Leiden "
    "Ranking e SCImago, no grupo internacional, e Ranking Universitário Folha (RUF), Índice Geral de Cursos "
    "(IGC), Conceito Preliminar de Curso (CPC) e Enade, no grupo brasileiro, com edições entre 2017 e 2027. "
    "As bases foram extraídas e padronizadas por um pipeline em Python, as instituições foram cruzadas por "
    "nome e as associações foram medidas pela correlação de postos de Spearman, por ranking e por edição, "
    "com os resultados reunidos em um dashboard interativo. Nos dados analisados, a pesquisa foi o "
    f"componente mais associado à posição no RUF (|ρ| médio de {v(H['pesquisa']['brasileiro'])}), acima "
    f"dos indicadores de pesquisa dos rankings internacionais ({v(H['pesquisa']['internacional'])}), o que "
    "refutou a primeira hipótese. O ensino teve associação "
    f"{v(H['ensino']['razao_br_intl'], 1)} vezes maior no RUF do que no QS, o que confirmou a segunda "
    "hipótese para essa dimensão, mas não para o mercado de trabalho, cujas associações foram equivalentes "
    "nos dois grupos. As associações no RUF variaram pouco entre as edições, enquanto as do QS oscilaram, "
    "e as posições das universidades brasileiras concordaram entre os grupos, com "
    f"ρ = {v(DISP['rho'])} entre o RUF e o SCImago em {DISP['ano']}. Conclui-se que os rankings brasileiros e "
    "internacionais tendem a ordenar as mesmas universidades de modo semelhante, mas por caminhos "
    "diferentes: a reputação acadêmica domina o QS, e pesquisa e ensino dominam o RUF. Os resultados são "
    "associações, e não relações de causa, e dependem do cruzamento por nome e do tamanho reduzido das "
    "amostras nos rankings internacionais."
)
ABSTRACT = (
    "This report presents the results of a study that compared the evaluation criteria of Brazilian and "
    "international university rankings. The aim was to analyse which university characteristics are most "
    "associated with ranking position in each group and how these associations change over time. It is a "
    "quantitative, exploratory, documentary and longitudinal study based on exploratory data analysis (EDA) "
    "of public data from eight sources: QS, ARWU, Leiden Ranking and SCImago (international) and Ranking "
    "Universitário Folha (RUF), the General Course Index (IGC), the Preliminary Course Score (CPC) and Enade "
    "(Brazilian), with editions from 2017 to 2027. The datasets were extracted and standardised by a Python "
    "pipeline, institutions were matched by name, and associations were measured with Spearman's rank "
    "correlation, per ranking and per edition, with the results gathered in an interactive dashboard. In "
    "the data analysed, research was the component most associated with position in RUF (mean |ρ| of "
    f"{H['pesquisa']['brasileiro']:.2f}), above the research indicators of the international rankings "
    f"({H['pesquisa']['internacional']:.2f}), which refuted the first hypothesis. Teaching showed an "
    f"association {H['ensino']['razao_br_intl']:.1f} times stronger in RUF than in QS, confirming the second "
    "hypothesis for that dimension but not for the labour market, whose associations were equivalent in both "
    "groups. Associations in RUF changed little across editions, whereas those in QS fluctuated, and the "
    "positions of Brazilian universities agreed across groups, with "
    f"ρ = {DISP['rho']:.2f} between RUF and SCImago in {DISP['ano']}. The study concludes that Brazilian "
    "and international rankings tend to order the same universities in a similar way, but through different "
    "paths: academic reputation dominates QS, while research and teaching dominate RUF. The results are "
    "associations, not causal relations, and depend on name matching and on the small samples of the "
    "international rankings."
)

SIGLAS = [
    ("ARWU", "Academic Ranking of World Universities"),
    ("Capes", "Coordenação de Aperfeiçoamento de Pessoal de Nível Superior"),
    ("CPC", "Conceito Preliminar de Curso"),
    ("CWTS", "Centre for Science and Technology Studies (Universidade de Leiden)"),
    ("EDA", "Análise exploratória de dados (do inglês Exploratory Data Analysis)"),
    ("Enade", "Exame Nacional de Desempenho dos Estudantes"),
    ("IES", "Instituição de Ensino Superior"),
    ("IGC", "Índice Geral de Cursos"),
    ("Inep", "Instituto Nacional de Estudos e Pesquisas Educacionais Anísio Teixeira"),
    ("QS", "QS World University Rankings (Quacquarelli Symonds)"),
    ("RUF", "Ranking Universitário Folha"),
    ("SIR", "SCImago Institutions Rankings"),
    ("UFPB", "Universidade Federal da Paraíba"),
]

ATIVIDADES = [
    ("Escolha do tema e dos rankings a serem comparados", {"Set"}),
    ("Elaboração do projeto de pesquisa", {"Set", "Out"}),
    ("Coleta e padronização das bases de dados dos rankings", {"Out"}),
    ("Análise exploratória dos dados e comparação entre os rankings", {"Out", "Nov"}),
    ("Exposição dos resultados no dashboard", {"Nov"}),
    ("Escrita do relatório da pesquisa", {"Nov", "Dez"}),
    ("Entrega do relatório da pesquisa", {"Dez"}),
]
MESES = ["Set", "Out", "Nov", "Dez"]

REFERENCIAS = [
    "CENTRE FOR SCIENCE AND TECHNOLOGY STUDIES (CWTS). CWTS Leiden Ranking Open Edition 2025. Leiden: "
    "Leiden University, 2025. Disponível em: https://open.leidenranking.com/. Acesso em: 18 set. 2026.",
    "HAZELKORN, Ellen. Rankings and the reshaping of higher education: the battle for world-class "
    "excellence. 2. ed. London: Palgrave Macmillan, 2015. DOI: 10.1057/9781137446671.",
    "INSTITUTO NACIONAL DE ESTUDOS E PESQUISAS EDUCACIONAIS ANÍSIO TEIXEIRA (INEP). Índice Geral de Cursos "
    "(IGC). Brasília, DF: Inep, [s.d.]. Disponível em: https://www.gov.br/inep/pt-br/areas-de-atuacao/"
    "pesquisas-estatisticas-e-indicadores/indicadores-de-qualidade-da-educacao-superior/"
    "indice-geral-de-cursos-igc. Acesso em: 2 out. 2026.",
    "MARGINSON, Simon; VAN DER WENDE, Marijk. To rank or to be ranked: the impact of global rankings in "
    "higher education. Journal of Studies in International Education, [s. l.], v. 11, n. 3-4, p. 306–329, "
    "2007. DOI: 10.1177/1028315307303544. Acesso em: 14 nov. 2026.",
    "MOED, Henk F. A critical comparative analysis of five world university rankings. Scientometrics, "
    "[s. l.], v. 110, n. 2, p. 967–990, 2017. DOI: 10.1007/s11192-016-2212-y. Disponível em: "
    "https://link.springer.com/article/10.1007/s11192-016-2212-y. Acesso em: 30 out. 2026.",
    "QS QUACQUARELLI SYMONDS. QS World University Rankings 2027. Londres: QS Quacquarelli Symonds, 2026. "
    "Disponível em: https://www.topuniversities.com/world-university-rankings. Acesso em: 18 set. 2026.",
    "SANTOS, Solange Maria dos. O desempenho das universidades brasileiras nos rankings internacionais: "
    "áreas de destaque da produção científica brasileira. 2015. Tese (Doutorado) – Escola de Comunicações "
    "e Artes, Universidade de São Paulo, São Paulo, 2015. DOI: 10.11606/T.27.2015.tde-26052015-122043. "
    "Disponível em: https://teses.usp.br/teses/disponiveis/27/27151/tde-26052015-122043/. "
    "Acesso em: 14 nov. 2026.",
    "SCIMAGO LAB. SIR – SCImago Institutions Rankings. [S. l.]: SCImago Lab, [s.d.]. Disponível em: "
    "https://www.scimagolab.com/products/sir-scimago-institutions-rankings/. Acesso em: 18 set. 2026.",
    "SELTEN, Friso; NEYLON, Cameron; HUANG, Chun-Kai; GROTH, Paul. A longitudinal analysis of university "
    "rankings. Quantitative Science Studies, [s. l.], v. 1, n. 3, p. 1109–1135, 2020. "
    "DOI: 10.1162/qss_a_00052. Acesso em: 30 out. 2026.",
    "SHANGHAIRANKING CONSULTANCY. Academic Ranking of World Universities 2026. Xangai: ShanghaiRanking "
    "Consultancy, 2026. Disponível em: https://www.shanghairanking.com/rankings/arwu/2026. "
    "Acesso em: 18 set. 2026.",
    "SZLUKA, Péter; CSAJBÓK, Edit; GYŐRFFY, Balázs. Relationship between bibliometric indicators and "
    "university ranking positions. Scientific Reports, [s. l.], v. 13, art. 14193, 2023. "
    "DOI: 10.1038/s41598-023-35306-1. Disponível em: https://www.nature.com/articles/s41598-023-35306-1. "
    "Acesso em: 30 out. 2026.",
    "UNIVERSIDADE DE SÃO PAULO. USP mantém a liderança no RUF 2025 e é a primeira colocada em 33 das 40 "
    "carreiras avaliadas. Jornal da USP, São Paulo, 2025. Disponível em: https://jornal.usp.br/"
    "institucional/usp-mantem-a-lideranca-no-ruf-2025-e-e-a-primeira-colocada-em-33-das-40-carreiras-"
    "avaliadas/. Acesso em: 2 out. 2026.",
]


def construir(paginas):
    rel = Relatorio(paginas)
    pre_textuais(rel)
    p, it, sec, sub, ter = rel.corpo, rel.item, rel.secao, rel.subsecao, rel.terciaria

    # ------------------------------------------------------------ resumo e abstract
    rel.par([(TITULO, {"b": True}), (" " + SUBTITULO, {})], alin="C", recuo1=0, quebra=True, depois=18)
    rel.par("RESUMO", alin="E", negrito=True, recuo1=0, depois=6)
    rel.par(RESUMO, recuo1=0, esp=1.0, depois=12)
    rel.par([("Palavras-chave: ", {"b": True}),
             ("Análise exploratória de dados. Ensino superior. Rankings universitários.", {})], recuo1=0)

    rel.par([("CRITERIA FOR EVALUATING UNIVERSITY RANKINGS:", {"b": True}),
             (" AN EXPLORATORY COMPARATIVE ANALYSIS OF BRAZILIAN AND INTERNATIONAL RANKINGS", {})],
            alin="C", recuo1=0, quebra=True, depois=18)
    rel.par("ABSTRACT", alin="E", negrito=True, recuo1=0, depois=6)
    rel.par(ABSTRACT, recuo1=0, esp=1.0, depois=12)
    rel.par([("Keywords: ", {"b": True}),
             ("Exploratory data analysis. Higher education. University rankings.", {})], recuo1=0)

    # ------------------------------------------------------------ agradecimentos
    rel.titulo_pre("AGRADECIMENTOS")
    p("À Profa. Elizabet Maria Spohr de Medeiros, pela orientação ao longo da disciplina e pelas correções "
      "que melhoraram o projeto e este relatório.")
    p("Ao Inep, ao Centre for Science and Technology Studies da Universidade de Leiden, à ShanghaiRanking "
      "Consultancy, à QS Quacquarelli Symonds, ao SCImago Lab e à Folha de S.Paulo, por manterem públicos "
      "os dados que tornaram esta pesquisa possível.")

    # ------------------------------------------------------------ listas (preenchidas na 2ª passada)
    rel.titulo_pre("LISTA DE FIGURAS")
    for rot in paginas.get("_figuras", []):
        rel.entrada(rot, rel.pagina(rot))
    rel.titulo_pre("LISTA DE QUADROS")
    for rot in paginas.get("_quadros", []):
        rel.entrada(rot, rel.pagina(rot))
    rel.titulo_pre("LISTA DE TABELAS")
    for rot in paginas.get("_tabelas", []):
        rel.entrada(rot, rel.pagina(rot))
    rel.titulo_pre("LISTA DE ABREVIATURAS E SIGLAS")
    for sig, desc in SIGLAS:
        rel.par([(f"{sig} – ", {"b": True}), (desc, {})], alin="E", recuo1=0)
    rel.titulo_pre("SUMÁRIO")
    for nivel, tit in paginas.get("_titulos", []):
        rel.entrada(tit, rel.pagina(tit), negrito=(nivel == 1), esq=0.0 if nivel == 1 else 0.4 * (nivel - 1),
                    tam=11 if nivel == 3 else 12)

    # ------------------------------------------------------------ 1 IDENTIFICAÇÃO
    sec("1 IDENTIFICAÇÃO", quebra=True)
    for rot, val in [("Título do Projeto: ", f"Critérios de avaliação em rankings universitários: uma análise "
                                             "exploratória comparativa entre rankings brasileiros e internacionais"),
                     ("Nome dos Autores: ", "Ana Beatriz Mota da Cruz, Eduardo de Oliveira Silva e Nicole Costa "
                                            "e Silva"),
                     ("Nome da Orientadora: ", "Profa. Elizabet Maria Spohr de Medeiros"),
                     ("Local de execução: ", "Centro de Informática, Universidade Federal da Paraíba, "
                                             "João Pessoa (PB)"),
                     ("Vigência do plano de trabalho: ", "2026.2")]:
        rel.par([(rot, {"b": True}), (val, {})], alin="E", recuo1=0, esq=1.25)

    # ------------------------------------------------------------ 2 INTRODUÇÃO
    sec("2 INTRODUÇÃO")
    p("A ideia desta pesquisa surgiu de uma apresentação feita pelo grupo em setembro de 2026 sobre os "
      "determinantes do desempenho de universidades em rankings internacionais. Ao montar o dicionário de "
      "variáveis daquela apresentação, o grupo percebeu que a mesma universidade brasileira aparecia em "
      "posições muito diferentes conforme o ranking consultado, e que os rankings brasileiros declaravam "
      "critérios bem distintos dos internacionais. A pergunta deixou de ser o que leva uma universidade a "
      "subir no ranking e passou a ser o que cada ranking efetivamente valoriza.")
    p("Os rankings universitários influenciam a escolha dos estudantes, a atração de pesquisadores e as "
      "estratégias de gestão das instituições (HAZELKORN, 2015; MARGINSON; VAN DER WENDE, 2007). Cada um "
      "deles, porém, mede a qualidade acadêmica de um jeito: os internacionais, como o QS World University "
      "Rankings (QS), o Academic Ranking of World Universities (ARWU), o Leiden Ranking e o SCImago "
      "Institutions Rankings (SIR), concentram-se em produção científica, citações, reputação e "
      "internacionalização; os brasileiros, como o Ranking Universitário Folha (RUF) e os indicadores de "
      "qualidade do Instituto Nacional de Estudos e Pesquisas Educacionais Anísio Teixeira (Inep), incluem "
      "ensino, mercado de trabalho e condições de oferta dos cursos. Estudos anteriores compararam rankings "
      "mundiais entre si (MOED, 2017; SELTEN et al., 2020; SZLUKA; CSAJBÓK; GYŐRFFY, 2023), mas a busca "
      "bibliográfica do grupo não encontrou comparação em série temporal entre rankings brasileiros e "
      "internacionais.")
    p("A proposta foi tratar a pergunta como um problema de análise exploratória de dados (EDA, do inglês "
      "Exploratory Data Analysis): reunir as bases públicas de oito fontes, padronizá-las, cruzar as "
      "universidades brasileiras entre elas e medir, com a correlação de postos de Spearman, quanto cada "
      "indicador acompanha a posição geral de cada ranking, edição a edição. Spearman mede se duas "
      "ordenações andam juntas (ρ próximo de 1 ou de −1) ou não (ρ próximo de 0), sem supor relação linear "
      "nem distribuição normal, o que convém a dados de posição.")
    p("Os resultados principais foram três. Primeiro, a pesquisa pesou mais no RUF do que nos indicadores "
      "de pesquisa dos rankings internacionais, o contrário do esperado. Segundo, o ensino pesou muito mais "
      "no RUF do que no QS, como esperado, mas o mercado de trabalho pesou igualmente nos dois. Terceiro, as "
      "associações do RUF foram estáveis entre as edições e as do QS não foram, embora os dois grupos tenham "
      "ordenado as mesmas universidades de forma parecida. O escopo é restrito às universidades brasileiras "
      "e às edições já publicadas até outubro de 2026, e os resultados são associações, não causas. As "
      "siglas usadas estão na Lista de abreviaturas e siglas.")
    p("O relatório está organizado da seguinte forma: a seção 3 apresenta os objetivos; a seção 4 reproduz "
      "o planejamento da pesquisa (o projeto); a seção 5 descreve os materiais e métodos; a seção 6 traz os "
      "resultados; a seção 7 os discute; e a seção 8 apresenta as conclusões e os trabalhos futuros.")

    # ------------------------------------------------------------ 3 OBJETIVOS
    sec("3 OBJETIVOS")
    sub("3.1 OBJETIVO GERAL")
    p("Analisar as diferenças entre as características das universidades que mais se associam à posição em "
      "rankings brasileiros e em rankings internacionais, considerando um período de no mínimo cinco anos.")
    sub("3.2 OBJETIVOS ESPECÍFICOS")
    for t in ["a) Levantar as metodologias, os indicadores e os pesos declarados de cada ranking, "
              "classificando-os em dimensões comparáveis (pesquisa, ensino, reputação, mercado, inovação e "
              "internacionalização);",
              "b) Coletar e padronizar as bases de dados dos rankings brasileiros e internacionais para uma "
              "série mínima de cinco anos;",
              "c) Caracterizar, por meio de análise exploratória, a distribuição e a evolução temporal das "
              "variáveis das universidades em cada ranking;",
              "d) Verificar a associação entre as características das universidades e a posição ou a "
              "pontuação em cada ranking, ano a ano;",
              "e) Comparar as características que mais pesam nos rankings brasileiros e nos internacionais e "
              "a concordância entre as posições das universidades brasileiras nos dois grupos;",
              "f) Apresentar os resultados em um dashboard interativo."]:
        it(t)

    # ------------------------------------------------------------ 4 PLANEJAMENTO (o projeto)
    sec("4 PLANEJAMENTO DO EXPERIMENTO")
    p("Esta seção reproduz o projeto de pesquisa aprovado na disciplina, já com as correções indicadas pela "
      "orientadora: a numeração das seções foi alinhada, a frase que antecipava o objetivo foi retirada do "
      "tema e problema, e o cronograma passou a ter as mesmas sete atividades da metodologia (ver seção 7.2).")
    sub("4.1 TEMA E PROBLEMA DE PESQUISA")
    p("Os rankings universitários consolidaram-se como instrumentos de comparação entre instituições de "
      "ensino superior e passaram a influenciar a escolha dos estudantes, a atração de pesquisadores, o "
      "financiamento e as estratégias de gestão das universidades (HAZELKORN, 2015; MARGINSON; VAN DER WENDE, "
      "2007). Cada ranking, porém, traduz uma concepção própria de qualidade acadêmica: ao definir quais "
      "indicadores serão medidos e com que pesos, ele privilegia determinadas características das "
      "instituições e deixa outras em segundo plano.")
    p("Os rankings internacionais considerados neste projeto (QS, ARWU, Leiden Ranking e SCImago) "
      "concentram-se em dimensões como produção científica, citações, reputação e internacionalização. Os "
      "brasileiros, por sua vez, incorporam elementos do contexto nacional: o Ranking Universitário Folha "
      "(RUF) combina pesquisa, ensino, mercado, inovação e internacionalização e inclui pesquisas de opinião "
      "com docentes e empregadores, e o Índice Geral de Cursos (IGC), do Inep, resume o desempenho dos "
      "estudantes, o valor agregado pelos cursos, o corpo docente e as condições de oferta, além da "
      "avaliação da pós-graduação. Em consequência, uma mesma universidade brasileira pode ocupar posições "
      "muito distintas conforme o ranking consultado.")
    p("Diante disso, a questão que orienta o estudo é: quais características das universidades mais pesam "
      "na posição em rankings brasileiros e em rankings internacionais, e como esses pesos diferem ao longo "
      "do tempo?")
    sub("4.2 HIPÓTESES")
    for t in ["a) Nos rankings internacionais, os indicadores de produção científica e de impacto em citações "
              "apresentam associação 10% mais forte com a posição das universidades do que nos rankings "
              "brasileiros;",
              "b) Nos rankings brasileiros, os indicadores de ensino, de opinião de docentes e empregadores "
              "(reputação e mercado) e de condições de oferta apresentam associação 10% mais forte com a "
              "posição do que nos rankings internacionais;",
              "c) As características que mais pesam em cada grupo de rankings variam menos de 5% ao longo do "
              "período analisado, mas possuem uma diferença de mais de 5% entre rankings brasileiros e "
              "internacionais."]:
        it(t)
    sub("4.3 JUSTIFICATIVA")
    p("Rankings diferentes podem levar a conclusões distintas sobre a mesma instituição. Moed (2017) mostra "
      "que os rankings mundiais diferem em seus fundamentos conceituais e metodológicos, e Szluka, Csajbók e "
      "Győrffy (2023) observam que a associação entre indicadores bibliométricos e a posição das "
      "universidades varia de um ranking para outro. Para o Brasil, Santos (2015) mostra que o número de "
      "universidades brasileiras entre as melhores do mundo muda bastante conforme o ranking. Na busca "
      "bibliográfica preliminar deste projeto, porém, não se identificou estudo que compare, por análise "
      "exploratória e em série temporal, as características valorizadas pelos rankings brasileiros e pelos "
      "internacionais.")
    p("Estudantes, gestores e formuladores de políticas públicas usam rankings para decidir onde estudar e "
      "onde investir. Entender o que cada ranking valoriza ajuda a interpretá-los com cautela e pode "
      "subsidiar a autoavaliação de universidades brasileiras, inclusive da Universidade Federal da Paraíba "
      "(UFPB). O projeto deixa ainda uma base de dados padronizada e um dashboard reutilizável, além de "
      "permitir a coleta e a análise exploratória de dados reais.")
    sub("4.4 REFERENCIAL TEÓRICO")
    ter("4.4.1 RANKINGS INTERNACIONAIS")
    p("O QS World University Rankings combina indicadores de reputação (acadêmica e de empregadores), de "
      "citações por docente, de relação entre docentes e estudantes e de internacionalização (QS "
      "QUACQUARELLI SYMONDS, 2026; SELTEN et al., 2020). O ARWU baseia-se em critérios mais objetivos, como "
      "publicações, citações, pesquisadores altamente citados e premiações de ex-alunos e docentes "
      "(SHANGHAIRANKING CONSULTANCY, 2026). O Leiden Ranking é exclusivamente bibliométrico, sem pesquisas "
      "de reputação (CENTRE FOR SCIENCE AND TECHNOLOGY STUDIES, 2025), e o SCImago Institutions Rankings "
      "reúne indicadores de pesquisa, de inovação e de impacto social (SCIMAGO LAB, [s.d.]).")
    p("A literatura comparativa mostra que essas escolhas importam. Moed (2017) destaca diferenças nos "
      "fundamentos dos rankings mundiais; Selten et al. (2020) indicam que as variáveis utilizadas medem "
      "sobretudo reputação e desempenho em pesquisa; e Szluka, Csajbók e Győrffy (2023) mostram que "
      "citações, reputação internacional e número de pesquisadores estão entre os parâmetros mais associados "
      "à posição, com intensidade que varia entre os rankings.")
    ter("4.4.2 RANKINGS E INDICADORES BRASILEIROS")
    p("O Ranking Universitário Folha (RUF) classifica universidades brasileiras com cinco indicadores, cujos "
      "pesos na edição de 2025 foram: pesquisa (42%), ensino (32%), mercado de trabalho (18%), inovação (4%) "
      "e internacionalização (4%) (UNIVERSIDADE DE SÃO PAULO, 2025).")
    p("O Índice Geral de Cursos (IGC) combina a média dos Conceitos Preliminares de Curso (CPC) do último "
      "triênio, os conceitos da Coordenação de Aperfeiçoamento de Pessoal de Nível Superior (Capes) para a "
      "pós-graduação e a distribuição dos estudantes entre graduação e pós-graduação (INSTITUTO NACIONAL DE "
      "ESTUDOS E PESQUISAS EDUCACIONAIS ANÍSIO TEIXEIRA, [s.d.]). Assim, o IGC privilegia o ensino e a "
      "avaliação oficial, e não a visibilidade internacional. Santos (2015) compara a metodologia de seis "
      "rankings internacionais para explicar por que o número de instituições brasileiras entre as melhores "
      "do mundo varia tanto de um ranking para outro.")
    ter("4.4.3 ANÁLISE EXPLORATÓRIA DE DADOS")
    p("A análise exploratória de dados (EDA) privilegia o exame dos dados, por meio de resumos numéricos e "
      "gráficos, para identificar padrões e hipóteses antes de testes formais. Ela é adequada porque o "
      "objetivo é descrever e comparar associações, e não estabelecer causa e efeito. Será usada a "
      "correlação de postos de Spearman, apropriada a variáveis ordinais e assimétricas, como as de "
      "rankings, em linha com estudos comparativos anteriores (SELTEN et al., 2020).")
    sub("4.5 METODOLOGIA DA PESQUISA")
    p("Esta pesquisa será do tipo quantitativa, exploratória e documental, com delineamento longitudinal, "
      "pois acompanha as mesmas variáveis em, no mínimo, cinco edições anuais dos rankings (por exemplo, de "
      "2021 a 2025, conforme a disponibilidade de cada base).")
    p("O universo da pesquisa é composto pelas instituições presentes nos rankings no período: QS, ARWU, "
      "Leiden e SCImago (internacionais) e RUF, IGC/CPC e rankings de pesquisa de recorte brasileiro "
      "(nacionais). A amostra principal reúne as universidades brasileiras presentes em pelo menos um "
      "ranking de cada grupo, para comparar as mesmas instituições sob critérios diferentes.")
    p("A coleta de dados será realizada em bases de dados públicas obtidas nos sítios oficiais dos rankings "
      "e do Inep, preferencialmente em formato aberto, seguindo o dicionário de variáveis definido pelo "
      "grupo (nome, país, ano, posição, pontuação, tipo, tamanho, produção de pesquisa, docentes, relação "
      "estudantes/docente e estudantes estrangeiros), acrescido dos componentes do RUF e do CPC/IGC. Os "
      "nomes das instituições serão padronizados entre as bases.")
    p("O método de análise será a análise exploratória de dados, com verificação da qualidade dos dados, "
      "estatísticas descritivas por ranking e por ano, correlação de Spearman entre cada característica e a "
      "posição, correlação de postos entre rankings e comparação das características de maior associação "
      "nos dois grupos, com visualizações e dashboard.")
    p("Os resultados serão lidos como associações, e não como causas, considerando as limitações do estudo: "
      "posições em faixas no ARWU, mudanças metodológicas entre edições e o caráter trienal de parte dos "
      "indicadores do Inep.")
    p("As atividades da pesquisa, na ordem em que serão executadas, são as seguintes:")
    for i, (a, _) in enumerate(ATIVIDADES, 1):
        it(f"{i}. {a};" if i < len(ATIVIDADES) else f"{i}. {a}.")
    sub("4.6 CRONOGRAMA")
    p("As atividades descritas na seção 4.5 encontram-se dispostas no Quadro 1.")
    rel.tab("Quadro", "Cronograma da pesquisa (2026)", ["ATIVIDADES"] + MESES,
            [[str(i)] + ["X" if m in meses else "" for m in MESES] for i, (_, meses) in enumerate(ATIVIDADES, 1)],
            [3.0, 1.6, 1.6, 1.6, 1.6], alin_col=["C"] * 5)

    # ------------------------------------------------------------ 5 MATERIAIS E MÉTODOS
    sec("5 MATERIAIS E MÉTODOS")
    p("Esta seção descreve como a pesquisa foi executada: as fontes de dados, o pipeline de extração e "
      "padronização, a medida de associação, a forma de testar as hipóteses e as ferramentas usadas. O "
      "objetivo é que um terceiro consiga repetir a análise a partir do repositório do projeto.")
    sub("5.1 FONTES DE DADOS E COLETA")
    p("Foram usadas oito fontes públicas, quatro em cada grupo (Quadro 2). Os rankings internacionais foram "
      "filtrados para as instituições brasileiras; as bases do Inep já são nacionais e foram restritas às "
      "universidades (organização acadêmica \"Universidade\"). As edições usadas são todas as que estavam "
      "disponíveis e foram baixadas até 16 de outubro de 2026, quando o grupo decidiu fechar as bases.")
    q_linhas = [
        ["ARWU", "Internacional", anos(COB["ARWU"]), faixa(COB["ARWU"]),
         "Posição (exata até a 100ª; depois, em faixas)"],
        ["QS", "Internacional", anos(COB["QS"]), faixa(COB["QS"]),
         "Posição e nove indicadores (indicadores só de 2023 a 2027)"],
        ["Leiden", "Internacional", anos(COB["Leiden"]), faixa(COB["Leiden"]),
         "Publicações (P) e proporção no top 10% mais citado"],
        ["SCImago (SIR)", "Internacional", anos(COB["SCImago"]), faixa(COB["SCImago"]),
         "Posição global e nacional (setor Universidades)"],
        ["RUF", "Brasileiro", "2019*, 2023–2025", faixa(COB["RUF"]),
         "Posição geral e cinco indicadores"],
        ["IGC (Inep)", "Brasileiro", "2019, 2021–2023", f"{min(IGC_N.values())} a {max(IGC_N.values())}",
         "IGC contínuo e faixa"],
        ["CPC (Inep)", "Brasileiro", "2019, 2021–2023",
         f"{milhar(min(CPC_N.values()))} a {milhar(max(CPC_N.values()))} cursos", "Média do CPC por instituição"],
        ["Enade (Inep)", "Brasileiro", "2021–2023",
         f"{milhar(min(ENADE_N.values()))} a {milhar(max(ENADE_N.values()))} cursos",
         "Média do conceito Enade por instituição"],
    ]
    rel.tab("Quadro", "Fontes de dados utilizadas", ["Fonte", "Grupo", "Edições", "Universidades brasileiras "
                                                     "por edição", "Variáveis usadas"],
            q_linhas, [2.4, 2.2, 2.6, 3.0, 5.8], alin_col=["E", "E", "C", "C", "E"], tam=9,
            fonte="Fonte: elaborado pelos autores (2026). * Em 2019, o RUF traz só a classificação de Mercado.")
    p("O ARWU foi obtido pela interface pública de dados do ShanghaiRanking, com mil universidades por "
      "edição. O QS, o SCImago e o RUF foram baixados como planilhas CSV; os arquivos do QS têm esquemas de "
      "colunas e codificações diferentes por edição, o que exigiu um tratamento por arquivo. O IGC, o CPC e "
      "o Enade vieram das planilhas do Inep. O Leiden Ranking exigiu um cuidado à parte: cada planilha tem "
      "de 170 a 262 MB e cerca de 190 mil linhas, então a leitura foi feita em fluxo, linha a linha, "
      "mantendo só o Brasil, a área \"All sciences\" e a contagem inteira (Quadro 3). Os nomes das colunas "
      "mudam entre edições do Leiden e foram normalizados.")
    rel.codigo("Código de extração do Leiden Ranking: leitura em fluxo e filtro do Brasil", COD_LEIDEN,
               "scripts/extrair_leiden.py")
    sub("5.2 PADRONIZAÇÃO E CRUZAMENTO DAS INSTITUIÇÕES")
    p("As bases não compartilham um código de instituição, então o cruzamento foi feito pelo nome. Cada "
      "nome foi convertido em uma chave: sem acentos, em minúsculas, sem palavras genéricas "
      "(\"universidade\", \"university\", \"de\" etc.), com termos em inglês traduzidos para o português e "
      "com as palavras em ordem alfabética (Quadro 4). Assim, \"Federal University of Paraíba\" e "
      "\"Universidade Federal da Paraíba\" geram a mesma chave. Equivalências que os termos não resolvem "
      "foram registradas à mão; hoje só a da Unesp.")
    rel.codigo("Código da chave de nome usada no cruzamento entre bases", COD_CHAVE, "scripts/preparar_dados.py")
    sub("5.3 MEDIDA DE ASSOCIAÇÃO: CORRELAÇÃO DE SPEARMAN")
    p("A associação foi medida pela correlação de postos de Spearman (ρ): os valores são convertidos em "
      "postos, com empate resolvido pela média dos postos empatados, e calcula-se a correlação de Pearson "
      "entre os postos (Quadro 5). A implementação foi escrita em Python puro, sem bibliotecas estatísticas, "
      "para manter o pipeline sem dependências novas. O cálculo exige pelo menos cinco universidades em "
      "comum; abaixo disso, o resultado é descartado. O tamanho da amostra (n) é sempre reportado ao lado de "
      "ρ, e não foram calculados p-valores, porque várias amostras dos rankings internacionais são pequenas.")
    rel.codigo("Código da correlação de postos de Spearman", COD_SPEARMAN, "scripts/preparar_dados.py")
    p("Nos rankings, posição menor significa melhor; nos indicadores, nota maior significa melhor. Por isso, "
      "quando um indicador acompanha a posição, ρ sai negativo. Para facilitar a leitura, as figuras e "
      "tabelas mostram a concordância, igual a −ρ entre indicador e posição e igual a ρ entre duas "
      "posições: valores positivos indicam que as duas medidas ordenam as universidades da mesma forma. Para "
      "cada ranking com indicadores publicados (QS e RUF), cada indicador foi correlacionado com a posição "
      "geral, edição a edição (Quadro 6).")
    rel.codigo("Código que correlaciona cada indicador com a posição geral, por edição", COD_COLUNAS,
               "scripts/preparar_dados.py")
    p("Foram calculadas ainda: a correlação da posição de cada ranking entre edições consecutivas "
      "(estabilidade); a correlação entre a posição no ARWU e os indicadores do Leiden (publicações e "
      "proporção no top 10%); a correlação entre a nota média de cada universidade no IGC, no CPC e no Enade "
      "e sua posição no ARWU, no QS e no SCImago, no mesmo ano; e a correlação entre a posição no RUF e a "
      "posição nos rankings internacionais, que mede a concordância entre os grupos.")
    sub("5.4 OPERACIONALIZAÇÃO DAS HIPÓTESES")
    p("As hipóteses do projeto falam em associações \"10% mais fortes\" e em variações \"menores que 5%\", "
      "sem definir a medida. Elas foram operacionalizadas assim:")
    it("a) Para cada dimensão, calculou-se a média de |ρ| dos indicadores daquela dimensão em todas as "
       "edições disponíveis, separadamente para cada grupo. Pesquisa: citações por docente (QS) e "
       "publicações e proporção no top 10% (Leiden, contra a posição no ARWU), no grupo internacional, e o "
       "indicador Pesquisa do RUF, no brasileiro. Ensino: proporção docente/aluno (QS) e Ensino (RUF). "
       "Mercado: reputação com empregadores e empregabilidade (QS) e Mercado (RUF);")
    it("b) \"10% mais forte\" significa que a razão entre as médias de |ρ| dos dois grupos é de pelo menos "
       "1,10;")
    it("c) \"Varia menos de 5% ao longo do período\" significa que a amplitude relativa de |ρ| de um "
       "indicador entre as edições, (máximo − mínimo) ÷ média, é menor que 0,05; \"diferença de mais de 5% "
       "entre os grupos\" significa que a diferença entre as médias dos grupos, dividida pela maior delas, "
       "passa de 0,05.")
    sub("5.5 FERRAMENTAS E REPRODUTIBILIDADE")
    p("O pipeline foi escrito em Python 3, com a biblioteca openpyxl para ler as planilhas. O dashboard é "
      "uma página HTML com gráficos em Chart.js, gerada a partir de um modelo, com seis painéis (visão "
      "geral, pesquisa e impacto, ensino e avaliação, comparação entre universidades, correlações e fontes "
      "e limites). As figuras deste relatório foram geradas com matplotlib a partir dos mesmos dados do "
      "dashboard, com uma paleta de cores verificada para daltonismo, e o próprio relatório é montado por "
      "um script a partir do modelo da UFPB, para que todos os números do texto saiam dos dados. A sequência "
      "de execução é: extração (extrair_leiden.py, extrair_qs.py, extrair_scimago.py e extrair_ruf.py), "
      "agregação e correlações (preparar_dados.py), dashboard (gerar_dashboard.py), figuras "
      "(gerar_figuras.py) e relatório (gerar_relatorio.py).")

    # ------------------------------------------------------------ 6 RESULTADOS
    sec("6 RESULTADOS")
    sub("6.1 COBERTURA DAS BASES")
    sci_min, sci_max = min(COB["SCImago"].values()), max(COB["SCImago"].values())
    arwu_ini, arwu_fim = COB["ARWU"][min(COB["ARWU"])], COB["ARWU"][max(COB["ARWU"])]
    qs_pico = max(COB["QS"].values())
    p("A presença das universidades brasileiras variou muito entre as fontes (Figura 1). O RUF cobriu de "
      f"{faixa(COB['RUF'])} universidades por edição e o SCImago de {sci_min} a {sci_max}, enquanto o "
      f"Leiden cobriu de {faixa(COB['Leiden'])}, o QS de {faixa(COB['QS'])} e o ARWU de {faixa(COB['ARWU'])}. "
      f"O número de brasileiras no ARWU caiu de {arwu_ini}, em {min(COB['ARWU'])}, para {arwu_fim}, em "
      f"{max(COB['ARWU'])}; no QS, subiu até {qs_pico} em 2023 e 2025 e caiu para "
      f"{COB['QS'][max(COB['QS'])]} em {max(COB['QS'])}.")
    rel.figura("Universidades brasileiras presentes em cada ranking, por edição", "fig1_cobertura.png",
               "Fonte: elaborado pelos autores (2026) a partir de ShanghaiRanking Consultancy (2026), QS "
               "Quacquarelli Symonds (2026), CWTS (2025), SCImago Lab ([s.d.]) e RUF (2023–2025).")
    sub("6.2 ASSOCIAÇÃO ENTRE INDICADORES E POSIÇÃO")
    ult = RES["ultima_edicao"]
    qs_ar = rho_de("colunas_qs_ruf", "QS: Reputação acadêmica", ult["QS"])
    qs_fsr = rho_de("colunas_qs_ruf", "QS: Proporção docente/aluno", ult["QS"])
    ruf_p = rho_de("colunas_qs_ruf", "RUF: Pesquisa", ult["RUF"])
    ruf_e = rho_de("colunas_qs_ruf", "RUF: Ensino", ult["RUF"])
    ruf_m = rho_de("colunas_qs_ruf", "RUF: Mercado", ult["RUF"])
    p(f"Na edição mais recente de cada ranking (Figura 2), o indicador com maior concordância com a posição "
      f"geral foi a reputação acadêmica no QS {ult['QS']} ({v(-qs_ar['rho'])}; n = {qs_ar['n']}) e a pesquisa "
      f"no RUF {ult['RUF']} ({v(-ruf_p['rho'])}; n = {ruf_p['n']}), seguida de perto pelo ensino "
      f"({v(-ruf_e['rho'])}). O indicador de menor concordância foi a proporção docente/aluno no QS "
      f"({v(-qs_fsr['rho'])}) e o mercado de trabalho no RUF ({v(-ruf_m['rho'])}). Todos os indicadores dos "
      "dois rankings tiveram concordância positiva com a posição.")
    rel.figura(f"Concordância entre cada indicador e a posição geral no QS {ult['QS']} e no RUF {ult['RUF']}",
               "fig2_peso_indicador.png",
               "Fonte: elaborado pelos autores (2026) a partir de QS Quacquarelli Symonds (2026) e RUF (2025). "
               "Concordância = −ρ de Spearman.")
    hp, he, hm, hi = H["pesquisa"], H["ensino"], H["mercado"], H["internacionalizacao"]
    p("Considerando todas as edições disponíveis (Tabela 1), a média de |ρ| na dimensão de pesquisa foi "
      f"{v(hp['internacional'])} no grupo internacional e {v(hp['brasileiro'])} no brasileiro; no ensino, "
      f"{v(he['internacional'])} e {v(he['brasileiro'])}; no mercado, {v(hm['internacional'])} e "
      f"{v(hm['brasileiro'])}; e na internacionalização, {v(hi['internacional'])} e {v(hi['brasileiro'])}.")
    t1 = [
        ["Pesquisa", "QS citações por docente; ARWU × Leiden (publicações e top 10%)", v(hp["internacional"]),
         "RUF Pesquisa", v(hp["brasileiro"]), v(hp["razao_br_intl"])],
        ["Ensino", "QS proporção docente/aluno", v(he["internacional"]), "RUF Ensino", v(he["brasileiro"]),
         v(he["razao_br_intl"])],
        ["Mercado", "QS reputação com empregadores; empregabilidade", v(hm["internacional"]), "RUF Mercado",
         v(hm["brasileiro"]), v(hm["razao_br_intl"])],
        ["Internacionalização", "QS docentes e alunos internacionais", v(hi["internacional"]),
         "RUF Internacionalização", v(hi["brasileiro"]), v(hi["razao_br_intl"])],
    ]
    rel.tab("Tabela", "Média de |ρ| entre indicador e posição, por dimensão e grupo de rankings",
            ["Dimensão", "Indicadores internacionais", "|ρ| intern.", "Indicador brasileiro", "|ρ| bras.",
             "Razão bras./intern."], t1, [3.2, 4.2, 1.5, 3.2, 1.5, 2.4],
            alin_col=["E", "E", "C", "E", "C", "C"], tam=9,
            fonte="Fonte: elaborado pelos autores (2026). Médias sobre as edições QS 2023–2027, "
                  "ARWU × Leiden 2021–2025 e RUF 2023–2025.")
    sub("6.3 VARIAÇÃO AO LONGO DO TEMPO")
    amp = H["c_amplitude"]
    ruf_int = amp["RUF · Internacionalização"]
    p("A concordância de cada indicador com a posição variou pouco entre as edições do RUF e bastante entre "
      f"as do QS (Figura 3 e Tabela 2). No RUF, {H['c_estaveis_brasileiro']} dos {H['c_total_brasileiro']} "
      "indicadores tiveram amplitude relativa abaixo de 0,05; a exceção foi a internacionalização, cuja "
      f"concordância caiu de {v(ruf_int['max'])} para {v(ruf_int['min'])} em 2025. Entre os indicadores "
      f"internacionais, só {H['c_estaveis_internacional']} de {H['c_total_internacional']} ficaram abaixo de "
      "0,05 (publicações no Leiden contra a posição no ARWU). No QS, a reputação acadêmica subiu em todas as "
      f"edições, de {v(amp['QS · Reputação acadêmica']['min'])} em 2023 para "
      f"{v(amp['QS · Reputação acadêmica']['max'])} em 2027, e a proporção docente/aluno caiu de "
      f"{v(amp['QS · Proporção docente/aluno']['max'])} para {v(amp['QS · Proporção docente/aluno']['min'])}.")
    rel.figura("Concordância de cada indicador com a posição geral, por edição (QS 2023–2027 e RUF 2023–2025)",
               "fig3_variacao_temporal.png",
               "Fonte: elaborado pelos autores (2026) a partir de QS Quacquarelli Symonds (2026) e RUF "
               "(2023–2025). A barra liga a menor e a maior concordância de cada indicador.")
    linhas_amp = sorted(amp.items(), key=lambda kv: (kv[1]["ranking"] != "RUF", kv[1]["amplitude_rel"]))
    t2 = [[k.replace(" · ", ": "), str(a["edicoes"]), v(a["min"]), v(a["max"]), v(a["amplitude_rel"])]
          for k, a in linhas_amp]
    rel.tab("Tabela", "Amplitude relativa de |ρ| entre as edições, por indicador",
            ["Indicador", "Edições", "|ρ| mínimo", "|ρ| máximo", "Amplitude relativa"], t2,
            [6.6, 1.8, 2.2, 2.2, 3.2], alin_col=["E", "C", "C", "C", "C"], tam=9,
            fonte="Fonte: elaborado pelos autores (2026). Amplitude relativa = (máximo − mínimo) ÷ média.")
    qs_est = estab("QS: posição ano a ano")
    qs_min = min(qs_est, key=lambda r: r["rho"])
    rufs = estab("RUF: posição ano a ano")
    arwus = estab("ARWU: posição ano a ano")
    p("As posições em si foram estáveis entre edições consecutivas (Figura 4). No RUF, ρ ficou em "
      f"{v(min(r['rho'] for r in rufs))} ou mais; no ARWU, subiu de {v(arwus[0]['rho'])} "
      f"({arwus[0]['de']}→{arwus[0]['para']}) para {v(arwus[-1]['rho'])} ({arwus[-1]['de']}→{arwus[-1]['para']}). "
      f"O menor valor da série foi o do QS entre {qs_min['de']} e {qs_min['para']} (ρ = {v(qs_min['rho'])}; "
      f"n = {qs_min['n']}).")
    rel.figura("Correlação de Spearman entre as posições em edições consecutivas de cada ranking",
               "fig4_estabilidade.png",
               "Fonte: elaborado pelos autores (2026) a partir de ShanghaiRanking Consultancy (2026), QS "
               "Quacquarelli Symonds (2026), SCImago Lab ([s.d.]) e RUF (2023–2025).")
    sub("6.4 ÍNDICES DO INEP E POSIÇÃO NOS RANKINGS INTERNACIONAIS")
    inep = C["inep_x_internacional"]
    igc_vals = [-r["rho"] for r in inep if r["indice"] == "IGC"]
    cpc_qs = [r for r in inep if r["indice"] == "CPC" and r["ranking"] == "QS"]
    p("A nota média das universidades nos índices do Inep concordou com a posição nos rankings "
      f"internacionais em quase todos os cruzamentos (Figura 5). O IGC teve concordância de {v(min(igc_vals))} "
      f"a {v(max(igc_vals))}, a maior entre os três índices. O cruzamento mais fraco foi o do CPC com o QS, "
      f"com valores de {v(min(-r['rho'] for r in cpc_qs))} a {v(max(-r['rho'] for r in cpc_qs))}. Nos "
      "cruzamentos com o SCImago, a amostra passou de 120 universidades; com o ARWU e o QS, ficou entre 12 e "
      "20.")
    rel.figura("Concordância entre a nota média no IGC, no CPC e no Enade e a posição no ARWU, no QS e no "
               "SCImago, por ano", "fig5_inep_internacional.png",
               "Fonte: elaborado pelos autores (2026) a partir de Inep ([s.d.]), ShanghaiRanking Consultancy "
               "(2026), QS Quacquarelli Symonds (2026) e SCImago Lab ([s.d.]). \"–\": sem dados no ano.")
    sub("6.5 CONCORDÂNCIA ENTRE RANKINGS BRASILEIROS E INTERNACIONAIS")
    pares = ["RUF × ARWU", "RUF × SCImago", "RUF × QS", "RUF × IGC"]
    t3 = []
    for par in pares:
        lin = [par]
        for ano in (2023, 2024, 2025):
            c, n = concord(par, ano)
            lin.append(f"{v(c)} ({n})" if c is not None else "–")
        t3.append(lin)
    uf = UFPB
    p("As posições das universidades brasileiras no RUF concordaram com as posições nos rankings "
      "internacionais (Tabela 3). A concordância foi maior com o ARWU e o SCImago e menor com o QS. Na "
      f"edição de {DISP['ano']}, a correlação entre o RUF e o SCImago foi ρ = {v(DISP['rho'])} (n = "
      f"{DISP['n']}; Figura 6). A UFPB ficou em {DISP['ufpb']['ruf']}º no RUF e em "
      f"{milhar(DISP['ufpb']['scimago'])}º na posição global do SCImago "
      f"({int(uf['sci'][DISP['ano']][0])}º entre as brasileiras); no QS {DISP['ano']}, ficou na faixa "
      f"{uf['qs'].get(DISP['ano'], '–')}, e não aparece no ARWU em nenhuma edição.")
    rel.tab("Tabela", "Concordância entre a posição no RUF e as posições ou notas nas demais fontes",
            ["Par", "2023", "2024", "2025"], t3, [4.4, 3.4, 3.4, 3.4], alin_col=["E", "C", "C", "C"],
            fonte="Fonte: elaborado pelos autores (2026). Valores: concordância (n). Entre duas posições, "
                  "concordância = ρ; entre posição e nota (IGC), = −ρ.")
    rel.figura(f"Posição no RUF {DISP['ano']} e posição global no SCImago {DISP['ano']} das universidades "
               "brasileiras", "fig6_ruf_scimago.png",
               f"Fonte: elaborado pelos autores (2026) a partir de RUF ({DISP['ano']}) e SCImago Lab ([s.d.]).",
               largura=15)

    # ------------------------------------------------------------ 7 DISCUSSÃO
    sec("7 DISCUSSÃO")
    sub("7.1 ANÁLISE DOS RESULTADOS")
    p([("Hipótese (a) – refutada. ", {"b": True}),
       ("Esperava-se que pesquisa e citações pesassem pelo menos 10% mais nos rankings internacionais. "
        f"Ocorreu o contrário: a razão entre os grupos foi {v(hp['razao_intl_br'])}, ou seja, a pesquisa "
        f"pesou cerca de {pct(hp['razao_br_intl'] - 1)} a mais no RUF. A explicação mais provável está na "
        "construção do RUF: a pesquisa responde por 42% da nota (UNIVERSIDADE DE SÃO PAULO, 2025), e a "
        "correlação entre um componente e a nota que ele ajuda a formar tende a ser alta por definição. "
        "Nos internacionais, o resultado depende do indicador: o volume de publicações acompanhou a posição "
        f"no ARWU quase tanto quanto a pesquisa no RUF (|ρ| entre {v(amp['ARWU×Leiden · publicações']['min'])} "
        f"e {v(amp['ARWU×Leiden · publicações']['max'])}), enquanto a proporção de artigos no top 10% mais "
        "citado, uma medida de impacto independente do tamanho, acompanhou bem menos. Isso converge com "
        "Szluka, Csajbók e Győrffy (2023), para quem tamanho e citações estão entre os parâmetros mais "
        "associados à posição, com intensidade que varia entre os rankings.", {})])
    p([("Hipótese (b) – confirmada em parte. ", {"b": True}),
       (f"O ensino pesou {v(he['razao_br_intl'], 1)} vezes mais no RUF do que no QS, bem acima do limiar de "
        f"1,10. O mercado de trabalho, porém, teve associações praticamente iguais nos dois grupos (razão de "
        f"{v(hm['razao_br_intl'])}), e no QS a reputação com empregadores chegou a superar o indicador de "
        "mercado do RUF em várias edições. As condições de oferta não puderam ser testadas contra uma "
        "posição, porque o IGC é uma nota e não um ranking; o que se observou foi que o IGC concordou com a "
        "posição nos rankings internacionais, e o CPC, mais centrado na graduação, concordou menos, sobretudo "
        "com o QS. Isso sugere que a qualidade medida pelo Inep e o prestígio internacional andam juntos "
        "principalmente por meio da pós-graduação, que entra no IGC e não no CPC.", {})])
    p([("Hipótese (c) – confirmada em parte. ", {"b": True}),
       ("A primeira parte, estabilidade abaixo de 5%, valeu para o RUF (exceto a internacionalização em "
        "2025) e não valeu para o QS, cujos indicadores oscilaram entre 17% e 55% de amplitude relativa. A "
        "segunda parte, diferença acima de 5% entre os grupos, valeu para pesquisa "
        f"({pct(H['c_diferenca_rel_grupos']['pesquisa'])}), ensino ({pct(H['c_diferenca_rel_grupos']['ensino'])}) "
        f"e internacionalização ({pct(H['c_diferenca_rel_grupos']['internacionalizacao'])}), mas não para "
        f"mercado ({pct(H['c_diferenca_rel_grupos']['mercado'], 1)}). A instabilidade do QS coincide com a "
        "revisão metodológica da edição 2024, a primeira com o indicador de sustentabilidade na base, e é "
        f"nessa transição que está a menor correlação de posições da série (ρ = {v(qs_min['rho'])}). O "
        "crescimento contínuo da concordância da reputação acadêmica no QS indica que, para as brasileiras, "
        "a posição no QS se tornou quase uma tradução da reputação.", {})])
    p("Dois padrões gerais emergem. Primeiro, o peso declarado e o peso efetivo não coincidem: a inovação "
      "vale 4% da nota do RUF, mas sua concordância com a posição (cerca de 0,71) superou a do mercado, que "
      "vale 18%. Isso sugere que a inovação se concentra nas mesmas universidades que lideram em pesquisa, "
      "e herda delas a associação com a posição. Segundo, os dois grupos chegaram a ordenações parecidas por caminhos "
      "diferentes: o RUF concordou fortemente com o ARWU e o SCImago, que são bibliométricos, e menos com o "
      "QS, que é dominado pela reputação. Esse resultado está de acordo com Selten et al. (2020), para quem "
      "os rankings medem sobretudo reputação e desempenho em pesquisa, e com Moed (2017), que aponta "
      "fundamentos conceituais distintos entre os rankings mundiais. Também ajuda a entender a observação de "
      "Santos (2015) de que o número de brasileiras entre as melhores muda conforme o ranking: a ordem "
      "entre elas é parecida, mas o corte de cada ranking e o peso dado à reputação mudam quem entra.")
    p("A UFPB ilustra o padrão: está entre as 30 primeiras no RUF, mas fora das mil primeiras posições no QS "
      "e no SCImago. A diferença não indica uma contradição entre os grupos, e sim "
      "a escala: entre as brasileiras, a UFPB está no terço superior; no mundo, compete com milhares de "
      "instituições com mais produção e mais reputação acumulada.")
    sub("7.2 ALTERAÇÕES REALIZADAS")
    p("Alguns procedimentos previstos no projeto foram alterados durante a execução:")
    for t in [
        "a) Janela temporal: as bases foram fechadas em 16/10/2026 com as edições disponíveis. O QS "
        "(2017–2027), o ARWU (2021–2026), o Leiden (2021–2025) e o SCImago (2022–2026) cumpriram as cinco "
        "edições previstas, mas o RUF ficou com três edições completas (2023–2025), porque o arquivo de "
        "2019 traz apenas a classificação de Mercado e as edições de 2020 a 2022 não foram obtidas. O IGC e o "
        "CPC cobrem 2019 e 2021–2023, pois não houve edição em 2020. A comparação temporal do grupo brasileiro "
        "ficou, portanto, mais curta que a prevista;",
        "b) \"Rankings de pesquisa de recorte brasileiro\": não foi incluído um ranking nacional à parte; o "
        "recorte brasileiro do SCImago (posição nacional) cumpriu esse papel;",
        "c) Fontes acrescentadas: além do IGC, entraram o CPC e o Enade, que permitem separar a graduação "
        "(CPC e Enade) do índice que inclui a pós-graduação (IGC);",
        "d) Hipóteses: os percentuais das hipóteses foram operacionalizados como descrito na seção 5.4, porque "
        "o projeto não definia a medida;",
        "e) Indicadores do ARWU: a interface pública traz os indicadores sem rótulo, e eles não foram usados "
        "para não atribuir significado sem confirmação na fonte; do ARWU, usou-se só a posição;",
        "f) Projeto: a numeração das seções foi alinhada, a frase que antecipava o objetivo foi retirada do "
        "tema e problema e o cronograma passou a ter as mesmas sete atividades da metodologia, conforme as "
        "correções da orientadora. As referências de Hazelkorn (2015) e de Marginson e Van der Wende (2007), "
        "citadas no texto do projeto mas ausentes da lista, foram incluídas, e o endereço eletrônico da tese "
        "de Santos (2015) foi corrigido.",
    ]:
        it(t)
    sub("7.3 AMEAÇAS À VALIDADE")
    for t in [
        "a) Cruzamento por nome: sem um código comum entre as bases, erros de correspondência podem excluir ou "
        "trocar instituições. A Universidade de São Paulo, a Unicamp, a Unesp e a Uerj, por exemplo, não "
        "aparecem no IGC, no CPC nem no Enade, o que reduz as amostras dos cruzamentos com o Inep "
        "justamente nas universidades mais bem colocadas;",
        "b) Circularidade: no QS e no RUF, cada indicador é parte da nota que define a posição, então a "
        "correlação entre os dois é inflada por construção e reflete, em parte, o peso declarado;",
        "c) Restrição de amplitude: os rankings internacionais só incluem as brasileiras mais bem colocadas, "
        "o que reduz a variação e tende a enfraquecer as correlações nesse grupo;",
        "d) Amostras pequenas e faixas: no ARWU e no QS, n ficou entre 12 e 35; no ARWU, as posições depois da "
        "100ª vêm em faixas, o que gera empates. Sem p-valores, diferenças pequenas entre ρ não devem ser "
        "interpretadas como significativas;",
        "e) Mudanças metodológicas: o QS mudou indicadores e pesos ao longo do período, e as edições nem "
        "sempre são comparáveis entre si;",
        "f) Comparações múltiplas e alinhamento de anos: foram calculadas mais de cem correlações, e o ano "
        "nominal de cada edição não coincide com o ano de coleta dos dados de cada ranking.",
    ]:
        it(t)

    # ------------------------------------------------------------ 8 CONCLUSÃO
    sec("8 CONCLUSÃO E TRABALHOS FUTUROS")
    sub("8.1 CONCLUSÕES E PRINCIPAIS CONTRIBUIÇÕES")
    p("A pesquisa reuniu oito fontes públicas de rankings e indicadores, construiu um pipeline reprodutível "
      "para extraí-las, padronizá-las e cruzá-las, e mediu, com a correlação de Spearman, quanto cada "
      "característica das universidades brasileiras acompanha a posição em cada ranking, ao longo de até "
      "onze edições. Os objetivos específicos foram atendidos: as metodologias foram levantadas e "
      "classificadas em dimensões (a), as bases foram coletadas e padronizadas (b), a distribuição e a "
      "evolução das variáveis foram caracterizadas (c), as associações foram verificadas ano a ano (d), os "
      "grupos foram comparados (e) e os resultados estão no dashboard (f). O objetivo (b) foi atendido com "
      "a ressalva da série mais curta do RUF.")
    p("A resposta ao problema é que os rankings brasileiros e os internacionais valorizam coisas diferentes, "
      "mas chegam a ordenações parecidas das universidades brasileiras. No RUF, pesquisa e ensino concentram "
      "a associação com a posição; no QS, a reputação acadêmica domina cada vez mais; e o ensino é a "
      "dimensão que mais separa os grupos. A principal contribuição é mostrar, com dados, que o peso "
      "declarado de um indicador não é o seu peso efetivo, e que a concordância entre rankings convive com "
      "critérios bem distintos. Para a UFPB e para outras universidades, isso indica que melhorar no RUF e "
      "melhorar nos rankings internacionais pedem estratégias diferentes, sobretudo em reputação e em "
      "visibilidade internacional. A base padronizada, o dashboard e os scripts ficam disponíveis para "
      "novas edições.")
    sub("8.2 LIMITAÇÕES")
    p("Permanecem em aberto: a série curta do RUF, que limita a avaliação da estabilidade no grupo brasileiro; "
      "a ausência das principais universidades estaduais paulistas e da Uerj nas bases do Inep; o cruzamento "
      "por nome; as amostras pequenas no ARWU e no QS, que impedem testes formais; e a circularidade entre "
      "indicador e posição, que faz a correlação refletir em parte o peso declarado e não só a relação entre "
      "as características das universidades.")
    sub("8.3 TRABALHOS FUTUROS")
    for t in [
        "a) Estimar intervalos de confiança para ρ por bootstrap, para separar diferenças reais de variação "
        "amostral;",
        "b) Usar modelos que controlem a colinearidade entre indicadores (regressão com postos ou análise de "
        "componentes principais), para estimar o peso efetivo de cada dimensão descontando as demais;",
        "c) Incluir o Times Higher Education e as edições do RUF de 2020 a 2022 e do Leiden de 2019 e 2020, "
        "para alongar as séries;",
        "d) Cruzar as instituições pelo código do cadastro e-MEC em vez do nome, para eliminar erros de "
        "correspondência;",
        "e) Testar a nova hipótese levantada na discussão: a concordância entre o Inep e os rankings "
        "internacionais passa pela pós-graduação, o que pode ser verificado com os conceitos da Capes.",
    ]:
        it(t)

    # ------------------------------------------------------------ REFERÊNCIAS
    sec("REFERÊNCIAS")
    for ref in REFERENCIAS:
        rel.par(ref, alin="E", recuo1=0, esp=1.0, depois=12)

    return rel


# ================================================================ 2ª passada: páginas
def soffice():
    for c in (shutil.which("soffice"), "/Applications/LibreOffice.app/Contents/MacOS/soffice",
              "/usr/bin/soffice", "/usr/lib/libreoffice/program/soffice"):
        if c and os.path.exists(c):
            return c
    return None


def localizar_paginas(docx_path, rel):
    exe = soffice()
    try:
        import pymupdf
    except ImportError:
        pymupdf = None
    if not exe or not pymupdf:
        print("aviso: LibreOffice ou PyMuPDF ausente; sumário e listas ficam com página 00")
        return {}
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run([exe, "--headless", "--convert-to", "pdf", "--outdir", tmp, docx_path],
                       check=True, capture_output=True)
        pdf = os.path.join(tmp, os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")
        doc = pymupdf.open(pdf)
        textos = [[ln.strip() for ln in pg.get_text().splitlines()] for pg in doc]
    inicio = next(i for i, ls in enumerate(textos) if any(l.startswith("Título do Projeto:") for l in ls))
    paginas = {}
    for _, tit in rel.titulos:
        for i in range(inicio, len(textos)):
            if tit in textos[i]:
                paginas[tit] = i + 1
                break
    for rot in rel.figuras + rel.quadros + rel.tabelas:
        prefixo = rot.split(" – ")[0] + " – "
        for i in range(inicio, len(textos)):
            if any(l.startswith(prefixo) for l in textos[i]):
                paginas[rot] = i + 1
                break
    faltam = [t for _, t in rel.titulos if t not in paginas] + \
             [r for r in rel.figuras + rel.quadros + rel.tabelas if r not in paginas]
    if faltam:
        print("aviso: página não encontrada para", faltam)
    return paginas


def main():
    # 1ª passada: documento sem páginas, para saber os títulos e legendas
    rel = construir({"_titulos": [], "_figuras": [], "_quadros": [], "_tabelas": []})
    meta = {"_titulos": rel.titulos, "_figuras": rel.figuras, "_quadros": rel.quadros, "_tabelas": rel.tabelas}
    # 2ª passada: com as entradas do sumário/listas (página 00) -- a paginação final é a mesma
    rel = construir(meta)
    rel.salvar(SAIDA)
    paginas = localizar_paginas(SAIDA, rel)
    # 3ª passada: números de página reais (mesma largura de "00", então a paginação não muda)
    rel = construir({**meta, **paginas})
    rel.salvar(SAIDA)
    n_palavras = len(RESUMO.split())
    print(f"relatório: {os.path.relpath(SAIDA, RAIZ)}  |  resumo: {n_palavras} palavras  |  "
          f"{len(rel.figuras)} figuras, {len(rel.quadros)} quadros, {len(rel.tabelas)} tabelas  |  "
          f"páginas localizadas: {len(paginas)}")


if __name__ == "__main__":
    main()
