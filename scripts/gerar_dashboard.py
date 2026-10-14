"""Gera dashboard_rankings.html embutindo dados/processado/dados_dashboard.json no modelo."""
import os

AQUI = os.path.dirname(os.path.abspath(__file__))
RAIZ = os.path.join(AQUI, "..")
modelo = open(os.path.join(AQUI, "dashboard_template.html"), encoding="utf-8").read()
dados = open(os.path.join(RAIZ, "dados", "processado", "dados_dashboard.json"), encoding="utf-8").read()
assert "__DADOS__" in modelo
# evita fechar a tag <script> por acidente se algum nome contiver "</"
html = modelo.replace("__DADOS__", dados.replace("</", "<\\/"))
saida = os.path.join(RAIZ, "dashboard_rankings.html")
open(saida, "w", encoding="utf-8").write(html)
print("salvo", os.path.abspath(saida), round(len(html) / 1024), "KB")
