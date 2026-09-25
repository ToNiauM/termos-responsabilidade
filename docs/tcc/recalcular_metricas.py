# -*- coding: utf-8 -*-
"""
Recalcula os indicadores da campanha censitária a partir do relatório exportado
pelo sistema de inventário e gera as figuras do esqueleto 11 em 11_figuras/.

Uso:  .venv/bin/python 11_make_figuras.py
Entrada: relatorio_20260922_211049.xlsx (posição de 22 set. 2026)
Saída:   11_figuras/*.png e 11_metricas.json (conferência dos números do texto)
"""
import json, re, shutil
from pathlib import Path
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

ROOT = Path(__file__).resolve().parent
XLSX = ROOT / "relatorio_20260922_211049.xlsx"   # export do sistema de conferência; não versionado (traz nomes de servidores)
OLD = ROOT / "10_figuras"
OUT = ROOT / "11_figuras"
OUT.mkdir(exist_ok=True)
HOJE = pd.Timestamp("2026-09-22")
GREY = "#666666"
DARK = "#333333"
LIGHT = "#b5b5b5"

plt.rcParams.update({
    "font.family": "sans-serif",
    "font.sans-serif": ["Arial", "Liberation Sans", "DejaVu Sans"],
    "font.size": 11, "text.color": "black", "axes.labelcolor": "black",
    "xtick.color": "black", "ytick.color": "black", "axes.edgecolor": "black",
})

def clean(ax, left=True, bottom=True):
    ax.grid(False)
    for s in ("top", "right"): ax.spines[s].set_visible(False)
    ax.spines["left"].set_visible(left); ax.spines["bottom"].set_visible(bottom)
    for s in ("left", "bottom"):
        if ax.spines[s].get_visible():
            ax.spines[s].set_linewidth(1.2)

def save(fig, name):
    fig.tight_layout()
    fig.savefig(OUT / name, dpi=300, bbox_inches="tight", facecolor="white")
    plt.close(fig)

def br(v, d=0):
    s = f"{v:,.{d}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")

# ------------------------------------------------------------------ dados
df = pd.read_excel(XLSX, header=4)
df.columns = ["foto", "pat", "sit", "desc", "compl", "classe", "loc_sis", "loc_inv",
              "dt_ent", "v_compra", "v_atual", "conserv", "integrante", "dh", "usuario"]
df["dh"] = pd.to_datetime(df["dh"], format="%d/%m/%Y %H:%M:%S", errors="coerce")
df["dt_ent"] = pd.to_datetime(df["dt_ent"], format="%d/%m/%Y", errors="coerce")
def money(s):
    return pd.to_numeric(s.astype(str).str.replace(".", "", regex=False)
                         .str.replace(",", ".", regex=False), errors="coerce")
df["valor"] = money(df["v_atual"])
df["loc_sis_n"] = df.loc_sis.fillna("").str.strip().str.upper()
df["loc_inv_n"] = df.loc_inv.fillna("").str.strip().str.upper()

def pav(loc):
    m = re.match(r"^(S[123]|T|\d{2})\s*-", loc or "")
    if not m: return "Localizações lógicas"
    p = m.group(1)
    if p.startswith("S"): return f"Subsolo {p[1]}"
    if p == "T": return "Térreo"
    return f"{int(p)}º pavimento"
df["pav"] = df.loc_sis_n.map(pav)
df["pav_inv"] = df.loc_inv_n.map(pav)

IMOVEIS = {"SEDE", "TERRENOS", "INSTALAÇÕES"}
at = df[df.sit == "ATIVO"].copy()
chk = df[df.dh.notna()].copy()
atc = at[at.dh.notna()].copy()
atc["ok"] = atc.loc_sis_n == atc.loc_inv_n
chk["dia"] = chk.dh.dt.date
atc["dia"] = atc.dh.dt.date

def etapa(d):
    if d <= pd.Timestamp("2026-08-21").date(): return 1
    if d <= pd.Timestamp("2026-08-28").date(): return 2
    return 3
chk["etapa"] = chk.dia.map(etapa)
atc["etapa"] = atc.dia.map(etapa)

M = {}
M["registros"] = len(df)
M["situacao"] = df.sit.value_counts().to_dict()
M["ativos"] = len(at)
M["localizacoes"] = int(at.loc_sis_n.replace("", np.nan).nunique())
M["classes_ativos"] = at.classe.value_counts().to_dict()
M["valor_ativos_M"] = round(at.valor.sum() / 1e6, 2)
mov = at[~at.classe.isin(IMOVEIS)].copy()
M["moveis_n"] = len(mov)
M["valor_moveis_M"] = round(mov.valor.sum() / 1e6, 2)
mov["idade"] = (HOJE - mov.dt_ent).dt.days / 365.25
M["idade_media"] = round(mov.idade.mean(), 1)
M["idade_mediana"] = round(mov.idade.median(), 1)
M["pct_mais_10_anos"] = round((mov.idade > 10).mean() * 100, 1)
M["n_desde_2020"] = int((mov.dt_ent.dt.year >= 2020).sum())
M["n_antes_2000"] = int((mov.dt_ent.dt.year < 2000).sum())

# ------------------------------------------------------------------ campanha
M["conferidos"] = len(chk)
M["conferidos_sit"] = chk.sit.value_counts().to_dict()
M["ativos_conferidos"] = len(atc)
M["cobertura"] = round(len(atc) / len(at) * 100, 1)
M["pendentes"] = len(at) - len(atc)
M["dias_campo"] = int(chk.dia.nunique())
M["servidores"] = int(chk.integrante.nunique())
top2 = chk.integrante.value_counts().head(2).sum()
M["pct_dois_servidores"] = round(top2 / len(chk) * 100, 1)
M["conservacao"] = chk.conserv.value_counts().to_dict()
M["concordantes"] = int(atc.ok.sum()); M["divergentes"] = int((~atc.ok).sum())
M["concordancia_pct"] = round(atc.ok.mean() * 100, 1)
M["divergencia_pct"] = round((~atc.ok).mean() * 100, 1)
movc = mov[mov.dh.notna()]
M["cobertura_valor_moveis_pct"] = round(movc.valor.sum() / mov.valor.sum() * 100, 1)
M["valor_moveis_conferido_M"] = round(movc.valor.sum() / 1e6, 2)

# produtividade por lote (unidade: lote de sincronização = mesmo instante)
def prod_rows(g):
    lotes = g.groupby("dh").size().sort_index()
    gap = lotes.index.to_series().diff().dt.total_seconds()
    rows = []
    for (ts, n), gp in zip(lotes.items(), gap):
        rows.append(dict(ts=ts, n=int(n), gap=gp, medido=bool(pd.notna(gp) and gp <= 3600)))
    return rows
diario = []
for dia, g in chk.groupby("dia"):
    rows = prod_rows(g)
    med = [r for r in rows if r["medido"]]
    tempo = sum(r["gap"] for r in med) / 3600
    bens = sum(r["n"] for r in med)
    diario.append(dict(dia=str(dia), etapa=etapa(dia), bens=len(g), ativos=int((g.sit == "ATIVO").sum()),
                       lotes=len(rows), tempo_h=round(tempo, 2), bens_medidos=bens,
                       prod=round(bens / tempo, 1) if tempo > 0 else None))
D = pd.DataFrame(diario)
M["diario"] = diario
M["lotes"] = int(D.lotes.sum())
M["tempo_efetivo_h"] = round(D.tempo_h.sum(), 1)
M["bens_medidos"] = int(D.bens_medidos.sum())
M["produtividade"] = round(D.bens_medidos.sum() / D.tempo_h.sum(), 1)
M["seg_por_bem"] = round(D.tempo_h.sum() * 3600 / D.bens_medidos.sum(), 1)
E = D.groupby("etapa").agg(dias=("dia", "nunique"), bens=("bens", "sum"), ativos=("ativos", "sum"),
                            lotes=("lotes", "sum"), tempo_h=("tempo_h", "sum"), medidos=("bens_medidos", "sum"))
E["prod"] = (E.medidos / E.tempo_h).round(1)
E["seg_bem"] = (E.tempo_h * 3600 / E.medidos).round(1)
M["etapas"] = E.round(2).reset_index().to_dict("records")
todos_lotes = [r for _, g in chk.groupby("dia") for r in prod_rows(g) if r["medido"]]
sb = np.array([r["gap"] / r["n"] for r in todos_lotes])
M["seg_por_bem_mediana_lote"] = round(float(np.median(sb)), 1)
M["lote_mediano"] = float(np.median([r["n"] for r in todos_lotes]))
# projeção de conclusão: pendentes ao ritmo da etapa 3
p3 = E.loc[3, "prod"]
M["horas_para_concluir_ritmo_e3"] = round(M["pendentes"] / p3, 1)
M["horas_total_projetado"] = round(M["tempo_efetivo_h"] + M["pendentes"] / p3, 1)

# divergências
d = atc[~atc.ok].copy()
def setor(l): return re.sub(r"\s*-\s*(SALA|CORREDOR).*$", "", l)
def cat(r):
    s, i = r.loc_sis_n, r.loc_inv_n
    if setor(s) == setor(i): return "Granularidade cadastral (sala de reunião ou corredor do mesmo setor)"
    if any(k in i for k in ("DEPÓSITO", "ALMOXARIFADO", "ARQUIVO")): return "Recolhimento a depósito, almoxarifado ou arquivo"
    if r.pav == "Localizações lógicas": return "Origem em localização lógica (bens novos, termos, CFC)"
    if r.pav == r.pav_inv: return "Remanejamento entre setores do mesmo pavimento"
    return "Remanejamento entre pavimentos"
d["cat"] = d.apply(cat, axis=1)
M["div_categorias"] = d.cat.value_counts().to_dict()
pares = d.groupby(["loc_sis", "loc_inv"]).size().sort_values(ascending=False)
M["div_pares_n"] = len(pares)
M["div_top10"] = [(a, b, int(n)) for (a, b), n in pares.head(10).items()]
M["div_top10_sum"] = int(pares.head(10).sum())
M["div_pares_ate5"] = int((pares <= 5).sum()); M["div_bens_pares_ate5"] = int(pares[pares <= 5].sum())
tp = pd.DataFrame({"conf": atc.groupby("pav").size(), "div": d.groupby("pav").size()}).fillna(0)
tp["taxa"] = (tp["div"] / tp.conf * 100).round(1)
M["div_taxa_pavimento"] = tp.sort_values("taxa", ascending=False).taxa.to_dict()
tc = pd.DataFrame({"conf": atc.groupby("classe").size(), "div": d.groupby("classe").size()}).fillna(0)
tc["taxa"] = (tc["div"] / tc.conf * 100).round(1)
M["div_taxa_classe"] = tc.sort_values("conf", ascending=False).taxa.head(5).to_dict()

# cobertura
cp = pd.DataFrame({"ativos": at.groupby("pav").size(), "conf": atc.groupby("pav").size()}).fillna(0).astype(int)
cp["cob"] = (cp.conf / cp.ativos * 100).round(1)
M["cobertura_pavimento"] = cp.sort_values("cob", ascending=False).reset_index().to_dict("records")
cc = pd.DataFrame({"ativos": at.groupby("classe").size(), "conf": atc.groupby("classe").size()}).fillna(0).astype(int)
cc["cob"] = (cc.conf / cc.ativos * 100).round(1)
M["cobertura_classe"] = cc.sort_values("ativos", ascending=False).reset_index().to_dict("records")
cl = pd.DataFrame({"ativos": at[at.loc_sis_n != ""].groupby("loc_sis").size(), "conf": atc.groupby("loc_sis").size()}).fillna(0).astype(int)
cl["cob"] = (cl.conf / cl.ativos * 100).round(1)
cl = cl.sort_values("ativos", ascending=False)
M["loc_100"] = int((cl.cob == 100).sum()); M["loc_90"] = int((cl.cob >= 90).sum()); M["loc_n"] = len(cl)
M["cobertura_top20"] = cl.head(20).reset_index().to_dict("records")
M["demais_locs"] = dict(n=len(cl) - 20, ativos=int(cl.iloc[20:].ativos.sum()))
pend = at[at.dh.isna()]
M["pendentes_loc"] = pend.loc_sis.fillna("(sem localização)").value_counts().head(8).to_dict()
M["pendentes_classe"] = pend.classe.value_counts().to_dict()
M["pendentes_valor_k"] = round(pend.valor.sum() / 1e3, 1)
M["pendentes_logicos"] = int(pend.pav.eq("Localizações lógicas").sum())
M["pendentes_valor_softwares_k"] = round(pend[pend.classe.str.contains("SOFTWARE", na=False)].valor.sum() / 1e3, 1)
# ABC
m2 = mov.dropna(subset=["valor"]).sort_values("valor", ascending=False).copy()
m2["cum"] = m2.valor.cumsum() / m2.valor.sum()
m2["abc"] = np.where(m2.cum <= 0.8, "A", np.where(m2.cum <= 0.95, "B", "C"))
ab = pd.DataFrame({"n": m2.groupby("abc").size(), "conf": m2[m2.dh.notna()].groupby("abc").size()})
ab["cob"] = (ab.conf / ab.n * 100).round(1); ab["pct_n"] = (ab.n / len(m2) * 100).round(1)
M["abc"] = ab.reset_index().to_dict("records")
# baixados encontrados
na = chk[chk.sit != "ATIVO"]
M["nao_ativos_encontrados"] = dict(n=len(na), sit=na.sit.value_counts().to_dict(),
                                   em_deposito=int(na.loc_inv_n.str.contains("DEPÓSITO").sum()))

(ROOT / "11_metricas.json").write_text(json.dumps(M, ensure_ascii=False, indent=1, default=str), encoding="utf-8")

# ------------------------------------------------------------------ figuras estáticas (inalteradas)
for f in ["figura_1_fluxo.png", "fluxograma_etapas.png", "piloto_produtividade_lote.png",
          "piloto_tempo_bem.png", "piloto_montecarlo.png"]:
    shutil.copy(OLD / f, OUT / f)

# ------------------------------------------------------------------ composição por classe
cls = at.classe.value_counts()
top = ["MÓVEIS E UTENSÍLIOS DE ESCRITÓRIO", "EQUIPAMENTOS DE PROCESSAMENTO DE DADOS",
       "MÁQUINAS E EQUIPAMENTOS", "UTENSÍLIOS DE COPA E COZINHA", "MUSEU E OBRAS DE ARTE"]
names = ["Móveis e utensílios de escritório", "Equip. de processamento de dados",
         "Máquinas e equipamentos", "Utensílios de copa e cozinha", "Museu e obras de arte", "Demais classes"]
vals = [int(cls[t]) for t in top] + [int(cls.drop(top).sum())]
fig, ax = plt.subplots(figsize=(7.4, 3.6))
pos = range(len(names)); bars = ax.barh(pos, vals, color=GREY, height=0.62)
ax.set_yticks(pos, names); ax.invert_yaxis(); clean(ax)
ax.set_xlabel("Quantidade de bens ativos"); ax.set_xlim(0, max(vals) * 1.12)
for b, v in zip(bars, vals):
    ax.text(v + 15, b.get_y() + b.get_height() / 2, br(v), va="center", fontsize=10)
save(fig, "composicao_classes.png")

# ------------------------------------------------------------------ idade do acervo
bins = [(None, 1999, "até 1999"), (2000, 2004, "2000–2004"), (2005, 2009, "2005–2009"), (2010, 2014, "2010–2014"),
        (2015, 2019, "2015–2019"), (2020, 2024, "2020–2024"), (2025, 2026, "2025–2026")]
yr = mov.dt_ent.dt.year
labels, vals = [], []
for lo, hi, lb in bins:
    m = (yr <= hi) if lo is None else ((yr >= lo) & (yr <= hi))
    labels.append(lb); vals.append(int(m.sum()))
fig, ax = plt.subplots(figsize=(7.4, 3.3))
bars = ax.bar(labels, vals, color=GREY, width=0.6); clean(ax)
ax.set_ylabel("Bens móveis ativos"); ax.set_xlabel("Período de entrada do bem no acervo")
ax.set_ylim(0, max(vals) * 1.15)
for b, v in zip(bars, vals):
    ax.text(b.get_x() + b.get_width() / 2, v + 10, str(v), ha="center", va="bottom", fontsize=10)
save(fig, "idade_acervo.png")

# ------------------------------------------------------------------ curva ABC
x = np.arange(1, len(m2) + 1) / len(m2) * 100; y = m2.cum.values * 100
xa = (m2.abc == "A").sum() / len(m2) * 100; xb = (m2.abc != "C").sum() / len(m2) * 100
fig, ax = plt.subplots(figsize=(7.4, 3.6))
ax.plot(x, y, color=DARK, lw=2); clean(ax)
ax.axhline(80, ls=":", color="#888"); ax.axhline(95, ls=":", color="#888")
ax.axvline(xa, ls="--", color="#888"); ax.axvline(xb, ls="--", color="#888")
ax.set_xlim(0, 100); ax.set_ylim(0, 102)
ax.set_xlabel("Bens móveis ativos ordenados por valor (%)"); ax.set_ylabel("Valor acumulado (%)")
for xx, lb in ((xa / 2, "A"), ((xa + xb) / 2, "B"), ((xb + 100) / 2, "C")):
    ax.text(xx, 42, lb, ha="center", fontsize=13)
ax.annotate(f"{br(xa, 1)}% dos bens\n= 80% do valor", xy=(xa, 80), xytext=(xa + 12, 68), fontsize=10,
            arrowprops=dict(arrowstyle="->", color="black", lw=0.8))
save(fig, "curva_abc.png")

# ------------------------------------------------------------------ bens ativos conferidos por dia
lab = [pd.Timestamp(r["dia"]).strftime("%d/%m") for r in diario]
vals = [r["ativos"] for r in diario]; et = [r["etapa"] for r in diario]
shade = {1: DARK, 2: GREY, 3: LIGHT}
fig, ax = plt.subplots(figsize=(7.4, 3.4))
bars = ax.bar(lab, vals, color=[shade[e] for e in et], width=0.7); clean(ax)
ax.set_ylabel("Bens ativos conferidos"); ax.set_xlabel("Dia de campo da campanha censitária (2026)")
ax.set_ylim(0, max(vals) * 1.18); ax.tick_params(axis="x", labelsize=9, rotation=45)
for b, v in zip(bars, vals):
    ax.text(b.get_x() + b.get_width() / 2, v + 8, str(v), ha="center", va="bottom", fontsize=8.5)
from matplotlib.patches import Patch
ax.legend(handles=[Patch(color=DARK, label="Etapa 1 (18–21 ago.)"), Patch(color=GREY, label="Etapa 2 (24–28 ago.)"),
                   Patch(color=LIGHT, label="Etapa 3 (1–22 set.)")], frameon=False, fontsize=9, loc="upper right")
save(fig, "figura_2_conferencias_dia.png")

# ------------------------------------------------------------------ produtividade por dia
vals = [r["prod"] or 0 for r in diario]
fig, ax = plt.subplots(figsize=(7.4, 3.4))
bars = ax.bar(lab, vals, color=[shade[e] for e in et], width=0.7); clean(ax)
ax.axhline(M["produtividade"], ls="--", color="#8b1a1a", lw=1.3,
           label=f"média ponderada da campanha ≈ {br(M['produtividade'])} bens/h")
ax.axhline(94, ls=":", color="black", lw=1.2, label="inventário-piloto ≈ 94 bens/h")
ax.set_ylabel("Produtividade efetiva (bens/hora)"); ax.set_xlabel("Dia de campo da campanha censitária (2026)")
ax.set_ylim(0, max(vals) * 1.22); ax.tick_params(axis="x", labelsize=9, rotation=45)
for b, v in zip(bars, vals):
    ax.text(b.get_x() + b.get_width() / 2, v + 5, str(int(round(v))) if v else "n.d.", ha="center", va="bottom", fontsize=8.5)
ax.legend(frameon=False, fontsize=9, loc="lower center", bbox_to_anchor=(0.5, 1.0), ncol=2)
save(fig, "campanha_produtividade_dia.png")

# ------------------------------------------------------------------ cobertura acumulada
acum = atc.groupby("dia").size().sort_index().cumsum() / len(at) * 100
fig, ax = plt.subplots(figsize=(7.4, 3.4))
dias = [pd.Timestamp(d) for d in acum.index]
ax.plot(dias, acum.values, color=DARK, lw=2, marker="o", ms=4); clean(ax)
ax.set_ylim(0, 105); ax.set_ylabel("Cobertura acumulada do acervo ativo (%)"); ax.set_xlabel("Data (2026)")
ax.axhline(100, ls=":", color="#888")
for xd, yv, lb in ((dias[3], acum.iloc[3], "fim da etapa 1"), (dias[8], acum.iloc[8], "fim da etapa 2"), (dias[-1], acum.iloc[-1], "posição atual")):
    ax.annotate(f"{lb}\n{br(yv, 1)}%", xy=(xd, yv), xytext=(0, -38 if lb != "posição atual" else -40), textcoords="offset points",
                ha="center", fontsize=9, arrowprops=dict(arrowstyle="-", color="#888", lw=0.8))
ax.xaxis.set_major_formatter(matplotlib.dates.DateFormatter("%d/%m")); ax.tick_params(axis="x", labelsize=9)
save(fig, "cobertura_acumulada.png")

# ------------------------------------------------------------------ concordância
fig, ax = plt.subplots(figsize=(7.4, 2.3))
a, dv = M["concordancia_pct"], M["divergencia_pct"]
ax.barh([0], [a], color=DARK, height=0.42, label="Concordantes")
ax.barh([0], [dv], left=[a], color=LIGHT, height=0.42, label="Divergentes")
clean(ax, left=False); ax.set_xlim(0, 100); ax.set_yticks([])
ax.set_xlabel("Proporção dos bens ativos conferidos (%)")
ax.legend(frameon=False, loc="upper center", bbox_to_anchor=(0.5, 1.3), ncol=2, fontsize=10)
ax.text(a / 2, 0, f"{br(a, 1)}% ({br(M['concordantes'])} bens)", ha="center", va="center", color="white", fontsize=10)
ax.text(a + dv / 2, 0, f"{br(dv, 1)}%\n({br(M['divergentes'])})", ha="center", va="center", color="black", fontsize=9)
save(fig, "figura_3_concordancia.png")

# ------------------------------------------------------------------ cobertura por classe
names = ["Móveis e utensílios", "Processamento de dados", "Máquinas e equipamentos", "Copa e cozinha", "Museu e obras de arte"]
vals = [float(cc.loc[t, "cob"]) for t in top]
fig, ax = plt.subplots(figsize=(7.4, 3.2))
pos = range(len(names)); bars = ax.barh(pos, vals, color=GREY, height=0.58)
ax.set_yticks(pos, names); ax.invert_yaxis(); ax.set_xlim(0, 110); ax.set_xlabel("Cobertura da classe (%)"); clean(ax)
for b, v in zip(bars, vals):
    ax.text(v + 1.2, b.get_y() + b.get_height() / 2, br(v, 1), va="center", fontsize=10)
save(fig, "figura_4_cobertura_classes.png")

# ------------------------------------------------------------------ cobertura por pavimento
cps = cp.sort_values(["cob", "ativos"], ascending=[False, False])
fig, ax = plt.subplots(figsize=(7.4, 5.2))
pos = range(len(cps)); bars = ax.barh(pos, cps.cob.values, color=GREY, height=0.7)
ax.set_yticks(pos, cps.index); ax.invert_yaxis(); ax.set_xlim(0, 112); ax.set_xlabel("Cobertura do inventário (%)"); clean(ax)
for b, v, n in zip(bars, cps.cob.values, cps.ativos.values):
    ax.text(v + 1, b.get_y() + b.get_height() / 2, f"{br(v, 1)}  (n = {n})", va="center", fontsize=9)
save(fig, "cobertura_pavimento.png")

# ------------------------------------------------------------------ comparação AS-IS
fig, ax = plt.subplots(figsize=(7.4, 3.4))
cats = ["Automatizado\nmedido na campanha\n(97,4% do acervo)", "Automatizado\nprojetado (Monte Carlo)\n(100% do acervo)", "Manual AS-IS\nestimado pela literatura"]
vals = [M["tempo_efetivo_h"], 37.6, 156.5]
bars = ax.bar(cats, vals, color=[DARK, GREY, LIGHT], width=0.55); clean(ax)
ax.errorbar([1], [37.6], yerr=[[37.6 - 30.6], [46.6 - 37.6]], fmt="none", ecolor="black", capsize=5, lw=1.2)
ax.errorbar([2], [156.5], yerr=[[156.5 - 125], [188 - 156.5]], fmt="none", ecolor="black", capsize=5, lw=1.2)
ax.set_ylabel("Horas efetivas de conferência"); ax.set_ylim(0, 205); ax.tick_params(axis="x", labelsize=9)
for b, v, t in zip(bars, vals, [f"{br(M['tempo_efetivo_h'], 1)} h", "37,6 h\n(IC 95%: 30,6–46,6)", "125 a 188 h"]):
    ax.text(b.get_x() + b.get_width() / 2, v + (12 if v < 100 else 36), t, ha="center", va="bottom", fontsize=9.5)
save(fig, "piloto_comparacao_asis.png")

print(json.dumps({k: M[k] for k in ["registros", "ativos", "conferidos", "ativos_conferidos", "cobertura", "pendentes",
                                    "lotes", "tempo_efetivo_h", "produtividade", "seg_por_bem", "concordancia_pct",
                                    "divergentes", "horas_total_projetado", "dias_campo"]}, ensure_ascii=False, indent=1))
print("etapas:"); print(E)
