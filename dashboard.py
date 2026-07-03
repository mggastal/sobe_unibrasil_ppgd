#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gerador Dashboard UniBrasil (interativo)
========================================
Mesmo padrão dos demais dashboards do estúdio: dados DIÁRIOS embutidos como
JSON e TODA a filtragem (datas, períodos, modalidade) acontece no navegador.

Camada específica UniBrasil: filtro de MODALIDADE no topo (Presencial / EAD / Semi),
onde EAD consolida EAD+SEMI mas o Semi é isolável. Ranking de CURSOS e RANKING
DE CRIATIVOS respeitam data + modalidade.

Fonte: .xlsx (padrão) ou Google Sheets (mesmas abas).
Desenvolvido por Sobé Estratégias.
"""

import pandas as pd
import json
from datetime import date
from pathlib import Path

# ───────────────────────── CONFIG ─────────────────────────
CLIENTE        = "UniBrasil · Mestrado & Doutorado"
TEMPLATE_FILE  = "template.html"
OUTPUT_FILE    = "index.html"
DATA_FILE      = "data.json"

XLSX_FILE   = "_UNIBRASIL_V2__PLANILHA_BASE_PARA_DASHBOARD.xlsx"
USAR_SHEETS = True
SHEET_ID    = "1SeR6m6p_KQ9NHPgOKfUbQc3zmxqYTFVZe4EJg03xgI4"

# ── Filtro de campanhas por nome (aplicado a TODAS as abas) ──
# Dois dashboards saem da MESMA planilha só trocando estas duas linhas:
#   • UniBrasil geral: EXCLUIR = ["MESTRADO", "DOUTORADO"]  /  APENAS = []
#   • Mestrado/Doutorado: EXCLUIR = []  /  APENAS = ["MESTRADO", "DOUTORADO"]
# A regra é por substring, sem acento/caixa. Se uma campanha casar em APENAS (quando
# a lista não é vazia) ela entra; se casar em EXCLUIR ela sai. APENAS tem precedência.
CAMP_EXCLUIR = []
CAMP_APENAS  = ["MESTRADO", "DOUTORADO"]

ABAS = {
    "meta": "meta-ads",
    "g_pesquisa": "google-ads-pesquisa",
    "g_outros": "google-ads-outros",
    "g_keywords": "google-ads",
    "bk_genage": "breakdown-gender-age",
    "bk_platform": "breakdown-platform",
    "g_bd_gender": "google-breakdown-gender",
    "g_bd_age": "google-breakdown-age",
    # blocos anuais fixos
    "meta_2024": "meta-2024",
    "meta_2025": "meta-2025",
    "g_2024_search": "google-2024-search",
    "g_2024_outros": "google-2024-outros",
    "g_2025_search": "google-2025-search",
    "g_2025_outros": "google-2025-outros",
}
MOEDA = "R$"

CURSO_MAP = {
    "ADM": "Administração", "ANALISE_DESENV_SOFTWARE": "Análise e Desenv. de Sistemas",
    "ARQ": "Arquitetura e Urbanismo", "BIOMED": "Biomedicina",
    "CIENCIAS_CONTABEIS": "Ciências Contábeis", "COMERCIO_EXT": "Comércio Exterior",
    "DESIGN_INTERIORES": "Design de Interiores", "DIRE": "Direito",
    "EDFIS": "Educação Física", "ENFERMAGEM": "Enfermagem", "ENG_CIVIL": "Eng. Civil",
    "ENG_MECANICA": "Eng. Mecânica", "ENG_PROD": "Eng. de Produção",
    "ENG_SOFTWARE": "Eng. de Software", "ENG": "Engenharias",
    "ESTETICA": "Estética e Cosmética", "FONO": "Fonoaudiologia",
    "GESTAO_E_PROCESSOS": "Gestão e Processos", "GESTAO_RH": "Gestão de RH",
    "MARKETING": "Marketing", "MEDVET": "Medicina Veterinária", "NUTRI": "Nutrição",
    "ODON": "Odontologia", "PEDAGOGIA": "Pedagogia", "PSICOLOGIA": "Psicologia",
    "PUBLI": "Publicidade e Propaganda",
}

# ───────────────────────── HELPERS ─────────────────────────
def _passa_campanha(nome):
    u = str(nome).upper()
    if CAMP_APENAS:
        return any(t.upper() in u for t in CAMP_APENAS)
    if CAMP_EXCLUIR:
        return not any(t.upper() in u for t in CAMP_EXCLUIR)
    return True

def _filtra_campanhas(df):
    """Remove linhas de campanhas fora do escopo (todas as abas têm 'Campaign Name')."""
    if "Campaign Name" in df.columns and (CAMP_APENAS or CAMP_EXCLUIR):
        return df[df["Campaign Name"].apply(_passa_campanha)].copy()
    return df

def carregar(key):
    aba = ABAS[key]
    if USAR_SHEETS:
        from urllib.parse import quote
        # gviz/tq lê a aba pelo nome; encodamos o nome para suportar acentos/espaços
        url = (f"https://docs.google.com/spreadsheets/d/{SHEET_ID}"
               f"/gviz/tq?tqx=out:csv&headers=1&sheet={quote(aba)}")
        df = pd.read_csv(url, dtype=str)  # tudo como texto; num() converte depois
    else:
        df = pd.read_excel(XLSX_FILE, sheet_name=aba)
    return _filtra_campanhas(df)

def num(s):
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce").fillna(0.0)
    c = s.astype(str).str.strip().str.replace("R$", "", regex=False).str.replace(" ", "", regex=False)
    tem_virgula = c.str.contains(",", regex=False)
    # BR (1.234,56): ponto = milhar, vírgula = decimal  → remove pontos, vírgula vira ponto
    br = c.str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    # US (1234.56 ou 8.000324): ponto já é o decimal  → mantém como está
    us = c
    out = us.where(~tem_virgula, br)
    return pd.to_numeric(out, errors="coerce").fillna(0.0)

def is_branding(c):
    c = str(c).upper()
    return c.startswith("BRAND") or c.startswith("GENERICA") or c.startswith("MATRICULA")

def modalidade(c):
    c = str(c).upper()
    if "PRESENCIAL" in c: return "PRESENCIAL"
    if "SEMI" in c: return "SEMI"
    if "EAD" in c: return "EAD"
    return "OUTROS"

def curso(c):
    if is_branding(c): return ""
    first = str(c).upper().split("-")[0]
    return CURSO_MAP.get(first, first.title())

def r2(x):
    try: return round(float(x), 2)
    except Exception: return 0.0

def numcol(df, name):
    """num() de uma coluna que pode não existir na aba (retorna zeros se ausente)."""
    return num(df[name]) if name in df.columns else pd.Series(0.0, index=df.index)

def thumb(row):
    for col in ("Thumbnail URL", "Image URL"):
        v = row.get(col)
        if isinstance(v, str) and v.startswith("http"):
            return v
    return ""

def parse_data(serie):
    """Datas podem vir como 2026-06-23 (xlsx/ISO) ou 23/06/2026 (Sheets BR).
    Tenta ISO primeiro; nas que falharem, tenta dia/mês/ano."""
    d = pd.to_datetime(serie, errors="coerce")
    faltam = d.isna()
    if faltam.any():
        d2 = pd.to_datetime(serie[faltam], errors="coerce", dayfirst=True)
        d.loc[faltam] = d2
    return d

def col_lead_meta(df):
    """(legado) coluna única, mantido por compatibilidade."""
    return "Conversões" if "Conversões" in df.columns else "Action Leads"

# Colunas que compõem o "Lead" do Meta — igual ao Looker:
# formulário (Action Leads) + conversas no Messenger + pixel da LP.
# OBS: 'Conversion Contact Total' é idêntica a 'Action FB Pixel Custom (Offsite Conversion)'
# (mesma conversão em duas colunas), então usamos apenas UMA para não contar em dobro.
META_LEAD_COLS = [
    "Action Leads",
    "Action Messaging Conversations Started (Onsite Conversion)",
    "Action FB Pixel Custom (Offsite Conversion)",
]
def leads_meta(df):
    """Soma as colunas de conversão do Meta que existirem na aba."""
    total = pd.Series(0.0, index=df.index)
    achou = False
    for c in META_LEAD_COLS:
        if c in df.columns:
            total = total + num(df[c]); achou = True
    if not achou:  # fallback de segurança
        total = num(df["Action Leads"])
    return total

def col_conv_google(df):
    """Coluna de conversão do Google: usa 'Conversions' (primárias, igual ao Looker) se existir;
    senão cai para 'All Conversions' (comportamento anterior)."""
    return "Conversions" if "Conversions" in df.columns else "All Conversions"

# ───────────────────────── META (função reutilizável) ─────────────────────────
def build_meta(df):
    df = df.copy()
    df["Date"] = parse_data(df["Date"])
    df = df[df["Date"].notna()].copy()
    df["sp"] = num(df["Spend (Cost, Amount Spent)"])
    df["ld"] = leads_meta(df)
    df["im"] = num(df["Impressions"])
    df["ck"] = num(df["Clicks"])
    df["tp"]  = numcol(df, "Video Thruplay Watched Actions")
    df["modal"] = df["Campaign Name"].apply(modalidade)
    df["curso"] = df["Campaign Name"].apply(curso)
    df = df[df["modal"].isin(["PRESENCIAL", "EAD", "SEMI"])].copy()
    rows, thumbs, status = [], {}, {}
    for _, r in df.iterrows():
        ad = str(r.get("Ad Name", ""))
        key = r["Campaign Name"] + "||" + ad
        t = thumb(r)
        if t:
            thumbs[key] = t
        st = str(r.get("Status", "")).upper()
        if st == "ACTIVE" or key not in status:
            status[key] = st
        rows.append({
            "d": r["Date"].strftime("%Y-%m-%d"),
            "c": r["Campaign Name"], "ad": ad,
            "m": r["modal"], "cu": r["curso"],
            "br": bool(is_branding(r["Campaign Name"])),
            "sp": r2(r["sp"]), "ld": int(r["ld"]),
            "im": int(r["im"]), "ck": int(r["ck"]), "tp": int(r["tp"]),
        })
    return rows, thumbs, status

print("→ Lendo dados…")
meta_rows, thumbs, status = build_meta(carregar("meta"))

# ───────────────────────── GOOGLE (funções reutilizáveis) ─────────────────────────
def gprep(df, canal_col=None, canal_fixo=None):
    df = df.copy()
    df["Date"] = parse_data(df["Date (Segment)"])
    df = df[df["Date"].notna()]
    cvcol = col_conv_google(df)
    df["sp"] = num(df["Cost (Spend, Amount Spent)"])
    df["cv"] = num(df[cvcol])
    df["im"] = num(df["Impressions"])
    df["ck"] = num(df["Clicks"])
    out = []
    for _, r in df.iterrows():
        out.append({
            "d": r["Date"].strftime("%Y-%m-%d"),
            "c": str(r.get("Campaign Name", "")),
            "ch": canal_fixo or str(r.get(canal_col, "")).title(),
            "sp": r2(r["sp"]), "cv": r2(r["cv"]),
            "im": int(r["im"]), "ck": int(r["ck"]),
        })
    return out

def build_keywords(df):
    """Keywords diárias a partir de uma aba search (nível palavra-chave)."""
    df = df.copy()
    df["Date"] = parse_data(df["Date (Segment)"])
    df = df[df["Date"].notna()].copy()
    if "Keyword (Ad Group Criterion)" not in df.columns:
        return []
    cvcol = col_conv_google(df)
    df["sp"] = num(df["Cost (Spend, Amount Spent)"])
    df["cv"] = num(df[cvcol])
    df["ck"] = num(df["Clicks"])
    df["im"] = num(df["Impressions"])
    df["kw"] = df["Keyword (Ad Group Criterion)"].astype(str)
    df["dstr"] = df["Date"].dt.strftime("%Y-%m-%d")
    agg = df.groupby(["dstr", "kw"], as_index=False).agg(
        sp=("sp", "sum"), cv=("cv", "sum"), ck=("ck", "sum"), im=("im", "sum"))
    return [{"d": r["dstr"], "kw": r["kw"], "sp": r2(r["sp"]),
             "cv": r2(r["cv"]), "ck": int(r["ck"]), "im": int(r["im"])}
            for _, r in agg.iterrows()]

def conv_por_dia_campanha(df_kw):
    """Mapa (dia, campanha) -> soma de 'Conversions' a partir da aba keyword.
    Usado para corrigir a aba de pesquisa quando ela não tem a coluna 'Conversions'."""
    if "Conversions" not in df_kw.columns:
        return None
    df = df_kw.copy()
    df["Date"] = parse_data(df["Date (Segment)"])
    df = df[df["Date"].notna()].copy()
    df["dstr"] = df["Date"].dt.strftime("%Y-%m-%d")
    df["c"] = df["Campaign Name"].astype(str)
    df["cv"] = num(df["Conversions"])
    agg = df.groupby(["dstr", "c"], as_index=False).agg(cv=("cv", "sum"))
    return {(r["dstr"], r["c"]): r["cv"] for _, r in agg.iterrows()}

def gprep_search(df_search, df_kw, canal_fixo="Pesquisa"):
    """Prepara a aba de pesquisa AGREGANDO por dia+campanha (a aba tem várias linhas
    por dia/campanha, ex.: por ad group). Se não houver 'Conversions' na pesquisa,
    usa o total de 'Conversions' da aba keyword (mesmas campanhas/dias)."""
    df = df_search.copy()
    df["Date"] = parse_data(df["Date (Segment)"])
    df = df[df["Date"].notna()].copy()
    df["dstr"] = df["Date"].dt.strftime("%Y-%m-%d")
    df["c"] = df["Campaign Name"].astype(str)
    df["sp"] = num(df["Cost (Spend, Amount Spent)"])
    df["im"] = num(df["Impressions"])
    df["ck"] = num(df["Clicks"])
    tem_conv = "Conversions" in df.columns
    if tem_conv:
        df["cv"] = num(df["Conversions"])
    agg = df.groupby(["dstr", "c"], as_index=False).agg(
        sp=("sp", "sum"), im=("im", "sum"), ck=("ck", "sum"),
        **({"cv": ("cv", "sum")} if tem_conv else {}))
    mapa = None if tem_conv else conv_por_dia_campanha(df_kw)
    out = []
    for _, r in agg.iterrows():
        cv = r["cv"] if tem_conv else (mapa.get((r["dstr"], r["c"]), 0.0) if mapa else 0.0)
        out.append({"d": r["dstr"], "c": r["c"], "ch": canal_fixo,
                    "sp": r2(r["sp"]), "cv": r2(cv), "im": int(r["im"]), "ck": int(r["ck"])})
    return out

# Google "corrente" (dashboard principal): pesquisa (conversões corrigidas pela keyword) + outros
gkw_df = carregar("g_keywords")
google_rows = gprep_search(carregar("g_pesquisa"), gkw_df, canal_fixo="Pesquisa") + \
              gprep(carregar("g_outros"), canal_col="Advertising Channel Type")
keyword_rows = build_keywords(gkw_df)

# ───────────────────────── BLOCOS ANUAIS FIXOS ─────────────────────────
def gprep_agg(df, canal_col=None, canal_fixo=None):
    """Igual ao gprep, mas agrega por dia+campanha+canal (reduz volume das abas anuais)."""
    df = df.copy()
    df["Date"] = parse_data(df["Date (Segment)"])
    df = df[df["Date"].notna()].copy()
    cvcol = col_conv_google(df)
    df["sp"] = num(df["Cost (Spend, Amount Spent)"])
    df["cv"] = num(df[cvcol])
    df["im"] = num(df["Impressions"])
    df["ck"] = num(df["Clicks"])
    df["dstr"] = df["Date"].dt.strftime("%Y-%m-%d")
    df["c"] = df["Campaign Name"].astype(str)
    df["ch"] = canal_fixo if canal_fixo else df[canal_col].astype(str).str.title()
    agg = df.groupby(["dstr", "c", "ch"], as_index=False).agg(
        sp=("sp", "sum"), cv=("cv", "sum"), im=("im", "sum"), ck=("ck", "sum"))
    return [{"d": r["dstr"], "c": r["c"], "ch": r["ch"], "sp": r2(r["sp"]),
             "cv": r2(r["cv"]), "im": int(r["im"]), "ck": int(r["ck"])}
            for _, r in agg.iterrows()]

def build_meta_ano(df):
    """Meta anual agregado por dia+campanha (sem nível de criativo) — leve para páginas anuais."""
    df = df.copy()
    df["Date"] = parse_data(df["Date"])
    df = df[df["Date"].notna()].copy()
    df["sp"] = num(df["Spend (Cost, Amount Spent)"])
    df["ld"] = leads_meta(df)
    df["im"] = num(df["Impressions"])
    df["ck"] = num(df["Clicks"])
    df["m"] = df["Campaign Name"].apply(modalidade)
    df["cu"] = df["Campaign Name"].apply(curso)
    df["br"] = df["Campaign Name"].apply(is_branding)
    df = df[df["m"].isin(["PRESENCIAL", "EAD", "SEMI"])].copy()
    df["dstr"] = df["Date"].dt.strftime("%Y-%m-%d")
    df["c"] = df["Campaign Name"].astype(str)
    agg = df.groupby(["dstr", "c", "m", "cu", "br"], as_index=False).agg(
        sp=("sp", "sum"), ld=("ld", "sum"), im=("im", "sum"), ck=("ck", "sum"))
    return [{"d": r["dstr"], "c": r["c"], "m": r["m"], "cu": r["cu"], "br": bool(r["br"]),
             "sp": r2(r["sp"]), "ld": int(r["ld"]), "im": int(r["im"]), "ck": int(r["ck"])}
            for _, r in agg.iterrows()]

def build_keywords_ano(df):
    """Top keywords do ano agregadas (sem nível diário) — leve."""
    df = df.copy()
    df["Date"] = parse_data(df["Date (Segment)"])
    df = df[df["Date"].notna()].copy()
    if "Keyword (Ad Group Criterion)" not in df.columns:
        return []
    cvcol = col_conv_google(df)
    df["sp"] = num(df["Cost (Spend, Amount Spent)"])
    df["cv"] = num(df[cvcol])
    df["ck"] = num(df["Clicks"])
    df["im"] = num(df["Impressions"])
    df["kw"] = df["Keyword (Ad Group Criterion)"].astype(str)
    agg = df.groupby(["kw"], as_index=False).agg(
        sp=("sp", "sum"), cv=("cv", "sum"), ck=("ck", "sum"), im=("im", "sum"))
    agg = agg.sort_values("sp", ascending=False).head(60)
    return [{"kw": r["kw"], "sp": r2(r["sp"]), "cv": r2(r["cv"]),
             "ck": int(r["ck"]), "im": int(r["im"])} for _, r in agg.iterrows()]

def build_google_ano(search_key, outros_key):
    search_df = carregar(search_key)
    rows = gprep_agg(search_df, canal_fixo="Pesquisa") + \
           gprep_agg(carregar(outros_key), canal_col="Advertising Channel Type")
    kws = build_keywords_ano(search_df)
    return rows, kws

meta_2024 = build_meta_ano(carregar("meta_2024"))
meta_2025 = build_meta_ano(carregar("meta_2025"))
g2024_rows, g2024_kw = build_google_ano("g_2024_search", "g_2024_outros")
g2025_rows, g2025_kw = build_google_ano("g_2025_search", "g_2025_outros")

# ───────────────────────── ÍNDICE DE CAMPANHAS ─────────────────────────
# Breakdowns (demo/posicionamento) referenciam campanhas por índice (ci) para não
# repetir o nome em dezenas de milhares de linhas — assim respondem ao filtro de
# campanhas sem inflar o arquivo.
_camps_meta = sorted({r["c"] for r in meta_rows})

# ───────────────────────── DEMOGRAFIA (diária, por campanha) ─────────────────────────
ga = carregar("bk_genage")
ga["Date"] = parse_data(ga["Date"])
ga = ga[ga["Date"].notna()].copy()
ga["sp"] = num(ga["Spend (Cost, Amount Spent)"])
ga["ld"] = leads_meta(ga)
ga["modal"] = ga["Campaign Name"].apply(modalidade)
ga["ag"] = ga["Age (Breakdown)"].astype(str)
ga["g"] = ga["Gender (Breakdown)"].astype(str).map({"female": "f", "male": "m"}).fillna("u")
ga = ga[ga["modal"].isin(["PRESENCIAL", "EAD", "SEMI"])].copy()
ga["dstr"] = ga["Date"].dt.strftime("%Y-%m-%d")
ga["c"] = ga["Campaign Name"].astype(str)
demo_agg = ga.groupby(["dstr", "c", "ag", "g"], as_index=False).agg(sp=("sp", "sum"), ld=("ld", "sum"))

# ───────────────────────── POSICIONAMENTO (diário, por campanha) ─────────────────────────
pt = carregar("bk_platform")
pt["Date"] = parse_data(pt["Date"])
pt = pt[pt["Date"].notna()].copy()
pt["sp"] = num(pt["Spend (Cost, Amount Spent)"])
pt["ld"] = leads_meta(pt)
pt["modal"] = pt["Campaign Name"].apply(modalidade)
pt = pt[pt["modal"].isin(["PRESENCIAL", "EAD", "SEMI"])].copy()
pt["dstr"] = pt["Date"].dt.strftime("%Y-%m-%d")
pt["c"] = pt["Campaign Name"].astype(str)
pt["pp"] = pt["Platform Position (Breakdown)"].astype(str)
pos_agg = pt.groupby(["dstr", "c", "pp"], as_index=False).agg(sp=("sp", "sum"), ld=("ld", "sum"))

# índice final = campanhas do meta-ads + qualquer extra que só apareça nos breakdowns
CAMPS_LIST = sorted(set(_camps_meta) | set(demo_agg["c"]) | set(pos_agg["c"]))
CAMPS_IDX = {c: i for i, c in enumerate(CAMPS_LIST)}
CAMPS_MODAL = [modalidade(c) for c in CAMPS_LIST]

# ── formato compacto: linhas viram arrays [dia, campanha, dim..., sp, ld] com índices ──
DAYS_LIST = sorted(set(demo_agg["dstr"]) | set(pos_agg["dstr"]))
DAYS_IDX = {d: i for i, d in enumerate(DAYS_LIST)}
AG_LIST = sorted(demo_agg["ag"].unique())
AG_IDX = {a: i for i, a in enumerate(AG_LIST)}
G_IDX = {"f": 0, "m": 1, "u": 2}
POS_LIST = sorted(pos_agg["pp"].unique())
POS_IDX = {p: i for i, p in enumerate(POS_LIST)}

demo_rows = [[DAYS_IDX[r["dstr"]], CAMPS_IDX[r["c"]], AG_IDX[r["ag"]], G_IDX[r["g"]],
              r2(r["sp"]), int(r["ld"])] for _, r in demo_agg.iterrows()]
pos_rows = [[DAYS_IDX[r["dstr"]], CAMPS_IDX[r["c"]], POS_IDX[r["pp"]],
             r2(r["sp"]), int(r["ld"])] for _, r in pos_agg.iterrows()]

# ───────────────────────── DEMOGRAFIA GOOGLE (diária) ─────────────────────────
def build_gdemo(key, col, out_key):
    df = carregar(key)
    df["Date"] = parse_data(df["Date (Segment)"])
    df = df[df["Date"].notna()].copy()
    df["sp"] = num(df["Cost (Spend, Amount Spent)"])
    df["cv"] = num(df[col_conv_google(df)])
    df["dstr"] = df["Date"].dt.strftime("%Y-%m-%d")
    df["k"] = df[col].astype(str)
    agg = df.groupby(["dstr", "k"], as_index=False).agg(sp=("sp", "sum"), cv=("cv", "sum"))
    return [{"d": r["dstr"], out_key: r["k"], "sp": r2(r["sp"]), "cv": r2(r["cv"])} for _, r in agg.iterrows()]

gdemo_gen = build_gdemo("g_bd_gender", "Gender (Ad Group Criterion)", "g")
gdemo_age = build_gdemo("g_bd_age", "Age (Ad Group Criterion)", "ag")

# As abas de demografia do Google não têm coluna de campanha — não dá para filtrá-las
# por nome. Se este dashboard usa um recorte de campanhas (APENAS/EXCLUIR ativo com
# poucas campanhas), zeramos para não misturar dados de fora do escopo.
if CAMP_APENAS:
    gdemo_gen, gdemo_age = [], []

# ───────────────────────── PAYLOAD ─────────────────────────
all_dates = [x["d"] for x in meta_rows] + [x["d"] for x in google_rows]
D = {
    "cliente": CLIENTE, "moeda": MOEDA, "gerado_em": date.today().strftime("%d/%m/%Y"),
    "dmin": min(all_dates), "dmax": max(all_dates),
    "meta": meta_rows, "google": google_rows, "demo": demo_rows,
    "pos": pos_rows, "gdemo": {"gen": gdemo_gen, "age": gdemo_age},
    "campsList": CAMPS_LIST, "campsModal": CAMPS_MODAL,
    "days": DAYS_LIST, "agList": AG_LIST, "posList": POS_LIST,
    "keywords": keyword_rows, "thumbs": thumbs, "status": status,
    "anos": {
        "2024": {"meta": meta_2024, "google": g2024_rows, "keywords": g2024_kw},
        "2025": {"meta": meta_2025, "google": g2025_rows, "keywords": g2025_kw},
    },
}
Path(DATA_FILE).write_text(json.dumps(D, ensure_ascii=False), encoding="utf-8")
print(f"✓ {DATA_FILE}: {len(meta_rows)} linhas Meta, {len(google_rows)} Google, {len(demo_rows)} demo, "
      f"{len(pos_rows)} posicionamento, {len(gdemo_gen)+len(gdemo_age)} demo Google")
print(f"  anuais → 2024: meta {len(meta_2024)}, google {len(g2024_rows)}, kw {len(g2024_kw)} | "
      f"2025: meta {len(meta_2025)}, google {len(g2025_rows)}, kw {len(g2025_kw)}")

tpl = Path(TEMPLATE_FILE).read_text(encoding="utf-8")
Path(OUTPUT_FILE).write_text(tpl.replace("/*__DATA__*/", json.dumps(D, ensure_ascii=False)), encoding="utf-8")
print(f"✓ {OUTPUT_FILE} gerado")
