# =========================================================
# app.py — Dashboard Executivo (SENAI) — CONSOLIDADO (FINAL)
# AJUSTE PRINCIPAL:
# - KPIs + Termômetro + Diagnóstico: SEM filtros (base institucional) -> bate com o diagnóstico
# - Filtros laterais aplicados SOMENTE em:
#   (1) Tabela NR (VIGENTE x VENCIDO por Treinamento/NR)
#   (2) Gráfico por Ano
#   (3) Base Detalhada + Exportação
# - Unidade Ouro + Ranking Top 5: SEM filtros (travados)
# =========================================================

import os
import io
import re
import time
import shutil
import unicodedata
from datetime import datetime

import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt


# =========================================================
# CONFIG
# =========================================================
st.set_page_config(page_title="Dashboard Executivo - Treinamentos SENAI", layout="wide")

ARQUIVO = "CONTROLE DE TREINAMENTOS SENAI - POWER BI.xlsx"
ABA = 0

# Cabeçalho "LINHA 4" no Excel -> header=3
HEADER_LINHA = 3

LOGO_PATH = "logo_senai.png"

# Paleta SENAI
SENAI_LARANJA = "#F37021"
SENAI_AZUL = "#003DA5"
CINZA_FUNDO = "#F5F6F8"
CINZA_BORDA = "#E5E7EB"
TEXTO_CINZA = "#6B7280"

# Status (conforme você pediu)
COR_VIGENTE = SENAI_AZUL
COR_VENCIDO = SENAI_LARANJA


# =========================================================
# CSS (EXECUTIVO + SIDEBAR BRANCA)
# =========================================================
st.markdown(
    f"""
<style>
[data-testid="stAppViewContainer"] {{
    background: {CINZA_FUNDO};
}}
.block-container {{
    padding-top: 1.0rem;
    padding-bottom: 2.0rem;
}}

[data-testid="stSidebar"] {{
    background: #ffffff !important;
    border-right: 1px solid {CINZA_BORDA};
}}
[data-testid="stSidebar"] * {{
    color: {SENAI_AZUL} !important;
}}

[data-testid="stSidebar"] div[data-baseweb="select"] > div {{
    background: #ffffff !important;
    border-radius: 12px !important;
    border: 1px solid {CINZA_BORDA} !important;
}}
[data-testid="stSidebar"] div[data-baseweb="select"] span {{
    color: {SENAI_AZUL} !important;
    font-weight: 900 !important;
}}
[data-testid="stSidebar"] div[data-baseweb="select"] input {{
    color: {SENAI_AZUL} !important;
}}
ul[role="listbox"] {{
    background: #ffffff !important;
}}
ul[role="listbox"] * {{
    color: {SENAI_AZUL} !important;
    font-weight: 900 !important;
}}

[data-testid="stSidebar"] details {{
    border-radius: 14px;
    background: #ffffff;
    border: 1px solid {CINZA_BORDA};
    padding: 6px 10px;
}}
[data-testid="stSidebar"] summary {{
    font-weight: 900 !important;
}}

.stButton>button {{
    background: {SENAI_LARANJA};
    color: #fff !important;
    border: none;
    border-radius: 12px;
    font-weight: 900;
    padding: 0.55rem 0.85rem;
}}
.stButton>button:hover {{
    background: #d85f18;
}}
.stDownloadButton>button {{
    border-radius: 12px;
    font-weight: 900;
}}

.header-wrap {{
    background: #ffffff;
    border: 1px solid {CINZA_BORDA};
    border-radius: 18px;
    padding: 14px 18px;
    box-shadow: 0 6px 18px rgba(0,0,0,.05);
}}
.h-title {{
    font-size: 26px;
    font-weight: 900;
    margin: 0;
    color: {SENAI_AZUL};
}}
.h-sub {{
    margin: 6px 0 0 0;
    color: {TEXTO_CINZA};
    font-weight: 800;
}}
.badge {{
    display: inline-block;
    padding: 2px 10px;
    border-radius: 999px;
    background: rgba(243,112,33,.12);
    border: 1px solid rgba(243,112,33,.35);
    color: {SENAI_LARANJA};
    font-weight: 900;
    font-size: 12px;
}}

.kpi-grid {{
    display: grid;
    grid-template-columns: repeat(4, minmax(0, 1fr));
    gap: 12px;
}}
.kpi {{
    background: #ffffff;
    border: 1px solid {CINZA_BORDA};
    border-radius: 18px;
    padding: 14px 16px;
    box-shadow: 0 6px 18px rgba(0,0,0,.05);
    position: relative;
    overflow: hidden;
    min-height: 140px;
}}
.kpi::before {{
    content: "";
    position: absolute;
    left: 0; top: 0;
    width: 8px; height: 100%;
    background: {SENAI_LARANJA};
}}
.kpi h4 {{
    margin: 0;
    font-size: 12px;
    text-transform: uppercase;
    color: {TEXTO_CINZA};
    font-weight: 900;
}}
.kpi .v {{
    margin-top: 6px;
    font-size: 30px;
    font-weight: 900;
    color: #111827;
}}
.kpi .hint {{
    margin-top: 6px;
    font-size: 12px;
    color: {TEXTO_CINZA};
    font-weight: 800;
}}

.section {{
    background: #ffffff;
    border: 1px solid {CINZA_BORDA};
    border-radius: 18px;
    padding: 14px 16px;
    box-shadow: 0 6px 18px rgba(0,0,0,.05);
}}
.section-title {{
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 10px;
    margin: 0 0 10px 0;
}}
.section-title h3 {{
    margin: 0;
    font-size: 18px;
    font-weight: 900;
    color: {SENAI_AZUL};
}}
.pill {{
    display: inline-flex;
    align-items: center;
    gap: 8px;
    padding: 6px 10px;
    border-radius: 999px;
    border: 1px solid {CINZA_BORDA};
    background: #f9fafb;
    font-weight: 900;
    font-size: 12px;
    color: #111827;
}}

.thermo-wrap {{
    background: #ffffff;
    border: 1px solid {CINZA_BORDA};
    border-radius: 18px;
    padding: 14px 16px;
    box-shadow: 0 6px 18px rgba(0,0,0,.05);
}}
.thermo-bar {{
    height: 18px;
    width: 100%;
    background: #eef2ff;
    border-radius: 999px;
    overflow: hidden;
    border: 1px solid {CINZA_BORDA};
}}
.thermo-vig {{
    height: 100%;
    background: {COR_VIGENTE};
    display: inline-block;
}}
.thermo-ven {{
    height: 100%;
    background: {COR_VENCIDO};
    display: inline-block;
}}
.legend {{
    display:flex;
    gap:12px;
    flex-wrap:wrap;
    margin-top:10px;
}}
.legend .item {{
    border: 1px solid {CINZA_BORDA};
    background:#f9fafb;
    border-radius:999px;
    padding:6px 10px;
    font-weight:900;
}}
</style>
""",
    unsafe_allow_html=True,
)


# =========================================================
# HELPERS
# =========================================================
def limpar_nome_coluna(c):
    return str(c).replace("\n", " ").replace("\r", " ").strip()


def normalize_text(s: str) -> str:
    s = "" if s is None else str(s)
    s = s.replace("\u00a0", " ").replace("\u200b", "")
    s = s.strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"\s+", " ", s)
    return s.upper().strip()


def status_norm(x) -> str:
    s = normalize_text(x)
    if "VENC" in s:
        return "VENCIDO"
    if "VIG" in s:
        return "VIGENTE"
    return s


def norm_col(s):
    return normalize_text(s).lower()


def fmt_int(n: int) -> str:
    try:
        return f"{int(n):,}".replace(",", ".")
    except Exception:
        return str(n)


def fmt_pct(x: float) -> str:
    return f"{x*100:.2f}%".replace(".", ",")


def gerar_excel_bytes(df_out: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_out.to_excel(writer, index=False, sheet_name="Relatorio")
    return output.getvalue()


def selectbox_com_todos(container, label, options, key=None):
    options = [o for o in options if str(o).strip() != ""]
    options = sorted(options, key=lambda x: str(x))
    if not options:
        return "TODOS"
    return container.selectbox(label, ["TODOS"] + options, index=0, key=key)


def multiselect_com_todos(container, label, options, default_all=True, key=None):
    options = [o for o in options if str(o).strip() != ""]
    options = sorted(options, key=lambda x: str(x))
    if not options:
        return []
    todos_label = "✅ TODOS"
    ui_options = [todos_label] + options
    default = [todos_label] if default_all else []
    sel = container.multiselect(label, ui_options, default=default, key=key)
    if todos_label in sel:
        return options
    return sel


@st.cache_data(show_spinner=False)
def carregar_dados():
    if not os.path.exists(ARQUIVO):
        return pd.DataFrame()

    tmp_path = os.path.join(os.getcwd(), "~base_temp.xlsx")
    last_err = None

    for _ in range(1, 9):
        try:
            shutil.copy2(ARQUIVO, tmp_path)
            df_local = pd.read_excel(tmp_path, sheet_name=ABA, header=HEADER_LINHA, engine="openpyxl")
            df_local = df_local.dropna(axis=1, how="all").dropna(axis=0, how="all")
            df_local.columns = [limpar_nome_coluna(c) for c in df_local.columns]
            return df_local
        except PermissionError as e:
            last_err = e
            time.sleep(0.6)
        except Exception as e:
            last_err = e
            break
        finally:
            try:
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)
            except Exception:
                pass

    st.error("❌ Não consegui acessar o Excel porque ele está travado (Excel/OneDrive).")
    st.caption(f"Detalhe do erro: {repr(last_err)}")
    return pd.DataFrame()


# =========================================================
# CARREGAR BASE
# =========================================================
df = carregar_dados()
if df.empty:
    st.stop()

cols_map = {norm_col(c): c for c in df.columns}


def getcol(*cands):
    for c in cands:
        k = norm_col(c)
        if k in cols_map:
            return cols_map[k]
        for kk, orig in cols_map.items():
            if k in kk:
                return orig
    return None


COL_UNIDADE = getcol("unidade")
COL_NOME = getcol("nome")
COL_MATRICULA = getcol("matricula", "matrícula")
COL_GES = getcol("ges")
COL_SITUACAO = getcol("situacao", "situação")
COL_NORMA = getcol("norma regulamentadora", "norma", "nr")
COL_ANO = getcol("ano")
COL_DATA = getcol("data")

if COL_SITUACAO:
    df[COL_SITUACAO] = df[COL_SITUACAO].apply(status_norm)

if COL_DATA:
    df[COL_DATA] = pd.to_datetime(df[COL_DATA], errors="coerce", dayfirst=True)

if not COL_ANO and COL_DATA:
    df["ANO_DERIVADO"] = df[COL_DATA].dt.year
    COL_ANO = "ANO_DERIVADO"

# Base institucional (SEM filtros)
df_exec = df.copy()

# Chave colaborador
COL_CHAVE_COLAB = COL_MATRICULA if COL_MATRICULA else COL_NOME


# =========================================================
# SIDEBAR (BRANCA) — FILTROS (DETALHAMENTO)
# =========================================================
if os.path.exists(LOGO_PATH):
    st.sidebar.image(LOGO_PATH, use_container_width=True)

st.sidebar.markdown("## 🎛️ Filtros (Detalhamento)")
st.sidebar.caption(
    "Os indicadores do topo, termômetro e diagnóstico são institucionais (SEM filtros). "
    "Os filtros afetam apenas: tabela por NR, gráfico por ano e base detalhada."
)
st.sidebar.divider()

df_f = df.copy()

with st.sidebar.expander("🏢 Unidade", expanded=True):
    if COL_UNIDADE:
        unidades = sorted(df_f[COL_UNIDADE].dropna().astype(str).unique().tolist())
        unidade_sel = selectbox_com_todos(st, "Selecione a unidade", unidades, key="f_unidade")
        if unidade_sel != "TODOS":
            df_f = df_f[df_f[COL_UNIDADE].astype(str) == unidade_sel]

with st.sidebar.expander("📘 Norma / NR", expanded=True):
    if COL_NORMA:
        normas = sorted(df_f[COL_NORMA].dropna().astype(str).unique().tolist())
        normas_sel = multiselect_com_todos(st, "Selecione as NRs", normas, default_all=True, key="f_normas")
        if normas_sel:
            df_f = df_f[df_f[COL_NORMA].astype(str).isin(normas_sel)]

with st.sidebar.expander("✅ Situação", expanded=True):
    if COL_SITUACAO:
        situacoes = sorted(df_f[COL_SITUACAO].dropna().astype(str).unique().tolist())
        sit_sel = multiselect_com_todos(st, "Selecione a situação", situacoes, default_all=True, key="f_situacao")
        if sit_sel:
            df_f = df_f[df_f[COL_SITUACAO].astype(str).isin(sit_sel)]

with st.sidebar.expander("👥 GES", expanded=False):
    if COL_GES:
        ges_lista = sorted(df_f[COL_GES].dropna().astype(str).unique().tolist())
        ges_sel = multiselect_com_todos(st, "Selecione o GES", ges_lista, default_all=True, key="f_ges")
        if ges_sel:
            df_f = df_f[df_f[COL_GES].astype(str).isin(ges_sel)]

with st.sidebar.expander("📅 Ano", expanded=False):
    if COL_ANO and df_f[COL_ANO].notna().any():
        anos = sorted(pd.to_numeric(df_f[COL_ANO], errors="coerce").dropna().astype(int).unique().tolist())
        ano_sel = multiselect_com_todos(st, "Selecione o ano", anos, default_all=True, key="f_ano")
        if ano_sel:
            df_f = df_f[pd.to_numeric(df_f[COL_ANO], errors="coerce").isin(ano_sel)]

with st.sidebar.expander("🔎 Busca", expanded=False):
    busca = st.text_input("Buscar (NR, GES ou Situação)", value="", key="f_busca").strip().lower()
    if busca:
        cols_busca = [c for c in [COL_NORMA, COL_GES, COL_SITUACAO] if c]
        mask = False
        for c in cols_busca:
            mask = mask | df_f[c].astype(str).str.lower().str.contains(busca, na=False)
        df_f = df_f[mask]

st.sidebar.divider()
if st.sidebar.button("🔄 Atualizar dados", use_container_width=True):
    st.cache_data.clear()
    st.rerun()

# Base filtrada (somente para as 3 seções solicitadas)
df_det = df_f.copy()


# =========================================================
# HEADER (TOPO)
# =========================================================
h1, h2 = st.columns([0.12, 0.88], vertical_alignment="center")
with h1:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, use_container_width=True)
with h2:
    st.markdown(
        f"""
<div class="header-wrap">
  <p class="h-title">Dashboard Executivo - Treinamentos SENAI</p>
  <p class="h-sub">Indicadores institucionais • Termômetro • Unidade Ouro • Ranking • Análises</p>
  <span class="badge">Meta institucional: 100% VIGENTE</span>
</div>
""",
        unsafe_allow_html=True,
    )

st.write("")


# =========================================================
# DIAGNÓSTICO (ANTES DOS FILTROS) — BASE INSTITUCIONAL
# =========================================================
with st.expander("🔎 Diagnóstico rápido da SITUAÇÃO (antes dos filtros)", expanded=True):
    if COL_SITUACAO:
        diag = df_exec[COL_SITUACAO].value_counts(dropna=False).reset_index()
        diag.columns = ["SITUAÇÃO", "count"]
        st.dataframe(diag, use_container_width=True, hide_index=True)
    else:
        st.warning("Coluna SITUAÇÃO não encontrada.")


# =========================================================
# KPIs (INSTITUCIONAIS) — SEM FILTROS
# =========================================================
total_exec_registros = len(df_exec)

total_colaboradores_exec = (
    df_exec[COL_CHAVE_COLAB].dropna().astype(str).nunique()
    if COL_CHAVE_COLAB
    else 0
)

vigente_exec = len(df_exec[df_exec[COL_SITUACAO] == "VIGENTE"]) if COL_SITUACAO else 0
vencido_exec = len(df_exec[df_exec[COL_SITUACAO] == "VENCIDO"]) if COL_SITUACAO else 0

total_status_exec = (vigente_exec + vencido_exec) if (vigente_exec + vencido_exec) > 0 else total_exec_registros

pct_vigente_exec = (vigente_exec / total_status_exec) if total_status_exec else 0.0
pct_vencido_exec = (vencido_exec / total_status_exec) if total_status_exec else 0.0

st.markdown(
    f"""
<div class="kpi-grid">
  <div class="kpi">
    <h4>Total de Colaboradores</h4>
    <div class="v">{fmt_int(total_colaboradores_exec)}</div>
    <div class="hint">Nomes únicos (sem duplicados)</div>
  </div>
  <div class="kpi">
    <h4>Meta (100% VIGENTE)</h4>
    <div class="v">{fmt_int(total_status_exec)}</div>
    <div class="hint">Meta aplicada ao total (VIGENTE+VENCIDO)</div>
  </div>
  <div class="kpi">
    <h4>VIGENTE</h4>
    <div class="v" style="color:{COR_VIGENTE};">{fmt_int(vigente_exec)}</div>
    <div class="hint">{fmt_pct(pct_vigente_exec)}</div>
  </div>
  <div class="kpi">
    <h4>VENCIDO</h4>
    <div class="v" style="color:{COR_VENCIDO};">{fmt_int(vencido_exec)}</div>
    <div class="hint">{fmt_pct(pct_vencido_exec)}</div>
  </div>
</div>
""",
    unsafe_allow_html=True,
)

st.write("")


# =========================================================
# TERMÔMETRO — INSTITUCIONAL (SEM FILTROS)
# =========================================================
st.markdown(
    f"""
<div class="thermo-wrap">
  <div class="section-title">
    <h3>🌡️ Termômetro Institucional (VIGENTE x VENCIDO)</h3>
    <div>
      <span class="pill">🕒 {datetime.now().strftime('%d/%m/%Y %H:%M')}</span>
      <span class="pill">🎯 Meta: 100%</span>
      <span class="pill">Atingimento: {fmt_pct(pct_vigente_exec)}</span>
    </div>
  </div>

  <div class="thermo-bar">
    <span class="thermo-vig" style="width:{pct_vigente_exec*100:.4f}%"></span>
    <span class="thermo-ven" style="width:{pct_vencido_exec*100:.4f}%"></span>
  </div>

  <div class="legend">
    <div class="item">🔷 VIGENTE: {fmt_pct(pct_vigente_exec)} ({fmt_int(vigente_exec)})</div>
    <div class="item">🔶 VENCIDO: {fmt_pct(pct_vencido_exec)} ({fmt_int(vencido_exec)})</div>
  </div>
</div>
""",
    unsafe_allow_html=True,
)

st.write("")


# =========================================================
# UNIDADE OURO — TODO PERÍODO (SEM FILTROS)
# =========================================================
st.markdown(
    """
<div class="section">
  <div class="section-title">
    <h3>🏅 Unidade Ouro (Todo período)</h3>
    <div><span class="pill">Maior % atingida (VIGENTE/TOTAL)</span></div>
  </div>
</div>
""",
    unsafe_allow_html=True,
)

ouro_info = None
if COL_UNIDADE and COL_CHAVE_COLAB and COL_SITUACAO and not df_exec.empty:
    base = df_exec.copy()

    agg = base.groupby(COL_UNIDADE).agg(
        EFETIVO=(COL_CHAVE_COLAB, lambda s: s.dropna().astype(str).nunique()),
        VIGENTE=(COL_SITUACAO, lambda s: (s == "VIGENTE").sum()),
        VENCIDO=(COL_SITUACAO, lambda s: (s == "VENCIDO").sum()),
    ).reset_index()

    agg["TOTAL_TREIN"] = (agg["VIGENTE"] + agg["VENCIDO"]).astype(int)
    agg = agg[agg["TOTAL_TREIN"] > 0].copy()

    agg["% ATINGIDA"] = (agg["VIGENTE"] / agg["TOTAL_TREIN"]).fillna(0) * 100
    agg = agg.sort_values(["% ATINGIDA", "VIGENTE", "EFETIVO"], ascending=[False, False, False])
    if not agg.empty:
        ouro_info = agg.iloc[0].to_dict()

if ouro_info:
    unidade_ouro = ouro_info.get(COL_UNIDADE, "—")
    efetivo_ouro = ouro_info.get("EFETIVO", 0)
    vigente_ouro = ouro_info.get("VIGENTE", 0)
    vencido_ouro = ouro_info.get("VENCIDO", 0)
    total_ouro = ouro_info.get("TOTAL_TREIN", 0)
    pct_ouro = ouro_info.get("% ATINGIDA", 0.0)

    st.markdown(
        f"""
<div class="kpi-grid" style="grid-template-columns: repeat(4, minmax(0, 1fr));">
  <div class="kpi"><h4>Unidade</h4><div class="v">{unidade_ouro}</div><div class="hint">Todo período</div></div>
  <div class="kpi"><h4>Efetivo</h4><div class="v">{fmt_int(efetivo_ouro)}</div><div class="hint">Colaboradores únicos</div></div>
  <div class="kpi"><h4>VIGENTE / VENCIDO</h4><div class="v"><span style="color:{COR_VIGENTE};">{fmt_int(vigente_ouro)}</span> / <span style="color:{COR_VENCIDO};">{fmt_int(vencido_ouro)}</span></div><div class="hint">Total: {fmt_int(total_ouro)}</div></div>
  <div class="kpi"><h4>% Atingida</h4><div class="v">{str(round(pct_ouro,2)).replace(".",",")}%</div><div class="hint">VIGENTE/TOTAL</div></div>
</div>
""",
        unsafe_allow_html=True,
    )
else:
    st.info("Sem base suficiente para calcular a Unidade Ouro.")

st.write("")


# =========================================================
# RANKING TOP 5 — POR % ATINGIDA (SEM FILTROS)
# =========================================================
st.markdown(
    """
<div class="section">
  <div class="section-title">
    <h3>🏆 Ranking Top 5 Unidades (por % atingida)</h3>
    <div><span class="pill">Período total • % = VIGENTE / (VIGENTE+VENCIDO)</span></div>
  </div>
</div>
""",
    unsafe_allow_html=True,
)

if COL_UNIDADE and COL_CHAVE_COLAB and COL_SITUACAO:
    base = df_exec.copy()

    grp_u = base.groupby(COL_UNIDADE).agg(
        EFETIVO=(COL_CHAVE_COLAB, lambda s: s.dropna().astype(str).nunique()),
        VIGENTE=(COL_SITUACAO, lambda s: (s == "VIGENTE").sum()),
        VENCIDO=(COL_SITUACAO, lambda s: (s == "VENCIDO").sum()),
    ).reset_index()

    grp_u["TOTAL_TREIN"] = (grp_u["VIGENTE"] + grp_u["VENCIDO"]).astype(int)
    grp_u = grp_u[grp_u["TOTAL_TREIN"] > 0].copy()

    grp_u["% ATINGIDA"] = (grp_u["VIGENTE"] / grp_u["TOTAL_TREIN"]).fillna(0) * 100

    grp_u = grp_u.sort_values(
        ["% ATINGIDA", "VIGENTE", "TOTAL_TREIN", "EFETIVO"],
        ascending=[False, False, False, False],
    )

    top5 = grp_u.head(5).reset_index(drop=True)
    top5["POSIÇÃO"] = top5.index + 1
    top5["MEDALHA"] = top5["POSIÇÃO"].map({1: "🥇", 2: "🥈", 3: "🥉"}).fillna("⭐")
    top5 = top5.rename(columns={COL_UNIDADE: "UNIDADE"})

    st.dataframe(
        top5[["POSIÇÃO", "MEDALHA", "UNIDADE", "EFETIVO", "TOTAL_TREIN", "VIGENTE", "VENCIDO", "% ATINGIDA"]],
        use_container_width=True,
        hide_index=True,
    )
else:
    st.info("Não foi possível gerar ranking (verifique colunas UNIDADE + NOME/MATRICULA + SITUAÇÃO).")

st.write("")


# =========================================================
# VIGENTE/VENCIDO POR NR — COM FILTROS (df_det)
# =========================================================
st.markdown(
    """
<div class="section">
  <div class="section-title">
    <h3>📘 VIGENTE x VENCIDO por Treinamento/NR</h3>
    <div><span class="pill">Com filtros do menu lateral</span></div>
  </div>
</div>
""",
    unsafe_allow_html=True,
)

if COL_NORMA and COL_SITUACAO and not df_det.empty:
    nr = df_det.groupby(COL_NORMA).agg(
        VIGENTE=(COL_SITUACAO, lambda s: (s == "VIGENTE").sum()),
        VENCIDO=(COL_SITUACAO, lambda s: (s == "VENCIDO").sum()),
    ).reset_index()

    nr["TOTAL"] = (nr["VIGENTE"] + nr["VENCIDO"]).astype(int)
    nr = nr[nr["TOTAL"] > 0].copy()

    nr["% VIGENTE"] = (nr["VIGENTE"] / nr["TOTAL"]).fillna(0) * 100
    nr["% VENCIDO"] = (nr["VENCIDO"] / nr["TOTAL"]).fillna(0) * 100

    nr = nr.sort_values(["% VENCIDO", "VENCIDO"], ascending=[False, False]).rename(columns={COL_NORMA: "NORMA"})

    st.dataframe(
        nr[["NORMA", "TOTAL", "VIGENTE", "VENCIDO", "% VIGENTE", "% VENCIDO"]],
        use_container_width=True,
        hide_index=True,
        height=360,
    )
else:
    st.info("Não foi possível montar a visão por NR com os filtros (verifique colunas NORMA e SITUAÇÃO).")

st.write("")


# =========================================================
# REGISTROS POR ANO — COM FILTROS (df_det)
# (com valores em cima das barras)
# =========================================================
st.markdown(
    """
<div class="section">
  <div class="section-title">
    <h3>📅 Quantidade de Registros por Ano</h3>
    <div><span class="pill">Com filtros do menu lateral</span></div>
  </div>
</div>
""",
    unsafe_allow_html=True,
)

if COL_ANO and not df_det.empty and df_det[COL_ANO].notna().any():
    anos_series = pd.to_numeric(df_det[COL_ANO], errors="coerce").dropna().astype(int)
    if not anos_series.empty:
        cont_ano = anos_series.value_counts().sort_index()

        fig = plt.figure(figsize=(10, 3.8))
        ax = plt.gca()

        bars = ax.bar(cont_ano.index.astype(str), cont_ano.values)

        ax.set_xlabel("Ano")
        ax.set_ylabel("Quantidade de registros")

        for b in bars:
            h = b.get_height()
            ax.annotate(
                f"{int(h)}",
                xy=(b.get_x() + b.get_width() / 2, h),
                xytext=(0, 4),
                textcoords="offset points",
                ha="center",
                va="bottom",
                fontsize=10,
                fontweight="bold",
            )

        plt.tight_layout()
        st.pyplot(fig, clear_figure=True)
    else:
        st.info("Não foi possível identificar anos válidos com os filtros.")
else:
    st.info("Coluna ANO não encontrada (nem derivada pela DATA) ou não há dados após filtros.")

st.write("")


# =========================================================
# BASE DETALHADA + EXPORTAÇÃO — COM FILTROS (df_det)
# =========================================================
st.markdown(
    """
<div class="section">
  <div class="section-title">
    <h3>🔎 Base Detalhada (com filtros) e Exportação</h3>
    <div><span class="pill">Filtros do menu lateral</span></div>
  </div>
</div>
""",
    unsafe_allow_html=True,
)

st.dataframe(df_det, use_container_width=True, height=420)

st.download_button(
    "⬇️ Baixar relatório filtrado (Excel)",
    data=gerar_excel_bytes(df_det),
    file_name="Relatorio_Treinamentos_Filtrado.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)

with st.expander("🔧 Diagnóstico (colunas detectadas)", expanded=False):
    st.write("Colunas encontradas:", list(df.columns))
    st.write({
        "UNIDADE": COL_UNIDADE,
        "NOME": COL_NOME,
        "MATRICULA": COL_MATRICULA,
        "CHAVE_COLAB": COL_CHAVE_COLAB,
        "GES": COL_GES,
        "NORMA": COL_NORMA,
        "SITUAÇÃO": COL_SITUACAO,
        "ANO": COL_ANO,
        "DATA": COL_DATA,
    })
