# =========================================================
# app_corrigido.py — Dashboard Executivo (SENAI) — CONSOLIDADO
# CORREÇÃO DEFINITIVA DOS "ZERADOS":
# 1) Usa os nomes EXATOS das colunas da sua planilha (linha 4)
# 2) Normaliza textos (UNIDADE / NR / SITUAÇÃO) com strip + upper
# 3) Tabela por NR usa crosstab (não depende de igualdade “perfeita”)
# 4) TOTAL por NR = quantidade de registros (size) -> nunca zera se há linhas
# =========================================================

import os
import io
import time
import shutil
import unicodedata
import re
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
HEADER_LINHA = 3  # linha 4

LOGO_PATH = "logo_senai.png"

# Paleta SENAI
SENAI_LARANJA = "#F37021"
SENAI_AZUL = "#003DA5"
CINZA_FUNDO = "#F5F6F8"
CINZA_BORDA = "#E5E7EB"
TEXTO_CINZA = "#6B7280"

COR_VIGENTE = SENAI_AZUL
COR_VENCIDO = SENAI_LARANJA

# =========================================================
# CSS
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
def normalize_text(s: str) -> str:
    s = "" if s is None else str(s)
    s = s.replace("\u00a0", " ").replace("\u200b", "")
    s = s.strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"\s+", " ", s)
    return s.upper().strip()

def fmt_int(n) -> str:
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
        st.error(f"❌ Arquivo não encontrado: {ARQUIVO}")
        return pd.DataFrame()

    tmp_path = os.path.join(os.getcwd(), "~base_temp.xlsx")
    last_err = None

    for _ in range(1, 9):
        try:
            shutil.copy2(ARQUIVO, tmp_path)
            df_local = pd.read_excel(tmp_path, sheet_name=ABA, header=HEADER_LINHA, engine="openpyxl")
            df_local = df_local.dropna(axis=1, how="all").dropna(axis=0, how="all")
            # padroniza nomes das colunas
            df_local.columns = [str(c).replace("\n", " ").replace("\r", " ").strip() for c in df_local.columns]
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

# =========================================================
# COLUNAS (EXATAS DO SEU EXCEL)
# =========================================================
COL_UNIDADE  = "UNIDADE"
COL_NOME     = "NOME"
COL_MATRICULA= "MATRÍCULA"
COL_GES      = "GES"
COL_NORMA    = "NORMA REGULAMENTADORA"
COL_TREIN    = "TREINAMENTOS DO GES"
COL_ANO      = "ANO"
COL_DATA     = "DATA ATUAL"
COL_SITUACAO = "SITUAÇÃO"

# garante que existem
for c in [COL_UNIDADE, COL_NOME, COL_MATRICULA, COL_GES, COL_NORMA, COL_ANO, COL_SITUACAO]:
    if c not in df.columns:
        st.error(f"❌ Coluna obrigatória não encontrada: {c}")
        st.write("Colunas detectadas:", list(df.columns))
        st.stop()

# =========================================================
# NORMALIZAÇÃO (aqui fica “à prova de erro”)
# =========================================================
df[COL_UNIDADE] = df[COL_UNIDADE].astype(str).str.strip()
df[COL_NORMA]   = df[COL_NORMA].astype(str).str.strip()
df[COL_TREIN]   = df[COL_TREIN].astype(str).str.strip() if COL_TREIN in df.columns else ""

# ✅ Chave única para a tabela "Treinamento/NR":
# - Se "NORMA REGULAMENTADORA" estiver vazia, usa "TREINAMENTOS DO GES"
# - Isso evita VENCIDOS "sumirem" quando a coluna de NR está em branco nas linhas vencidas
def _vazio(x: str) -> bool:
    x = "" if x is None else str(x)
    x = x.strip().upper()
    return x in ("", "NAN", "NONE", "NULL", "-")

CHAVE_TREIN_NR = "TREIN_NR"
df[CHAVE_TREIN_NR] = df.apply(
    lambda r: (r.get(COL_NORMA, "") if not _vazio(r.get(COL_NORMA, "")) else r.get(COL_TREIN, "SEM TREINAMENTO")),
    axis=1,
)
df[CHAVE_TREIN_NR] = df[CHAVE_TREIN_NR].astype(str).str.strip()
df[COL_SITUACAO]= df[COL_SITUACAO].astype(str).apply(normalize_text)  # VIGENTE/VENCIDO

# se vier qualquer variação (ex.: "VENCIDO " ou "VENC."), garante:
df.loc[df[COL_SITUACAO].str.contains("VENC", na=False), COL_SITUACAO] = "VENCIDO"
df.loc[~df[COL_SITUACAO].isin(["VIGENTE", "VENCIDO"]), COL_SITUACAO] = "VIGENTE"  # fallback seguro

# Data/ano
if COL_DATA in df.columns:
    df[COL_DATA] = pd.to_datetime(df[COL_DATA], errors="coerce", dayfirst=True)

# Base institucional (SEM filtros)
df_exec = df.copy()

# Chave colaborador
COL_CHAVE_COLAB = COL_MATRICULA if COL_MATRICULA in df.columns else COL_NOME

# =========================================================
# SIDEBAR — FILTROS (DETALHAMENTO)
# =========================================================
if os.path.exists(LOGO_PATH):
    st.sidebar.image(LOGO_PATH, width="stretch")

st.sidebar.markdown("## 🎛️ Filtros (Detalhamento)")
st.sidebar.caption(
    "Os indicadores do topo, termômetro e diagnóstico são institucionais (SEM filtros). "
    "Os filtros afetam apenas: tabela por NR, gráfico por ano e base detalhada."
)
st.sidebar.divider()
# ✅ Botão para resetar filtros (evita estado antigo prender somente VIGENTE)
if st.sidebar.button("🧹 Limpar filtros (reset)", width="stretch"):
    for k in ["f_unidade", "f_normas", "f_situacao_v2", "f_ges", "f_ano", "f_busca"]:
        if k in st.session_state:
            del st.session_state[k]
    st.rerun()

df_base = df.copy()  # filtros sem SITUAÇÃO (para tabela por NR)
df_det = df.copy()   # filtros com SITUAÇÃO (para base detalhada/export)

with st.sidebar.expander("🏢 Unidade", expanded=True):
    unidades = sorted(df_det[COL_UNIDADE].dropna().astype(str).unique().tolist())
    unidade_sel = selectbox_com_todos(st, "Selecione a unidade", unidades, key="f_unidade")
    if unidade_sel != "TODOS":
        df_base = df_base[df_base[COL_UNIDADE].astype(str) == unidade_sel]
        df_det  = df_det[df_det[COL_UNIDADE].astype(str) == unidade_sel]

with st.sidebar.expander("📘 Norma / NR", expanded=True):
    normas = sorted(df_det[COL_NORMA].dropna().astype(str).unique().tolist())
    normas_sel = multiselect_com_todos(st, "Selecione as NRs", normas, default_all=True, key="f_normas")
    if normas_sel:
        df_base = df_base[df_base[COL_NORMA].astype(str).isin(normas_sel)]
        df_det  = df_det[df_det[COL_NORMA].astype(str).isin(normas_sel)]

with st.sidebar.expander("✅ Situação", expanded=True):
    situacoes = sorted(df_det[COL_SITUACAO].dropna().astype(str).unique().tolist())
    sit_sel = multiselect_com_todos(st, "Selecione a situação", situacoes, default_all=True, key="f_situacao_v2")
    if sit_sel:
        df_det = df_det[df_det[COL_SITUACAO].astype(str).isin(sit_sel)]

with st.sidebar.expander("👥 GES", expanded=False):
    ges_lista = sorted(df_det[COL_GES].dropna().astype(str).unique().tolist())
    ges_sel = multiselect_com_todos(st, "Selecione o GES", ges_lista, default_all=True, key="f_ges")
    if ges_sel:
        df_base = df_base[df_base[COL_GES].astype(str).isin(ges_sel)]
        df_det  = df_det[df_det[COL_GES].astype(str).isin(ges_sel)]

with st.sidebar.expander("📅 Ano", expanded=False):
    anos = sorted(pd.to_numeric(df_det[COL_ANO], errors="coerce").dropna().astype(int).unique().tolist())
    ano_sel = multiselect_com_todos(st, "Selecione o ano", anos, default_all=True, key="f_ano")

    # ✅ CORREÇÃO: a maioria dos VENCIDOS está com ANO vazio (NaN).
    # Se filtrarmos apenas pelos anos selecionados, removemos os vencidos sem ano.
    # Portanto, quando houver filtro de ano, mantemos também os registros sem ano.
    if ano_sel:
        ano_base = pd.to_numeric(df_base[COL_ANO], errors="coerce")
        ano_det  = pd.to_numeric(df_det[COL_ANO], errors="coerce")
        mask_base = ano_base.isna() | ano_base.isin(ano_sel)
        mask_det  = ano_det.isna()  | ano_det.isin(ano_sel)
        df_base = df_base[mask_base]
        df_det  = df_det[mask_det]


with st.sidebar.expander("🔎 Busca", expanded=False):
    busca = st.text_input("Buscar (NR, GES ou Situação)", value="", key="f_busca").strip().lower()
    if busca:
        mask_base = (
            df_base[COL_NORMA].astype(str).str.lower().str.contains(busca, na=False)
            | df_base[COL_GES].astype(str).str.lower().str.contains(busca, na=False)
            | df_base[COL_SITUACAO].astype(str).str.lower().str.contains(busca, na=False)
        )
        df_base = df_base[mask_base]

        mask_det = (
            df_det[COL_NORMA].astype(str).str.lower().str.contains(busca, na=False)
            | df_det[COL_GES].astype(str).str.lower().str.contains(busca, na=False)
            | df_det[COL_SITUACAO].astype(str).str.lower().str.contains(busca, na=False)
        )
        df_det = df_det[mask_det]

st.sidebar.divider()
if st.sidebar.button("🔄 Atualizar dados", width="stretch"):
    st.cache_data.clear()
    st.rerun()

# =========================================================
# HEADER (TOPO)
# =========================================================
h1, h2 = st.columns([0.12, 0.88], vertical_alignment="center")
with h1:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width="stretch")
with h2:
    st.markdown(
        """
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
# DIAGNÓSTICO (ANTES DOS FILTROS) — INSTITUCIONAL
# =========================================================
with st.expander("🔎 Diagnóstico rápido da SITUAÇÃO (antes dos filtros)", expanded=True):
    diag = df_exec[COL_SITUACAO].value_counts(dropna=False).reset_index()
    diag.columns = ["SITUAÇÃO", "count"]
    st.dataframe(diag, width="stretch", hide_index=True)

# =========================================================
# KPIs (INSTITUCIONAIS) — SEM FILTROS
# =========================================================
total_colaboradores_exec = df_exec[COL_CHAVE_COLAB].dropna().astype(str).nunique()
vigente_exec = int((df_exec[COL_SITUACAO] == "VIGENTE").sum())
vencido_exec = int((df_exec[COL_SITUACAO] == "VENCIDO").sum())
total_status_exec = vigente_exec + vencido_exec

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
# TERMÔMETRO — INSTITUCIONAL
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
# UNIDADE OURO + TOP 5 (SEM FILTROS)
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

grp = df_exec.groupby(COL_UNIDADE).agg(
    EFETIVO=(COL_CHAVE_COLAB, lambda s: s.dropna().astype(str).nunique()),
    VIGENTE=(COL_SITUACAO, lambda s: (s == "VIGENTE").sum()),
    VENCIDO=(COL_SITUACAO, lambda s: (s == "VENCIDO").sum()),
).reset_index()

grp["TOTAL_TREIN"] = (grp["VIGENTE"] + grp["VENCIDO"]).astype(int)
grp = grp[grp["TOTAL_TREIN"] > 0].copy()
grp["% ATINGIDA"] = (grp["VIGENTE"] / grp["TOTAL_TREIN"]).fillna(0) * 100
grp = grp.sort_values(["% ATINGIDA", "VIGENTE", "EFETIVO"], ascending=[False, False, False])

if not grp.empty:
    ouro = grp.iloc[0]
    st.markdown(
        f"""
<div class="kpi-grid" style="grid-template-columns: repeat(4, minmax(0, 1fr));">
  <div class="kpi"><h4>Unidade</h4><div class="v">{ouro[COL_UNIDADE]}</div><div class="hint">Todo período</div></div>
  <div class="kpi"><h4>Efetivo</h4><div class="v">{fmt_int(ouro['EFETIVO'])}</div><div class="hint">Colaboradores únicos</div></div>
  <div class="kpi"><h4>VIGENTE / VENCIDO</h4><div class="v"><span style="color:{COR_VIGENTE};">{fmt_int(ouro['VIGENTE'])}</span> / <span style="color:{COR_VENCIDO};">{fmt_int(ouro['VENCIDO'])}</span></div><div class="hint">Total: {fmt_int(ouro['TOTAL_TREIN'])}</div></div>
  <div class="kpi"><h4>% Atingida</h4><div class="v">{str(round(float(ouro['% ATINGIDA']),2)).replace('.',',')}%</div><div class="hint">VIGENTE/TOTAL</div></div>
</div>
""",
        unsafe_allow_html=True,
    )
else:
    st.info("Sem base suficiente para calcular a Unidade Ouro.")

st.write("")

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

top5 = grp.head(5).reset_index(drop=True).copy()
top5["POSIÇÃO"] = top5.index + 1
top5["MEDALHA"] = top5["POSIÇÃO"].map({1: "🥇", 2: "🥈", 3: "🥉"}).fillna("⭐")
top5 = top5.rename(columns={COL_UNIDADE: "UNIDADE"})
st.dataframe(
    top5[["POSIÇÃO", "MEDALHA", "UNIDADE", "EFETIVO", "TOTAL_TREIN", "VIGENTE", "VENCIDO", "% ATINGIDA"]],
    width="stretch",
    hide_index=True,
)
st.write("")

# =========================================================
# ✅ VIGENTE x VENCIDO por NR — COM FILTROS (SEM filtro de SITUAÇÃO) — SEM ZERAR
# (crosstab é mais robusto)
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

if df_base.empty:
    st.warning("Sem dados após filtros. Ajuste os filtros do menu lateral.")
else:
    tab = pd.crosstab(df_base[CHAVE_TREIN_NR], df_base[COL_SITUACAO]).reset_index().rename(columns={CHAVE_TREIN_NR: "TREINAMENTO/NR"})

    # garante colunas (caso uma situação não apareça na base filtrada)
    if "VIGENTE" not in tab.columns:
        tab["VIGENTE"] = 0
    if "VENCIDO" not in tab.columns:
        tab["VENCIDO"] = 0

    tab["TOTAL"] = (tab["VIGENTE"] + tab["VENCIDO"]).astype(int)
    tab = tab[tab["TOTAL"] > 0].copy()
    tab["% VIGENTE"] = (tab["VIGENTE"] / tab["TOTAL"]).fillna(0) * 100
    tab["% VENCIDO"] = (tab["VENCIDO"] / tab["TOTAL"]).fillna(0) * 100

    tab = tab.sort_values(["% VENCIDO", "VENCIDO"], ascending=[False, False])

    st.dataframe(
        tab[["TREINAMENTO/NR", "TOTAL", "VIGENTE", "VENCIDO", "% VIGENTE", "% VENCIDO"]],
        width="stretch",
        hide_index=True,
        height=360,
    )

st.write("")
# =========================================================
# REGISTROS POR ANO — COM FILTROS (df_det)
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

anos_series = pd.to_numeric(df_det[COL_ANO], errors="coerce").dropna().astype(int) if not df_det.empty else pd.Series([], dtype=int)

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
    st.info("Sem anos válidos após filtros.")

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

st.dataframe(df_det, width="stretch", height=420)

st.download_button(
    "⬇️ Baixar relatório filtrado (Excel)",
    data=gerar_excel_bytes(df_det),
    file_name="Relatorio_Treinamentos_Filtrado.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    width="stretch",
)

with st.expander("🔧 Diagnóstico (colunas detectadas)", expanded=False):
    st.write("Colunas detectadas:", list(df.columns))
    st.write("Linhas após filtros:", len(df_det))
    st.write("Amostra SITUAÇÃO (pós-normalização):", df[COL_SITUACAO].value_counts().to_dict())
