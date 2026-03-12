import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime, timedelta, date
import io

# ─────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Gestão Integrada · Cronograma",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─────────────────────────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@300;400;500;600;700&display=swap');
html, body, [class*="css"] { font-family: 'IBM Plex Sans', sans-serif !important; }

.stApp { background: #F0F2F8; }
.block-container { padding-top: 20px !important; padding-bottom: 40px !important; }

.app-header {
    background: #1C1F2E; border-radius: 10px;
    padding: 16px 26px; margin-bottom: 16px;
}
.app-title { color: #FFFFFF; font-size: 18px; font-weight: 700; }
.app-sub   { color: #8B92B8; font-size: 12px; margin-top: 3px; }

.filter-card {
    background: #FFFFFF; border: 1px solid #D8DCE8; border-radius: 10px;
    padding: 14px 20px 10px 20px; margin-bottom: 14px;
    box-shadow: 0 1px 4px rgba(0,0,0,.06);
}
.filter-badge {
    font-size: 11px; font-weight: 700; text-transform: uppercase;
    letter-spacing: .09em; color: #1C1F2E; background: #E8EBF4;
    padding: 3px 10px; border-radius: 4px; display: inline-block; margin-bottom: 10px;
}

label, .stMultiSelect label, .stSelectbox label,
.stDateInput label, .stRadio label, [data-testid="stWidgetLabel"] {
    color: #1C1F2E !important; font-size: 12px !important; font-weight: 600 !important;
}

.mcard {
    background: white; border: 1px solid #D8DCE8; border-radius: 10px;
    padding: 13px 16px; box-shadow: 0 1px 3px rgba(0,0,0,.05);
}
.mcard-num { font-size: 26px; font-weight: 700; line-height: 1; }
.mcard-lbl { font-size: 10px; color: #6B7394; text-transform: uppercase;
             letter-spacing: .08em; margin-top: 4px; }
.mcard-sub { font-size: 11px; color: #9BA3BF; margin-top: 2px; }

.legend-wrap { display: flex; flex-wrap: wrap; gap: 6px; padding: 10px 4px 4px 4px; }
.legend-pill {
    display: inline-flex; align-items: center; gap: 6px;
    background: white; border: 1px solid #D8DCE8; border-radius: 20px;
    padding: 3px 10px 3px 6px; font-size: 11px; color: #1C1F2E;
    font-weight: 500; white-space: nowrap;
}
.legend-dot { width: 10px; height: 10px; border-radius: 50%; flex-shrink: 0; display: inline-block; }

div[data-testid="stHorizontalBlock"] { gap: 10px; }

/* ── Fix tab label colors ── */
button[data-baseweb="tab"] p,
button[data-baseweb="tab"] span,
button[data-baseweb="tab"] div {
    color: #1C1F2E !important;
}
button[data-baseweb="tab"][aria-selected="true"] p,
button[data-baseweb="tab"][aria-selected="true"] span {
    color: #4F46E5 !important;
    font-weight: 700 !important;
}
button[data-baseweb="tab"]:hover p,
button[data-baseweb="tab"]:hover span {
    color: #4F46E5 !important;
}

/* ── Toggle buttons (expand/collapse initiatives): force white text ── */
button[kind="secondary"] p,
button[kind="secondary"] span,
[data-testid="baseButton-secondary"] p,
[data-testid="baseButton-secondary"] span {
    color: #FFFFFF !important;
}

/* ── Fix text only on light backgrounds (captions, expanders, markdown) ── */
[data-testid="stCaption"] p,
[data-testid="stCaption"] span,
[data-testid="stMarkdownContainer"] > p,
[data-testid="stMarkdownContainer"] > ul > li,
[data-testid="stExpander"] summary p,
[data-testid="stExpander"] summary span,
[data-testid="stExpander"] > div[data-testid="stExpanderDetails"] p,
[data-testid="stExpander"] > div[data-testid="stExpanderDetails"] span {
    color: #1C1F2E !important;
}

/* Keep header white */
.app-title { color: #FFFFFF !important; }
.app-sub   { color: #8B92B8 !important; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────────
INI_PALETTE = [
    "#C0392B","#E67E22","#D4AC0D","#27AE60","#1A8FD1",
    "#8E44AD","#16A085","#2E4053","#CB4335","#1F618D",
    "#117A65","#884EA0","#B7950B","#1E8449","#6E2FD1",
    "#A04000","#1A5276","#0E6655","#6D4C41","#37474F","#AD1457",
]
RESP_UG_PALETTE = {
    "Felipe":"#1A8FD1","Marcelo":"#27AE60","Ana Clara":"#E67E22",
    "Priscila":"#AD1457","Marcia":"#8E44AD","Marcia Torquatto":"#8E44AD",
    "Paula":"#D97706","-":"#9BA3BF",
}
RESP_IL_PALETTE = {
    "Uales":"#16A085","Ryan":"#C0392B","Ryan e Uales":"#D35400",
    "Ryan ":"#C0392B","-":"#9BA3BF",
}
# Combined palette for "Responsável (atividade)" — all known names
RESP_ATY_PALETTE = {
    "Felipe":"#1A8FD1","Marcelo":"#27AE60","Ana Clara":"#E67E22",
    "Priscila":"#AD1457","Marcia":"#8E44AD","Marcia Torquatto":"#8E44AD",
    "Paula":"#D97706","Uales":"#16A085","Ryan":"#C0392B",
    "Bruno Mazzoco":"#6E2FD1","-":"#9BA3BF",
}
SP_COLORS = {"1":"#4F46E5","2":"#0891B2","3":"#059669","4":"#D97706"}
FERR_COLORS         = {"Ferramenta":"#7C3AED", "-":"#9BA3BF"}
PILOTO_ILOS_COLORS  = {"SIM":"#059669", "NÃO":"#DC2626", "A DEFINIR":"#D97706", "-":"#9BA3BF"}
PILOTO_ROLLOUT_COLORS = {"Piloto 1":"#7C3AED", "Piloto 2":"#1442E6", "Roll out":"#0ED055", "Geral":"#B3730D", "-":"#9BA3BF"}
# Combined: (piloto_ilos, piloto) → color
PILOTO_COMB_COLORS  = {
    ("SIM","SIM"):       "#059669",  # verde    — ambos piloto
    ("NÃO","NÃO"):       "#DC2626",  # vermelho — nenhum piloto
    ("SIM","NÃO"):       "#2563EB",  # azul     — só ILOS piloto
    ("NÃO","SIM"):       "#F97316",  # laranja  — só UG piloto
    ("A DEFINIR","SIM"): "#A855F7",  # roxo     — ILOS a definir, UG sim
    ("SIM","A DEFINIR"): "#0EA5E9",  # azul-ceu — ILOS sim, UG a definir
    ("A DEFINIR","NÃO"): "#F43F5E",  # rosa     — ILOS a definir, UG não
    ("NÃO","A DEFINIR"): "#FB923C",  # laranja-claro — ILOS não, UG a definir
    ("A DEFINIR","A DEFINIR"): "#D97706",  # âmbar — ambos a definir
}
PILOTO_COMB_LABELS  = {
    ("SIM","SIM"):       "Piloto ILOS + UG",
    ("NÃO","NÃO"):       "Não piloto",
    ("SIM","NÃO"):       "Piloto só ILOS",
    ("NÃO","SIM"):       "Piloto só UG",
    ("A DEFINIR","SIM"): "ILOS a definir · UG sim",
    ("SIM","A DEFINIR"): "ILOS sim · UG a definir",
    ("A DEFINIR","NÃO"): "ILOS a definir · UG não",
    ("NÃO","A DEFINIR"): "ILOS não · UG a definir",
    ("A DEFINIR","A DEFINIR"): "Ambos a definir",
}
PILOTO_COLORS       = {"SIM":"#2563EB", "NÃO":"#F43F5E", "A DEFINIR":"#F59E0B", "-":"#9BA3BF"}
PILOTO_COLORS       = {"SIM":"#059669", "NÃO":"#DC2626", "A DEFINIR":"#D97706", "-":"#9BA3BF"}
PILOTO_COLORS       = {"SIM":"#059669", "NÃO":"#DC2626", "A DEFINIR":"#D97706", "-":"#9BA3BF"}

PLOTLY_CFG = {
    "scrollZoom": False,
    "displayModeBar": True,
    "modeBarButtonsToRemove": ["zoom2d","zoomIn2d","zoomOut2d","autoScale2d","resetScale2d"],
    "displaylogo": False,
}

# ─────────────────────────────────────────────────────────────────
# LOAD & STATE
# ─────────────────────────────────────────────────────────────────
COLS_NEEDED = [
    "Sprint", "Início Sprint", "Fim Sprint", "Interações", "Cod Interação",
    "Iniciativa", "Responsável iniciativa Ultragaz", "Responsável iniciativa ILOS",
    "Atividade", "Início planejado", "Fim planejado",
    "Início planejado UG", "Fim planejado UG",
    "Início efetivo", "Fim efetivo", "Responsável", "Ferramenta", "Piloto", "Piloto ILOS",
    "Iniciativa resumo", "Piloto/roll out",
]

@st.cache_data
def load_data(path):
    # Try "Gestão integrada" first (skiprows=2), fall back to "Data"
    try:
        df = pd.read_excel(path, sheet_name="Gestão integrada", skiprows=2)
    except Exception:
        df = pd.read_excel(path, sheet_name="Data")

    # Keep only the columns we need (ignore daily-date columns and extras)
    existing = [c for c in COLS_NEEDED if c in df.columns]
    df = df[existing].copy()

    df["Início planejado"]    = pd.to_datetime(df["Início planejado"],    errors="coerce")
    df["Fim planejado"]       = pd.to_datetime(df["Fim planejado"],       errors="coerce")
    df["Início planejado UG"] = pd.to_datetime(df.get("Início planejado UG"), errors="coerce")
    df["Fim planejado UG"]    = pd.to_datetime(df.get("Fim planejado UG"),    errors="coerce")
    df = df.dropna(subset=["Início planejado","Fim planejado","Iniciativa","Atividade"])
    for col in ["Responsável iniciativa Ultragaz","Responsável iniciativa ILOS",
                "Interações","Cod Interação","Responsável","Ferramenta","Piloto","Piloto ILOS","Piloto/roll out"]:
        if col in df.columns:
            df[col] = df[col].fillna("-").astype(str).str.strip()
        else:
            df[col] = "-"
    df["Sprint"] = df["Sprint"].fillna(1).astype(int).astype(str)
    # Use "Iniciativa resumo" if available, else fall back to "Iniciativa"
    if "Iniciativa resumo" in df.columns:
        df["Iniciativa resumo"] = df["Iniciativa resumo"].fillna(df["Iniciativa"]).astype(str).str.strip()
    else:
        df["Iniciativa resumo"] = df["Iniciativa"]
    return df

DEFAULT_XLSX = "Cronograma Sprint 1_v3.xlsx"

def load_marcos(path):
    """Load marcos from 'Marcos' or 'Marco' sheet if it exists."""
    try:
        xl = pd.ExcelFile(path)
        sheet = next((s for s in xl.sheet_names
                      if s.strip().lower() in ("marco", "marcos")), None)
        if sheet is None:
            return []
        df_m = pd.read_excel(path, sheet_name=sheet)
        df_m.columns = [c.strip() for c in df_m.columns]
        col_desc = next((c for c in df_m.columns if any(k in c.lower() for k in ("descri","nome","name"))), None)
        col_ini  = next((c for c in df_m.columns if any(k in c.lower() for k in ("inicio","início","start"))), None)
        col_fim  = next((c for c in df_m.columns if any(k in c.lower() for k in ("fim","end"))), None)
        col_cor  = next((c for c in df_m.columns if any(k in c.lower() for k in ("cor","color","colour"))), None)
        if not col_desc or not col_ini:
            return []
        COR_MAP = {
            "roxo":"#7C3AED","azul":"#2563EB","verde":"#059669",
            "laranja":"#F97316","preto":"#1C1F2E","vermelho":"#DC2626",
            "rosa":"#EC4899","cinza":"#64748B",
        }
        marcos = []
        for _, row in df_m.iterrows():
            nome = str(row[col_desc]).strip() if pd.notna(row[col_desc]) else ""
            if not nome or nome == "nan":
                continue
            try:
                data_ini = pd.Timestamp(row[col_ini])
            except Exception:
                continue
            data_fim = None
            if col_fim and pd.notna(row.get(col_fim)):
                try: data_fim = pd.Timestamp(row[col_fim])
                except Exception: pass
            cor = "#F97316"
            if col_cor and pd.notna(row.get(col_cor)):
                cor_raw = str(row[col_cor]).strip().lower()
                cor = COR_MAP.get(cor_raw, cor_raw if cor_raw.startswith("#") else "#F97316")
            marcos.append({"nome": nome, "data": data_ini, "data_fim": data_fim, "cor": cor})
        return marcos
    except Exception as e:
        return []

if "df"       not in st.session_state: st.session_state.df       = load_data(DEFAULT_XLSX)
if "df_orig"  not in st.session_state: st.session_state.df_orig  = st.session_state.df.copy()
if "expanded_inis" not in st.session_state: st.session_state.expanded_inis = set()
if "marcos_excel"  not in st.session_state: st.session_state.marcos_excel  = load_marcos(DEFAULT_XLSX)
if "marcos"        not in st.session_state: st.session_state.marcos        = []  # manual


def get_df():
    return st.session_state.df.copy()

# ─────────────────────────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────────────────────────
st.markdown("""
<div class="app-header">
  <div class="app-title">Cronograma de Implementação · Gestão Integrada</div>
  <div class="app-sub">Sprint 1–4 · Visualização hierárquica por iniciativa e atividade</div>
</div>
""", unsafe_allow_html=True)

with st.expander("📂 Carregar outro arquivo Excel"):
    uploaded = st.file_uploader("Planilha", type=["xlsx"], label_visibility="collapsed")
    if uploaded:
        import tempfile
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(uploaded.read())
        st.session_state.df           = load_data(tmp.name)
        st.session_state.df_orig      = st.session_state.df.copy()
        st.session_state.marcos_excel = load_marcos(tmp.name)
        st.rerun()

df_base = get_df()

# ─────────────────────────────────────────────────────────────────
# FILTER CARD
# ─────────────────────────────────────────────────────────────────
st.markdown('<div class="filter-card">', unsafe_allow_html=True)
st.markdown('<span class="filter-badge">🔍 Filtros</span>', unsafe_allow_html=True)

fc1, fc2, fc3, fc4, fc5, fc6, fc7 = st.columns([2, 2, 2, 1.2, 1.2, 2, 2])
with fc1:
    ug_opts  = sorted([x for x in df_base["Responsável iniciativa Ultragaz"].unique() if x != "-"])
    resp_ug  = st.multiselect("Responsável Ultragaz", ug_opts,  default=[], key="f_ug")
with fc2:
    il_opts  = sorted([x for x in df_base["Responsável iniciativa ILOS"].unique() if x != "-"])
    resp_ilos= st.multiselect("Responsável ILOS",    il_opts,  default=[], key="f_il")
with fc3:
    int_opts = sorted(df_base["Interações"].unique())
    interacao= st.multiselect("Interação",           int_opts, default=[], key="f_int")
with fc4:
    sp_opts  = sorted(df_base["Sprint"].unique())
    sprint   = st.multiselect("Sprint",              sp_opts,  default=[], key="f_sp")
with fc5:
    gr_opts  = sorted(df_base["Cod Interação"].unique())
    grupo    = st.multiselect("Grupo",               gr_opts,  default=[], key="f_gr")
with fc6:
    ini_opts = sorted(df_base["Iniciativa"].unique())
    ini_fil  = st.multiselect("Iniciativa",          ini_opts, default=[], key="f_ini")
with fc7:
    resp_atv_opts = sorted([x for x in df_base["Responsável"].unique() if x != "-"])
    resp_atv      = st.multiselect("Responsável atividade", resp_atv_opts, default=[], key="f_resp")

# Second filter row for Ferramenta + Piloto
ff1, ff2, ff3, ff4 = st.columns([2, 2, 2, 6])
with ff1:
    ferr_opts = sorted([x for x in df_base["Ferramenta"].unique() if x != "-"])
    ferr_fil  = st.multiselect("Ferramenta", ferr_opts, default=[], key="f_ferr")
with ff2:
    piloto_opts = sorted([x for x in df_base["Piloto"].unique() if x != "-"])
    piloto_fil  = st.multiselect("Piloto", piloto_opts, default=[], key="f_piloto")
with ff3:
    piloto_ilos_opts = sorted([x for x in df_base["Piloto ILOS"].unique() if x != "-"])
    piloto_ilos_fil  = st.multiselect("Piloto ILOS", piloto_ilos_opts, default=[], key="f_piloto_ilos")

st.markdown('</div>', unsafe_allow_html=True)

# Apply filters
df_f = df_base.copy()
if resp_ug:   df_f = df_f[df_f["Responsável iniciativa Ultragaz"].isin(resp_ug)]
if resp_ilos: df_f = df_f[df_f["Responsável iniciativa ILOS"].isin(resp_ilos)]
if interacao: df_f = df_f[df_f["Interações"].isin(interacao)]
if sprint:    df_f = df_f[df_f["Sprint"].isin(sprint)]
if grupo:     df_f = df_f[df_f["Cod Interação"].isin(grupo)]
if ini_fil:   df_f = df_f[df_f["Iniciativa"].isin(ini_fil)]
if resp_atv:  df_f = df_f[df_f["Responsável"].isin(resp_atv)]
if ferr_fil:        df_f = df_f[df_f["Ferramenta"].isin(ferr_fil)]
if piloto_fil:      df_f = df_f[df_f["Piloto"].isin(piloto_fil)]
if piloto_ilos_fil: df_f = df_f[df_f["Piloto ILOS"].isin(piloto_ilos_fil)]

if df_f.empty:
    st.warning("Nenhuma atividade encontrada com os filtros selecionados.")
    st.stop()

# ─────────────────────────────────────────────────────────────────
# METRICS
# ─────────────────────────────────────────────────────────────────
date_min = df_f["Início planejado"].min()
date_max = df_f["Fim planejado"].max()
semanas  = max(1, (date_max - date_min).days // 7)

m1, m2, m3, m4 = st.columns(4)
for col, num, lbl, sub, color in [
    (m1, df_f["Iniciativa"].nunique(), "Iniciativas",  f"{len(df_f)} atividades", "#4F46E5"),
    (m2, len(df_f),                    "Atividades",    f"de {date_min.strftime('%b/%y')} a {date_max.strftime('%b/%y')}", "#0891B2"),
    (m3, semanas,                      "Semanas",       f"{(date_max-date_min).days} dias total", "#27AE60"),
    (m4, df_f["Sprint"].nunique(),     "Sprints",       "no período filtrado", "#E67E22"),
]:
    with col:
        st.markdown(f"""
        <div class="mcard">
          <div class="mcard-num" style="color:{color}">{num}</div>
          <div class="mcard-lbl">{lbl}</div>
          <div class="mcard-sub">{sub}</div>
        </div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────────────────────────
tab_gantt, tab_mensal, tab_editor = st.tabs([
    "📊 Gantt Hierárquico",
    "📅 Atividades por Período",
    "✏️ Editar Datas",
])

# ═══════════════════════════════════════════════════════════════
# TAB 1 — GANTT
# ═══════════════════════════════════════════════════════════════
with tab_gantt:

    abs_min = df_f["Início planejado"].min().date()
    abs_max = df_f["Fim planejado"].max().date()

    ug_row1, ug_row2 = st.columns([2, 10])
    with ug_row1:
        show_ug_toggle = st.checkbox("Mostrar planejado UG", value=True, key="show_ug_bars")
    with ug_row2:
        # ── Marcos (vertical lines) ──
        with st.expander("📍 Marcos personalizados", expanded=False):
            mc_cols = st.columns([3, 2, 1.5, 1])
            with mc_cols[0]: mc_nome = st.text_input("Nome do marco", key="mc_nome", placeholder="Ex: Entrega fase 1")
            with mc_cols[1]: mc_data = st.date_input("Data", key="mc_data", value=datetime.today().date())
            with mc_cols[2]:
                mc_cor = st.selectbox("Cor", ["🟣 Roxo","🔵 Azul","🟢 Verde","🟠 Laranja","⚫ Preto"], key="mc_cor")
                COR_MAP = {"🟣 Roxo":"#7C3AED","🔵 Azul":"#2563EB","🟢 Verde":"#059669","🟠 Laranja":"#F97316","⚫ Preto":"#1C1F2E"}
            with mc_cols[3]:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("＋ Adicionar", key="mc_add", use_container_width=True):
                    if mc_nome.strip():
                        st.session_state.marcos.append({
                            "nome": mc_nome.strip(),
                            "data": pd.Timestamp(mc_data),
                            "cor":  COR_MAP.get(mc_cor, "#7C3AED"),
                        })
                        st.rerun()

            if st.session_state.marcos_excel:
                st.markdown("**Marcos do Excel:**")
                for m in st.session_state.marcos_excel:
                    fim_str = f" → {m['data_fim'].strftime('%d/%m/%Y')}" if m.get("data_fim") else ""
                    st.markdown(
                        f"<span style='color:{m["cor"]};font-weight:700'>▏</span> "
                        f"{m['nome']} — {m['data'].strftime('%d/%m/%Y')}{fim_str}",
                        unsafe_allow_html=True
                    )
                st.divider()

            if st.session_state.marcos:
                st.markdown("**Marcos manuais:**")
                for i, m in enumerate(st.session_state.marcos):
                    rm1, rm2 = st.columns([8, 1])
                    with rm1:
                        st.markdown(
                            f"<span style='color:{m["cor"]};font-weight:700'>▏</span> "
                            f"**{m['nome']}** — {m['data'].strftime('%d/%m/%Y')}",
                            unsafe_allow_html=True
                        )
                    with rm2:
                        if st.button("✕", key=f"mc_rm_{i}", use_container_width=True):
                            st.session_state.marcos.pop(i)
                            st.rerun()

    go1, go2, go3, go4, go5, go6, go7, go8 = st.columns([2, 2, 2, 2, 2, 2, 1.5, 2])
    with go1:
        modo_cor = st.selectbox("Colorir por",
            ["Iniciativa","Responsável UG","Responsável ILOS","Sprint","Responsável (atividade)","Ferramenta","Piloto","Piloto ILOS","Piloto (combinado)","Piloto/roll out"], key="cor")
    with go2:
        modo_exp = st.selectbox("Atividades",
            ["Expandir individualmente","Mostrar todas","Só iniciativas","Só atividades"], key="exp")
    with go3:
        if modo_exp == "Expandir individualmente":
            col_exp_btns = st.columns([1,1])
            with col_exp_btns[0]:
                if st.button("＋ Expandir todas", key="exp_all", use_container_width=True):
                    st.session_state.expanded_inis = set(df_f["Iniciativa"].unique())
                    st.rerun()
            with col_exp_btns[1]:
                if st.button("− Colapsar todas", key="col_all", use_container_width=True):
                    st.session_state.expanded_inis = set()
                    st.rerun()
    with go4:
        filter_from = st.date_input("Início a partir de", value=abs_min,
                                    min_value=abs_min, max_value=abs_max, key="ff")
    with go5:
        view_start = st.date_input("Janela — início", value=abs_min,
                                   min_value=abs_min, max_value=abs_max, key="vs")
    with go6:
        default_end = min(abs_max, (date_min + timedelta(days=185)).date())
        view_end = st.date_input("Janela — fim", value=default_end,
                                 min_value=abs_min, max_value=abs_max, key="ve")
    with go7:
        tick_mode = st.selectbox("Escala X", ["Mensal","Semanal"], key="tick")
    with go8:
        sort_mode = st.selectbox("Ordenar iniciativas por",
            ["Data de início (↑ mais cedo)","Data de início (↓ mais tarde)",
             "Data de fim (↑ mais cedo)","Nome (A→Z)"], key="sort_ini")

    # Apply start-date filter
    df_g = df_f[df_f["Início planejado"].dt.date >= filter_from].copy()
    if df_g.empty:
        st.warning("Nenhuma atividade com início a partir desta data.")
        st.stop()
    if view_start >= view_end:
        st.warning("A janela de início deve ser anterior à data fim.")
        st.stop()

    # ── Color map ──
    ini_list      = sorted(df_g["Iniciativa"].unique())
    ini_color_map = {ini: INI_PALETTE[i % len(INI_PALETTE)] for i, ini in enumerate(ini_list)}

    def bar_color(ini_name, rug, ril, sp, resp="-", ferr="-", piloto_ilos="-", piloto="-", piloto_ro="-"):
        if modo_cor == "Iniciativa":              return ini_color_map.get(ini_name, "#9BA3BF")
        if modo_cor == "Responsável UG":          return RESP_UG_PALETTE.get(rug,   "#9BA3BF")
        if modo_cor == "Responsável ILOS":        return RESP_IL_PALETTE.get(ril,   "#9BA3BF")
        if modo_cor == "Responsável (atividade)": return RESP_ATY_PALETTE.get(resp, "#9BA3BF")
        if modo_cor == "Ferramenta":               return FERR_COLORS.get(ferr, "#9BA3BF")
        if modo_cor == "Piloto":                   return PILOTO_COLORS.get(piloto, "#9BA3BF")
        if modo_cor == "Piloto ILOS":              return PILOTO_ILOS_COLORS.get(piloto_ilos, "#9BA3BF")
        if modo_cor == "Piloto (combinado)":       return PILOTO_COMB_COLORS.get((piloto_ilos, piloto), "#9BA3BF")
        if modo_cor == "Piloto/roll out":           return PILOTO_ROLLOUT_COLORS.get(piloto_ro, "#9BA3BF")
        return SP_COLORS.get(str(sp), "#9BA3BF")

    # ── Aggregate & SORT iniciativas by start date ──
    ini_agg = (
        df_g.groupby("Iniciativa")
        .agg(inicio=("Início planejado","min"), fim=("Fim planejado","max"),
             sprint=("Sprint","first"),
             resp_ug=("Responsável iniciativa Ultragaz","first"),
             resp_ilos=("Responsável iniciativa ILOS","first"),
             inicio_ug=("Início planejado UG","min"),
             fim_ug=("Fim planejado UG","max"),
             ferramenta=("Ferramenta","first"),
             piloto_ilos=("Piloto ILOS","first"),
             piloto=("Piloto","first"),
             resumo=("Iniciativa resumo","first"),
             piloto_ro=("Piloto/roll out","first"))
        .reset_index()
    )
    if sort_mode == "Data de início (↑ mais cedo)":
        ini_agg = ini_agg.sort_values("inicio", ascending=True)
    elif sort_mode == "Data de início (↓ mais tarde)":
        ini_agg = ini_agg.sort_values("inicio", ascending=False)
    elif sort_mode == "Data de fim (↑ mais cedo)":
        ini_agg = ini_agg.sort_values("fim", ascending=True)
    else:
        ini_agg = ini_agg.sort_values("Iniciativa", ascending=True)

    # ── Build bar rows ──
    bar_rows     = []
    y_order      = []
    legend_items = {}

    # For palette-based modes, pre-build legend from all unique values in the data
    _palette_modes = {
        "Ferramenta":      (FERR_COLORS,          "Ferramenta"),
        "Piloto":          (PILOTO_COLORS,         "Piloto"),
        "Piloto ILOS":     (PILOTO_ILOS_COLORS,    "Piloto ILOS"),
        "Piloto/roll out": (PILOTO_ROLLOUT_COLORS, "Piloto/roll out"),
        "Sprint":          (SP_COLORS,             "Sprint"),
    }
    if modo_cor in _palette_modes:
        palette, col = _palette_modes[modo_cor]
        if col in df_g.columns:
            for val in sorted(df_g[col].fillna("-").unique()):
                if val != "-":
                    legend_items[val] = palette.get(val, "#9BA3BF")
        legend_items["-"] = "#9BA3BF"   # always show "sem valor" last

    for ini_idx, (_, ir) in enumerate(ini_agg.iterrows()):
        ini_name = ir["Iniciativa"]
        atv_df   = df_g[df_g["Iniciativa"] == ini_name].sort_values("Início planejado")
        show_atv = (
            modo_exp == "Mostrar todas" or
            modo_exp == "Só atividades" or
            (modo_exp == "Expandir individualmente" and ini_name in st.session_state.expanded_inis)
        )
        show_ini = modo_exp != "Só atividades"

        color     = bar_color(ini_name, ir["resp_ug"], ir["resp_ilos"], ir["sprint"], ir["resp_ug"], ir.get("ferramenta","-"), ir.get("piloto_ilos","-"), ir.get("piloto","-"), ir.get("piloto_ro","-"))
        ini_resumo = ir.get("resumo", ini_name) if "resumo" in ir.index else ini_name
        ini_resumo_short = (ini_resumo[:56] + "…") if len(ini_resumo) > 58 else ini_resumo
        ini_short  = "​" * ini_idx + ini_resumo_short
        if modo_exp == "Expandir individualmente":
            icon = "▼ " if ini_name in st.session_state.expanded_inis else "▶ "
            ini_short = icon + ini_short

        # Legend key
        if   modo_cor == "Iniciativa":              lk, lc = ini_short,                color
        elif modo_cor == "Responsável UG":          lk, lc = ir["resp_ug"],            color
        elif modo_cor == "Responsável ILOS":        lk, lc = ir["resp_ilos"],          color
        elif modo_cor == "Responsável (atividade)":
            resps_in_ini = atv_df["Responsável"].unique()
            for rv in resps_in_ini:
                rc = RESP_ATY_PALETTE.get(rv, "#9BA3BF")
                if rv not in legend_items: legend_items[rv] = rc
            lk, lc = resps_in_ini[0] if len(resps_in_ini) else "-", color
        elif modo_cor == "Ferramenta":              lk, lc = ir.get("ferramenta","-"),        color
        elif modo_cor == "Piloto":                  lk, lc = ir.get("piloto","-"),            color
        elif modo_cor == "Piloto ILOS":             lk, lc = ir.get("piloto_ilos","-"),       color
        elif modo_cor == "Piloto (combinado)":
            comb_key = (ir.get("piloto_ilos","-"), ir.get("piloto","-"))
            lk, lc   = PILOTO_COMB_LABELS.get(comb_key, "Outros"), color
        elif modo_cor == "Piloto/roll out":         lk, lc = ir.get("piloto_ro","-"),          color
        else:                                       lk, lc = f"Sprint {ir['sprint']}", color
        if modo_cor != "Responsável (atividade)" and lk not in legend_items:
            legend_items[lk] = lc

        if show_ini:

            y_order.append(ini_short)
            bar_rows.append(dict(
                y=ini_short, inicio=ir["inicio"], fim=ir["fim"], color=color,
                ug_inicio=ir.get("inicio_ug"), ug_fim=ir.get("fim_ug"),
                bar_width=0.70, opacity=1.0, is_ini=True,
                border_color="#1C1F2E", border_width=1.5,
                resp_ug=ir["resp_ug"], resp_ilos=ir["resp_ilos"], sprint=ir["sprint"],
                hover_name=ini_name,
                ini_name=ini_name,
                legendgroup=lk,
            ))

        if show_atv:
            for _, ar in atv_df.iterrows():
                raw   = ar["Atividade"]
                short = (raw[:50] + "…") if len(raw) > 52 else raw
                # zero-width space prefix per initiative = unique Y, display unchanged
                label = "​" * ini_idx + "  ⤷ " + short
                y_order.append(label)
                atv_color = bar_color(ini_name, ar["Responsável iniciativa Ultragaz"],
                                      ar["Responsável iniciativa ILOS"], ar["Sprint"],
                                      ar["Responsável"], ar.get("Ferramenta","-"), ar.get("Piloto ILOS","-"), ar.get("Piloto","-"), ar.get("Piloto/roll out","-"))
                bar_rows.append(dict(
                    y=label,
                    inicio=ar["Início planejado"], fim=ar["Fim planejado"],
                    ug_inicio=ar.get("Início planejado UG"),
                    ug_fim=ar.get("Fim planejado UG"),
                    color=atv_color, bar_width=0.34, opacity=0.42, is_ini=False,
                    border_color=atv_color, border_width=0,
                    resp_ug=ar["Responsável iniciativa Ultragaz"],
                    resp_ilos=ar["Responsável iniciativa ILOS"],
                    sprint=ar["Sprint"],
                    hover_name=ar["Atividade"],
                    ini_name=ini_name,
                    legendgroup=lk,
                ))

    # ── Figure ──
    today  = datetime.today()
    height = max(480, 64 + len(bar_rows) * 28)
    fig    = go.Figure()

    for r in bar_rows:
        real_ini = r["inicio"]
        real_fim = r["fim"]
        vis_ini  = max(real_ini, pd.Timestamp(view_start))
        vis_fim  = min(real_fim, pd.Timestamp(view_end))

        if vis_fim <= vis_ini:
            fig.add_trace(go.Bar(
                name="", x=[0], y=[r["y"]],
                base=[pd.Timestamp(view_start).timestamp()*1000],
                orientation="h", marker_color="rgba(0,0,0,0)",
                legendgroup=r.get("legendgroup", r["color"]),
                width=r["bar_width"], showlegend=False, hoverinfo="skip",
            ))
            continue

        days   = (real_fim - real_ini).days
        dur_ms = (vis_fim - vis_ini).total_seconds() * 1000

        # Build UG date line for hover (if available)
        ug_i = r.get("ug_inicio")
        ug_f = r.get("ug_fim")
        if ug_i is not None and not pd.isna(ug_i) and ug_f is not None and not pd.isna(ug_f):
            ug_days  = (pd.Timestamp(ug_f) - pd.Timestamp(ug_i)).days
            ug_diff  = (real_ini - pd.Timestamp(ug_i)).days
            diff_str = f"+{ug_diff}d" if ug_diff > 0 else (f"{ug_diff}d" if ug_diff < 0 else "=")
            ug_line  = (
                f"<span style='color:#aaa'>── Planejado UG ──</span><br>"
                f"Início UG: {pd.Timestamp(ug_i).strftime('%d/%m/%Y')} · "
                f"Fim UG: {pd.Timestamp(ug_f).strftime('%d/%m/%Y')}<br>"
                f"Duração UG: {ug_days} dias  <b>({diff_str} vs planejado)</b><br>"
            )
        else:
            ug_line = ""

        if r["is_ini"]:
            hover = (
                f"<b>📌 {r['hover_name']}</b><br>"
                f"<span style='color:#aaa'>Iniciativa</span><br>"
                f"Início: {real_ini.strftime('%d/%m/%Y')} · Fim: {real_fim.strftime('%d/%m/%Y')}<br>"
                f"Duração total: {days} dias<br>"
                f"{ug_line}"
                f"UG: {r['resp_ug']}  ·  ILOS: {r['resp_ilos']}<br>"
                f"Sprint: {r['sprint']}<extra></extra>"
            )
        else:
            full    = r["hover_name"]
            wrapped = "<br>".join([full[i:i+64] for i in range(0, len(full), 64)])
            hover = (
                f"<b>⤷ {wrapped}</b><br>"
                f"<span style='color:#aaa'>Atividade</span><br>"
                f"Iniciativa: {r.get('ini_name','')}<br>"
                f"Início: {real_ini.strftime('%d/%m/%Y')} · Fim: {real_fim.strftime('%d/%m/%Y')}<br>"
                f"Duração: {days} dias<br>"
                f"{ug_line}"
                f"UG: {r['resp_ug']}  ·  ILOS: {r['resp_ilos']}<br>"
                f"Sprint: {r['sprint']}<extra></extra>"
            )

        fig.add_trace(go.Bar(
            name="", showlegend=False,
            legendgroup=r.get("legendgroup", r["color"]),
            legendgrouptitle=dict(text=""),
            x=[dur_ms], y=[r["y"]],
            base=[vis_ini.timestamp() * 1000],
            orientation="h",
            marker=dict(color=r["color"], opacity=r["opacity"],
                        line=dict(color=r["border_color"], width=r["border_width"])),
            width=r["bar_width"],
            hovertemplate=hover,
        ))





# ── UG planned comparison — using go.Bar with Pattern ──
    if show_ug_toggle:
        for r in bar_rows:
            ug_i = r.get("ug_inicio")
            ug_f = r.get("ug_fim")
            
            if not ug_i or not ug_f or pd.isna(ug_i) or pd.isna(ug_f):
                continue
                
            ug_vis_ini = max(pd.Timestamp(ug_i), pd.Timestamp(view_start))
            ug_vis_fim = min(pd.Timestamp(ug_f), pd.Timestamp(view_end))
            
            if ug_vis_fim <= ug_vis_ini:
                continue

            dur_ms_ug = (ug_vis_fim - ug_vis_ini).total_seconds() * 1000

            fig.add_trace(go.Bar(
                name="⬚ Planejado UG",
                showlegend=False,
                legendgroup="__ug__",
                x=[dur_ms_ug],
                y=[r["y"]],
                base=[ug_vis_ini.timestamp() * 1000],
                orientation="h",
                marker=dict(
                    color="rgba(0,0,0,0)", 
                    line=dict(color="#1C1F2E", width=1.5), # Borda sólida
                    pattern=dict(shape="/", fillmode="replace", fgcolor="#1C1F2E") # Listras diagonais
                ),
                width=r["bar_width"] * 1.1,
                hoverinfo="skip"
            ))

        # Traço invisível da legenda
        fig.add_trace(go.Bar(
            name="⬚ Planejado UG",
            x=[None], y=[None],
            marker=dict(
                color="rgba(0,0,0,0)",
                line=dict(color="#1C1F2E", width=1.5),
                pattern=dict(shape="/", fillmode="replace", fgcolor="#1C1F2E")
            ),
            showlegend=True,
            legendgroup="__ug__",
        ))

    # Today + Sprint lines
    if pd.Timestamp(view_start) <= today <= pd.Timestamp(view_end):
        fig.add_vline(x=today.timestamp()*1000, line_dash="dot", line_color="#EF4444",
                      line_width=2, annotation_text=f"Hoje {today.strftime('%d/%m')}",
                      annotation_font=dict(color="#EF4444", size=11),
                      annotation_position="top right")

    for sp_val in sorted(df_g["Sprint"].unique()):
        sp_end = df_g[df_g["Sprint"]==sp_val]["Fim planejado"].max()
        if pd.Timestamp(view_start) <= sp_end <= pd.Timestamp(view_end):
            fig.add_vline(x=sp_end.timestamp()*1000, line_dash="dash",
                          line_color="#CBD5E1", line_width=1,
                          annotation_text=f"  Sprint {sp_val}",
                          annotation_font=dict(size=10, color="#94A3B8"),
                          annotation_position="top left")

    # ── Custom marcos (Excel + manual) ──
    all_marcos = list(st.session_state.marcos_excel) + list(st.session_state.marcos)
    v_start = pd.Timestamp(view_start)
    v_end   = pd.Timestamp(view_end)
    for m in all_marcos:
        m_ini = m["data"]
        m_fim = m.get("data_fim")
        # Shaded region if fim provided
        if m_fim and m_fim > m_ini:
            shade_ini = max(m_ini, v_start)
            shade_fim = min(m_fim, v_end)
            if shade_fim > shade_ini:
                fig.add_vrect(
                    x0=shade_ini.timestamp()*1000, x1=shade_fim.timestamp()*1000,
                    fillcolor=m["cor"], opacity=0.08, line_width=0, layer="below",
                )
        if v_start <= m_ini <= v_end:
            fig.add_vline(
                x=m_ini.timestamp()*1000,
                line_dash="dashdot", line_color=m["cor"], line_width=2,
                annotation_text=f"  📍 {m['nome']}",
                annotation_font=dict(color=m["cor"], size=11, family="IBM Plex Sans, sans-serif"),
                annotation_position="top left",
            )
        if m_fim and v_start <= m_fim <= v_end:
            fig.add_vline(
                x=m_fim.timestamp()*1000,
                line_dash="dot", line_color=m["cor"], line_width=1.5,
                annotation_text=f"  ⏹ {m['nome']}",
                annotation_font=dict(color=m["cor"], size=10, family="IBM Plex Sans, sans-serif"),
                annotation_position="top left",
            )

    dtick_val  = "W1"    if tick_mode == "Semanal" else "M1"
    tick_fmt   = "%d/%m" if tick_mode == "Semanal" else "%b/%y"
    tick_angle = 45      if tick_mode == "Semanal" else 0

    fig.update_xaxes(
        type="date",
        range=[pd.Timestamp(view_start).timestamp()*1000,
               pd.Timestamp(view_end).timestamp()*1000],
        dtick=dtick_val, tickformat=tick_fmt, tickangle=tick_angle,
        showgrid=True, gridcolor="#E8ECF4", gridwidth=1,
        tickfont=dict(size=10, color="#6B7394"), fixedrange=True,
        side="top",
    )

    # Bold iniciativa labels
    ini_set    = {r["y"] for r in bar_rows if r["is_ini"]}
    tick_texts = [f"<b>{lbl}</b>" if lbl in ini_set else lbl for lbl in y_order]

    fig.update_yaxes(
        autorange="reversed", title="",
        tickfont=dict(size=11, color="#1C1F2E", family="IBM Plex Sans, sans-serif"),
        showgrid=True, gridcolor="#F2F4FA", gridwidth=1,
        tickmode="array", tickvals=y_order, ticktext=tick_texts,
        ticklabeloverflow="allow",
    )
    # Space allocation (top of figure, top to bottom):
    #   [top margin] → legend → x-axis labels (angled if weekly) → plot bars
    n_items     = len(legend_items)
    legend_rows = max(1, -(-n_items // 5))   # ~5 pills per row
    legend_px   = legend_rows * 24 + 6
    # Weekly ticks are angled 45° → need more vertical room than monthly
    xaxis_px    = 52 if tick_mode == "Semanal" else 28
    top_margin  = legend_px + xaxis_px + 12

    # legend_y > 1.0 places it in the margin space above the plot+xaxis
    # We push it high enough to clear the angled tick labels
    legend_y    = 1.0 + (xaxis_px + 8) / height

    fig.update_layout(
        barmode="overlay", height=height,
        margin=dict(l=8, r=20, t=top_margin, b=8),
        paper_bgcolor="white", plot_bgcolor="#FAFBFF",
        font=dict(family="IBM Plex Sans, sans-serif"),
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="bottom", y=legend_y,
            xanchor="left",   x=0,
            font=dict(size=11, color="#1C1F2E"),
            bgcolor="rgba(255,255,255,0.92)",
            bordercolor="#D8DCE8", borderwidth=1,
            tracegroupgap=0,
        ),
    )

    # ── Add invisible legend traces so Plotly renders a proper legend ──
    for lbl, clr in legend_items.items():
        short = (lbl[:42] + "…") if len(lbl) > 44 else lbl
        fig.add_trace(go.Bar(
            name=short,
            x=[None], y=[None],
            orientation="h",
            marker_color=clr,
            showlegend=True,
            legendgroup=lbl,
        ))

    st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CFG)

    # ── Toggle buttons per iniciativa (only in individual expand mode) ──
    if modo_exp == "Expandir individualmente":
        st.markdown('<div style="font-size:11px;color:#6B7394;text-transform:uppercase;letter-spacing:.08em;font-weight:700;margin:8px 0 6px;border-left:3px solid #4F46E5;padding-left:10px">Clique para expandir / colapsar iniciativas</div>', unsafe_allow_html=True)
        # Render in rows of 3
        ini_names_sorted = ini_agg["Iniciativa"].tolist()
        rows_of_3 = [ini_names_sorted[i:i+3] for i in range(0, len(ini_names_sorted), 3)]
        for row_items in rows_of_3:
            btn_cols = st.columns(3)
            for col_btn, ini_n in zip(btn_cols, row_items):
                with col_btn:
                    is_open  = ini_n in st.session_state.expanded_inis
                    icon     = "▼" if is_open else "▶"
                    short_n  = (ini_n[:38] + "…") if len(ini_n) > 40 else ini_n
                    btn_type = "primary" if is_open else "secondary"
                    if st.button(f"{icon} {short_n}", key=f"tog_{ini_n}",
                                 use_container_width=True, type=btn_type):
                        if is_open:
                            st.session_state.expanded_inis.discard(ini_n)
                        else:
                            st.session_state.expanded_inis.add(ini_n)
                        st.rerun()



# ═══════════════════════════════════════════════════════════════
# TAB 2 — EDITOR COMPLETO
# ═══════════════════════════════════════════════════════════════
with tab_editor:
    st.caption(
        "Edite qualquer campo diretamente na tabela e clique em **Aplicar** para atualizar o Gantt. "
        "Use **Exportar Excel** para salvar o arquivo com todas as alterações."
    )

    # Build editor table from the full (unfiltered by view) df
    df_edit_src = st.session_state.df.copy()

    # Only show rows matching active filters so the table isn't overwhelming
    df_edit_src_f = df_f.copy()  # already filtered

    SPRINT_OPTS   = sorted(df_edit_src["Sprint"].unique().tolist())
    INTERAC_OPTS  = sorted(df_edit_src["Interações"].unique().tolist())
    GRUPO_OPTS    = sorted(df_edit_src["Cod Interação"].unique().tolist())
    UG_OPTS       = sorted(df_edit_src["Responsável iniciativa Ultragaz"].unique().tolist())
    ILOS_OPTS     = sorted(df_edit_src["Responsável iniciativa ILOS"].unique().tolist())

    editor_rows = []
    editor_idx  = []   # store original df index for apply
    for orig_idx, row in df_edit_src_f.iterrows():
        editor_rows.append({
            "Iniciativa":  row["Iniciativa"],
            "Atividade":   row["Atividade"],
            "Início":      row["Início planejado"].date(),
            "Fim":         row["Fim planejado"].date(),
            "Dias":        (row["Fim planejado"] - row["Início planejado"]).days,
            "Resp. UG":    row["Responsável iniciativa Ultragaz"],
            "Resp. ILOS":  row["Responsável iniciativa ILOS"],
            "Sprint":      row["Sprint"],
            "Interação":   row["Interações"],
            "Grupo":       row["Cod Interação"],
        })
        editor_idx.append(orig_idx)

    editor_df = pd.DataFrame(editor_rows)

    edited_df = st.data_editor(
        editor_df,
        use_container_width=True,
        num_rows="fixed",
        hide_index=True,
        column_config={
            "Iniciativa": st.column_config.TextColumn("Iniciativa", width="large"),
            "Atividade":  st.column_config.TextColumn("Atividade",  width="large"),
            "Início":     st.column_config.DateColumn("Início",     format="DD/MM/YYYY", width="small"),
            "Fim":        st.column_config.DateColumn("Fim",        format="DD/MM/YYYY", width="small"),
            "Dias":       st.column_config.NumberColumn("Dias",     width="small",  disabled=True),
            "Resp. UG":   st.column_config.SelectboxColumn("Resp. UG",   options=UG_OPTS,    width="small"),
            "Resp. ILOS": st.column_config.SelectboxColumn("Resp. ILOS", options=ILOS_OPTS,  width="small"),
            "Sprint":     st.column_config.SelectboxColumn("Sprint",     options=SPRINT_OPTS,width="small"),
            "Interação":  st.column_config.SelectboxColumn("Interação",  options=INTERAC_OPTS,width="medium"),
            "Grupo":      st.column_config.SelectboxColumn("Grupo",      options=GRUPO_OPTS, width="small"),
        },
        key="editor_tab",
    )

    e1, e2, e3 = st.columns([1.5, 1.5, 5])
    with e1:
        if st.button("✅ Aplicar alterações", type="primary", use_container_width=True):
            df_new = st.session_state.df.copy()
            for i, er in edited_df.iterrows():
                orig_idx = editor_idx[i]
                df_new.loc[orig_idx, "Iniciativa"]                        = er["Iniciativa"]
                df_new.loc[orig_idx, "Atividade"]                         = er["Atividade"]
                df_new.loc[orig_idx, "Início planejado"]                  = pd.to_datetime(er["Início"])
                df_new.loc[orig_idx, "Fim planejado"]                     = pd.to_datetime(er["Fim"])
                df_new.loc[orig_idx, "Responsável iniciativa Ultragaz"]   = er["Resp. UG"]
                df_new.loc[orig_idx, "Responsável iniciativa ILOS"]       = er["Resp. ILOS"]
                df_new.loc[orig_idx, "Sprint"]                            = er["Sprint"]
                df_new.loc[orig_idx, "Interações"]                        = er["Interação"]
                df_new.loc[orig_idx, "Cod Interação"]                     = er["Grupo"]
            st.session_state.df = df_new
            st.success("✅ Alterações aplicadas! O Gantt foi atualizado.")
            st.rerun()

    with e2:
        if st.button("📥 Exportar Excel", use_container_width=True):
            df_exp = get_df()
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as w:
                df_exp.to_excel(w, sheet_name="Data_Atualizado", index=False)
            buf.seek(0)
            st.download_button(
                "⬇️ Baixar .xlsx", data=buf,
                file_name=f"Cronograma_{datetime.today().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    with e3:
        if st.button("🔄 Resetar todas as alterações", use_container_width=True):
            st.session_state.df = st.session_state.df_orig.copy()
            st.rerun()

# ═══════════════════════════════════════════════════════════════
# TAB 2 — ATIVIDADES POR PERÍODO
# ═══════════════════════════════════════════════════════════════
with tab_mensal:

    st.caption("Selecione o tipo de período e um intervalo para ver as atividades previstas.")

    # ── Helpers: build period labels ──
    def build_periods(d_min, d_max, tipo):
        """Return list of (label, periodo_inicio, periodo_fim) for the given period type."""
        periods = []
        if tipo == "Semana":
            # ISO weeks
            cur = d_min - timedelta(days=d_min.weekday())  # Monday of first week
            while cur <= d_max:
                end = cur + timedelta(days=6)
                periods.append((f"Sem {cur.strftime('%d/%m')}–{end.strftime('%d/%m/%Y')}", cur, end))
                cur += timedelta(weeks=1)
        elif tipo == "Mês":
            cur = d_min.replace(day=1)
            while cur <= d_max:
                end = (cur + pd.DateOffset(months=1) - timedelta(days=1)).to_pydatetime()
                periods.append((cur.strftime("%B/%Y").capitalize(), cur, end))
                cur = (cur + pd.DateOffset(months=1)).to_pydatetime()
        elif tipo == "Trimestre":
            cur = d_min.replace(day=1, month=((d_min.month - 1) // 3) * 3 + 1)
            while cur <= d_max:
                end = (cur + pd.DateOffset(months=3) - timedelta(days=1)).to_pydatetime()
                q   = (cur.month - 1) // 3 + 1
                periods.append((f"T{q}/{cur.year}", cur, end))
                cur = (cur + pd.DateOffset(months=3)).to_pydatetime()
        elif tipo == "Semestre":
            cur = d_min.replace(day=1, month=(1 if d_min.month <= 6 else 7))
            while cur <= d_max:
                end = (cur + pd.DateOffset(months=6) - timedelta(days=1)).to_pydatetime()
                s   = 1 if cur.month <= 6 else 2
                periods.append((f"S{s}/{cur.year}", cur, end))
                cur = (cur + pd.DateOffset(months=6)).to_pydatetime()
        elif tipo == "Ano":
            cur = d_min.replace(day=1, month=1)
            while cur <= d_max:
                end = cur.replace(month=12, day=31)
                periods.append((str(cur.year), cur, end))
                cur = cur.replace(year=cur.year + 1)
        return periods

    # ── Row 1: Tipo de período + Período + Equipe ──
    d_min_p = df_f["Início planejado"].min().to_pydatetime()
    d_max_p = df_f["Fim planejado"].max().to_pydatetime()

    pr1, pr2, pr3 = st.columns([2, 4, 2])
    with pr1:
        tipo_periodo = st.selectbox("Tipo de período",
            ["Semana","Mês","Trimestre","Semestre","Ano"], index=1, key="m_tipo")

    periods      = build_periods(d_min_p, d_max_p, tipo_periodo)
    period_labels = [p[0] for p in periods]
    today_p       = datetime.today()

    # Default: period containing today
    def_idx = 0
    for i, (_, pi, pf) in enumerate(periods):
        if pi <= today_p <= pf:
            def_idx = i
            break

    with pr2:
        per_sel   = st.selectbox("Período", period_labels, index=def_idx, key="m_per")
    _, per_inicio, per_fim = periods[period_labels.index(per_sel)]
    per_inicio = pd.Timestamp(per_inicio)
    per_fim    = pd.Timestamp(per_fim)

    with pr3:
        lado_m = st.radio("Equipe", ["Ultragaz (UG)", "ILOS"], horizontal=True, key="m_lado")

    resp_col_m = ("Responsável iniciativa Ultragaz" if "UG" in lado_m
                  else "Responsável iniciativa ILOS")

    # ── Row 2: Filtros de responsável + ferramenta ──
    mc3, mc4, mc5, _ = st.columns([2, 2, 2, 2])
    with mc3:
        pessoas_m = sorted([x for x in df_f[resp_col_m].unique() if x != "-"])
        pessoa_m  = st.selectbox("Responsável iniciativa", ["Todos"] + pessoas_m, key="m_pessoa")
    with mc4:
        resp_atv_opts_m = sorted([x for x in df_f["Responsável"].unique() if x != "-"])
        pessoa_atv_m    = st.selectbox("Responsável atividade", ["Todos"] + resp_atv_opts_m, key="m_resp_atv")
    with mc5:
        ferr_opts_m = sorted([x for x in df_f["Ferramenta"].unique() if x != "-"])
        ferr_m      = st.selectbox("Ferramenta", ["Todos"] + ferr_opts_m, key="m_ferr")

    # ── Filter: activities overlapping the period ──
    df_m = df_f[
        (df_f["Início planejado"] <= per_fim) &
        (df_f["Fim planejado"]    >= per_inicio)
    ].copy()
    if pessoa_m    != "Todos": df_m = df_m[df_m[resp_col_m]    == pessoa_m]
    if pessoa_atv_m != "Todos": df_m = df_m[df_m["Responsável"] == pessoa_atv_m]
    if ferr_m      != "Todos": df_m = df_m[df_m["Ferramenta"]   == ferr_m]

    lbl_periodo = tipo_periodo.lower()

    if df_m.empty:
        st.info(f"Nenhuma atividade prevista para este {lbl_periodo} / filtros selecionados.")
    else:
        # ── Summary metrics ──
        sm1, sm2, sm3 = st.columns(3)
        with sm1:
            st.markdown(f"""<div class="mcard">
              <div class="mcard-num" style="color:#4F46E5">{df_m["Iniciativa"].nunique()}</div>
              <div class="mcard-lbl">Iniciativas</div>
            </div>""", unsafe_allow_html=True)
        with sm2:
            st.markdown(f"""<div class="mcard">
              <div class="mcard-num" style="color:#0891B2">{len(df_m)}</div>
              <div class="mcard-lbl">Atividades no {lbl_periodo}</div>
            </div>""", unsafe_allow_html=True)
        with sm3:
            pessoas_ativas = df_m[resp_col_m].nunique()
            st.markdown(f"""<div class="mcard">
              <div class="mcard-num" style="color:#27AE60">{pessoas_ativas}</div>
              <div class="mcard-lbl">Responsáveis ativos</div>
            </div>""", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # ── X-axis tick config per period type ──
        TICK_CFG = {
            "Semana":    dict(dtick=86400000,     tickformat="%a %d/%m"),   # daily
            "Mês":       dict(dtick="W1",          tickformat="%d/%m"),      # weekly
            "Trimestre": dict(dtick="M1",          tickformat="%b/%y"),      # monthly
            "Semestre":  dict(dtick="M1",          tickformat="%b/%y"),
            "Ano":       dict(dtick="M3",          tickformat="%b/%y"),      # quarterly
        }
        tick_cfg = TICK_CFG.get(tipo_periodo, dict(dtick="W1", tickformat="%d/%m"))

        # ── Mini Gantt clipped to the period ──
        ini_list_m      = sorted(df_m["Iniciativa"].unique())
        ini_color_map_m = {ini: INI_PALETTE[i % len(INI_PALETTE)] for i, ini in enumerate(ini_list_m)}

        fig_m  = go.Figure()
        rows_m = []

        ini_agg_m = (
            df_m.groupby("Iniciativa")
            .agg(inicio=("Início planejado","min"), fim=("Fim planejado","max"),
                 resp_ug=("Responsável iniciativa Ultragaz","first"),
                 resp_ilos=("Responsável iniciativa ILOS","first"),
                 sprint=("Sprint","first"),
                 resumo=("Iniciativa resumo","first"))
            .reset_index()
            .sort_values("inicio")
        )

        y_m = []
        for m_idx, (_, ir) in enumerate(ini_agg_m.iterrows()):
            ini_n   = ir["Iniciativa"]
            ini_res = ir.get("resumo", ini_n) or ini_n
            color   = ini_color_map_m.get(ini_n, "#9BA3BF")
            # Unique Y label: resumo + zero-width-space prefix per index
            short   = "​" * m_idx + ((ini_res[:56] + "…") if len(ini_res) > 58 else ini_res)

            y_m.append(short)
            rows_m.append(dict(
                y=short, inicio=ir["inicio"], fim=ir["fim"],
                color=color, width=0.70, opacity=1.0, is_ini=True,
                hover=f"<b>📌 {ini_res}</b><br>UG: {ir['resp_ug']} · ILOS: {ir['resp_ilos']}<br>Sprint: {ir['sprint']}<extra></extra>",
            ))

            for a_idx, (_, ar) in enumerate(df_m[df_m["Iniciativa"] == ini_n].sort_values("Início planejado").iterrows()):
                raw   = ar["Atividade"]
                lbl   = "​" * m_idx + "  ⤷ " + ((raw[:50] + "…") if len(raw) > 52 else raw)
                days  = (ar["Fim planejado"] - ar["Início planejado"]).days
                resp  = ar[resp_col_m]
                y_m.append(lbl)
                rows_m.append(dict(
                    y=lbl, inicio=ar["Início planejado"], fim=ar["Fim planejado"],
                    color=color, width=0.36, opacity=0.45, is_ini=False,
                    hover=(
                        f"<b>⤷ {'<br>'.join([raw[i:i+64] for i in range(0,len(raw),64)])}</b><br>"
                        f"Iniciativa: {ini_res}<br>"
                        f"{resp_col_m.replace('Responsável iniciativa ','')}: {resp}<br>"
                        f"Início: {ar['Início planejado'].strftime('%d/%m/%Y')} · "
                        f"Fim: {ar['Fim planejado'].strftime('%d/%m/%Y')}<br>"
                        f"Duração: {days} dias<extra></extra>"
                    ),
                ))

        for r in rows_m:
            vis_ini = max(r["inicio"], per_inicio)
            vis_fim = min(r["fim"],    per_fim)
            if vis_fim <= vis_ini:
                continue
            dur_ms = (vis_fim - vis_ini).total_seconds() * 1000
            fig_m.add_trace(go.Bar(
                name="", showlegend=False,
                x=[dur_ms], y=[r["y"]],
                base=[vis_ini.timestamp() * 1000],
                orientation="h",
                marker=dict(
                    color=r["color"], opacity=r["opacity"],
                    line=dict(color="#1C1F2E" if r["is_ini"] else r["color"],
                              width=1.5 if r["is_ini"] else 0),
                ),
                width=r["width"],
                hovertemplate=r["hover"],
            ))

        if per_inicio <= pd.Timestamp(today_p) <= per_fim:
            fig_m.add_vline(
                x=today_p.timestamp()*1000,
                line_dash="dot", line_color="#EF4444", line_width=2,
                annotation_text=f"Hoje {today_p.strftime('%d/%m')}",
                annotation_font=dict(color="#EF4444", size=11),
            )

        ini_set_m  = {r["y"] for r in rows_m if r["is_ini"]}
        tick_txt_m = [f"<b>{l}</b>" if l in ini_set_m else l for l in y_m]

        height_m = max(360, 60 + len(rows_m) * 28)
        fig_m.update_xaxes(
            type="date",
            range=[per_inicio.timestamp()*1000, (per_fim + timedelta(days=1)).timestamp()*1000],
            dtick=tick_cfg["dtick"], tickformat=tick_cfg["tickformat"], tickangle=45,
            showgrid=True, gridcolor="#E8ECF4",
            tickfont=dict(size=10, color="#6B7394"),
            side="top", fixedrange=True,
        )
        fig_m.update_yaxes(
            autorange="reversed", title="",
            tickmode="array", tickvals=y_m, ticktext=tick_txt_m,
            tickfont=dict(size=11, color="#1C1F2E", family="IBM Plex Sans, sans-serif"),
            showgrid=True, gridcolor="#F2F4FA",
            ticklabeloverflow="allow",
        )
        fig_m.update_layout(
            barmode="overlay", height=height_m,
            margin=dict(l=8, r=20, t=60, b=8),
            paper_bgcolor="white", plot_bgcolor="#FAFBFF",
            font=dict(family="IBM Plex Sans, sans-serif"),
            showlegend=False,
        )

        st.plotly_chart(fig_m, use_container_width=True, config=PLOTLY_CFG)

        # ── Detail table ──
        st.markdown('<div style="font-size:11px;color:#6B7394;text-transform:uppercase;letter-spacing:.08em;font-weight:700;margin:16px 0 8px;border-left:3px solid #4F46E5;padding-left:10px">Detalhes das atividades</div>', unsafe_allow_html=True)

        table_rows = []
        for _, row in df_m.sort_values(["Iniciativa","Início planejado"]).iterrows():
            ini_dt = row["Início planejado"]
            fim_dt = row["Fim planejado"]
            starts = "✅" if ini_dt >= per_inicio else "◀ em andamento"
            ends   = "✅" if fim_dt <= per_fim    else "▶ continua"
            table_rows.append({
                "Iniciativa":    row.get("Iniciativa resumo") or row["Iniciativa"],
                "Atividade":     row["Atividade"],
                "Início":        ini_dt.strftime("%d/%m/%Y"),
                "Fim":           fim_dt.strftime("%d/%m/%Y"),
                "Dias totais":   (fim_dt - ini_dt).days,
                resp_col_m.replace("Responsável iniciativa ",""): row[resp_col_m],
                f"Inicia no {lbl_periodo}":  starts,
                f"Encerra no {lbl_periodo}": ends,
            })

        st.dataframe(pd.DataFrame(table_rows), use_container_width=True, hide_index=True)