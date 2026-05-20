# ============================================================
# AZIMUT PORTFOLIO ANALYZER — app.py
# Streamlit Web App — Production Ready for Streamlit Cloud
# ============================================================

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from openpyxl import load_workbook
import io
import datetime
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from reportlab.lib.pagesizes import A4
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer,
    Table, TableStyle, Image as RLImage, HRFlowable
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors as rl_colors
from reportlab.lib.units import cm

# ── PAGE CONFIG ─────────────────────────────────────────────
st.set_page_config(
    page_title="Azimut | Portfolio Analyzer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CONSTANTS ───────────────────────────────────────────────
GROUP_NAMES = {"ALLOCATION", "AZIONARI (LONG)", "BOND"}

# Column indices (0-based) in PTF sheets
COL_A, COL_B, COL_C, COL_G, COL_K, COL_O, COL_R = 0, 1, 2, 6, 10, 14, 17

PROFILES = ["CONSERVATIVO", "EQUILIBRATO", "ACCRESCITIVO", "CENTRALE"]
PROFILE_ICONS = {"CONSERVATIVO": "🛡️", "EQUILIBRATO": "⚖️", "ACCRESCITIVO": "📈", "CENTRALE": "🎯"}
PROFILE_W_COL = {
    "CONSERVATIVO": "w_cons",
    "EQUILIBRATO":  "w_equil",
    "ACCRESCITIVO": "w_accr",
    "CENTRALE":     "w_centrale",
}

MACRO_COLORS = {
    "Azionari":              "#1B4FBB",
    "Bilanciati/Flessibili": "#C9A84C",
    "Obbligazionari":        "#2D9D78",
    "Alternativi":           "#8B5CF6",
    "Monetario":             "#F59E0B",
    "Altro":                 "#94A3B8",
}

# Multiple shades per macro-category for funds
SHADES = {
    "Azionari":              ["#0D3080", "#1B4FBB", "#2563EB", "#3B82F6", "#60A5FA", "#93C5FD", "#BFDBFE"],
    "Bilanciati/Flessibili": ["#92650A", "#B8860B", "#C9A84C", "#D4B572", "#DFC298", "#E9CEB4", "#F3DACD"],
    "Obbligazionari":        ["#065F46", "#14855F", "#2D9D78", "#34B98A", "#6DE5BC", "#9AEFD2", "#C5F7E7"],
    "Alternativi":           ["#5B21B6", "#7C3AED", "#8B5CF6", "#A78BFA", "#C4B5FD", "#DDD6FE"],
    "Altro":                 ["#475569", "#64748B", "#94A3B8", "#CBD5E1", "#E2E8F0"],
}

# Default AZ% by macro-category for free portfolio
DEFAULT_AZ = {
    "Azionari": 0.92,
    "Bilanciati/Flessibili": 0.50,
    "Obbligazionari": 0.06,
    "Alternativi": 0.30,
    "Monetario": 0.02,
    "Altro": 0.50,
}

# ── CUSTOM CSS ──────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@600;700&family=DM+Sans:opsz,wght@9..40,300;9..40,400;9..40,500;9..40,600&display=swap');

html, body, [class*="css"] { font-family:'DM Sans',sans-serif; }
h1,h2,h3 { font-family:'Cormorant Garamond',serif !important; }

/* Sidebar */
[data-testid="stSidebar"] {
    background: linear-gradient(170deg,#06101e 0%,#0d1f3c 55%,#0a1628 100%);
    border-right: 1px solid #1a3050;
}
[data-testid="stSidebar"] .stFileUploader label,
[data-testid="stSidebar"] .stRadio  > label,
[data-testid="stSidebar"] .stSelectbox > label {
    color:#4a6582 !important;
    font-size:0.68rem !important;
    letter-spacing:0.12em !important;
    text-transform:uppercase !important;
    font-weight:600 !important;
}
[data-testid="stSidebar"] .stRadio [data-testid="stMarkdownContainer"] p {
    color:#c0cfe0 !important;
    font-size:0.9rem !important;
}
[data-testid="stSidebar"] .stSelectbox > div > div {
    background:#132035 !important;
    border:1px solid #243d5a !important;
    color:#dde6f0 !important;
    border-radius:6px !important;
}
[data-testid="stSidebar"] .stSelectbox svg { color:#5b7fa0 !important; }
[data-testid="stSidebar"] .stFileUploader > div {
    background:#132035 !important;
    border:1px dashed #2a4a6a !important;
    border-radius:8px !important;
}
[data-testid="stSidebar"] .stFileUploader p,
[data-testid="stSidebar"] .stFileUploader span { color:#8aa5c0 !important; font-size:0.8rem !important; }

/* Main */
.main { background:#f6f8fb !important; }
.block-container { padding-top:1.8rem !important; max-width:1300px; }

/* Header */
.az-header {
    background:linear-gradient(130deg,#081420 0%,#0f2644 50%,#162e52 100%);
    border-radius:16px; padding:2rem 2.5rem; margin-bottom:1.8rem;
    position:relative; overflow:hidden;
    box-shadow: 0 8px 32px rgba(0,0,0,0.15);
}
.az-header::after {
    content:''; position:absolute; bottom:-60px; right:-40px;
    width:220px; height:220px; border-radius:50%;
    background:radial-gradient(circle,rgba(201,168,76,.18) 0%,transparent 70%);
}
.az-eyebrow { font-size:.65rem; letter-spacing:.2em; color:#4a7098; text-transform:uppercase; font-weight:600; }
.az-title { font-family:'Cormorant Garamond',serif; font-size:2.1rem; font-weight:700; color:#f0f6ff; margin:.2rem 0 .4rem; line-height:1.1; }
.az-rule { width:38px; height:3px; background:#C9A84C; border-radius:2px; margin:.6rem 0; }
.az-meta { font-size:.88rem; color:#6b8fb0; }

/* KPI cards */
.kpi { background:#fff; border:1px solid #e4eaf3; border-radius:12px; padding:1.2rem 1.4rem; box-shadow:0 1px 4px rgba(0,0,0,.05); }
.kpi-label { font-size:.65rem; text-transform:uppercase; letter-spacing:.1em; color:#94a3b8; font-weight:500; margin-bottom:.3rem; }
.kpi-value { font-size:1.9rem; font-weight:700; color:#0d1b2a; font-family:'Cormorant Garamond',serif; line-height:1; }
.kpi-sub { font-size:.75rem; color:#64748b; margin-top:.3rem; }

/* Section title */
.sec-title {
    font-family:'Cormorant Garamond',serif; font-size:1.25rem; font-weight:600;
    color:#0d1b2a; border-bottom:2px solid #c9a84c;
    display:inline-block; padding-bottom:.4rem; margin-bottom:.9rem;
}

/* Fund list card */
.fund-group-hdr {
    background:#f0f4f9; padding:.45rem 1rem;
    font-size:.65rem; font-weight:700; letter-spacing:.12em; text-transform:uppercase;
    color:#64748b; border-bottom:1px solid #e2e8f0;
}
.fund-row {
    display:flex; align-items:center; gap:10px;
    padding:.65rem 1rem; border-bottom:1px solid #f1f5f9;
}
.fund-row:last-child { border-bottom:none; }
.fund-dot { width:8px; height:34px; border-radius:3px; flex-shrink:0; }
.fund-name { font-size:.83rem; color:#1e293b; font-weight:500; flex:1; min-width:0; }
.fund-cat  { font-size:.68rem; color:#64748b; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
.fund-pct  { font-size:1rem; font-weight:700; color:#0d1b2a; min-width:2.8rem; text-align:right; }

/* Download button */
[data-testid="stDownloadButton"] > button {
    background:linear-gradient(135deg,#0f2d6b 0%,#1b4fbb 100%) !important;
    color:#fff !important; font-size:1.05rem !important; font-weight:600 !important;
    padding:.9rem 2rem !important; border-radius:10px !important;
    border:none !important; width:100% !important; letter-spacing:.02em !important;
    box-shadow:0 4px 18px rgba(27,79,187,.35) !important;
    transition:all .25s ease !important;
}
[data-testid="stDownloadButton"] > button:hover {
    box-shadow:0 6px 24px rgba(27,79,187,.55) !important;
    transform:translateY(-2px) !important;
}

/* Weight alerts */
.w-ok   { background:#d1fae5; border:1px solid #6ee7b7; border-radius:8px; padding:.7rem 1rem; font-size:.84rem; color:#065f46; }
.w-warn { background:#fef3c7; border:1px solid #fcd34d; border-radius:8px; padding:.7rem 1rem; font-size:.84rem; color:#92400e; }

/* Expander clean */
[data-testid="stExpander"] { border:1px solid #e2e8f0 !important; border-radius:10px !important; }

/* Number input */
.stNumberInput input { text-align:right; }

div[data-testid="stVerticalBlock"] > div:has(> [data-testid="stDownloadButton"]) { margin-top:.5rem; }
</style>
""", unsafe_allow_html=True)


# ── HELPERS ─────────────────────────────────────────────────

def get_macro(cat: str) -> str:
    if not cat or cat == "-":
        return "Altro"
    c = cat.lower()
    if "azionari" in c or "equity" in c:
        return "Azionari"
    if any(x in c for x in ["obbligazionari", "bond", "credit", "debt", "sukuk", "reddito"]):
        return "Obbligazionari"
    if any(x in c for x in ["bilanciati", "allocation", "flessibili", "balanced", "flexible", "prudenti", "moderati"]):
        return "Bilanciati/Flessibili"
    if any(x in c for x in ["alternativi", "alternative", "commodity", "commodit"]):
        return "Alternativi"
    if any(x in c for x in ["monetari", "money market", "liquidit"]):
        return "Monetario"
    return "Altro"


def assign_colors(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    counter = {k: 0 for k in SHADES}
    colors = []
    for _, row in df.iterrows():
        mc = row.get("macro_cat", "Altro")
        shades = SHADES.get(mc, SHADES["Altro"])
        colors.append(shades[counter.get(mc, 0) % len(shades)])
        counter[mc] = counter.get(mc, 0) + 1
    df["color"] = colors
    return df


# ── DATA PARSING ────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def parse_excel(file_bytes: bytes) -> dict:
    wb = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=False)
    out = {}
    for sname in ["PTF FULL", "PTF SHORT"]:
        if sname in wb.sheetnames:
            out[sname] = _parse_ptf(wb, sname)
    if "FIDA" in wb.sheetnames:
        out["FIDA"] = _parse_fida(wb)
    wb.close()
    return out


def _parse_ptf(wb, sheet_name: str) -> pd.DataFrame:
    ws = wb[sheet_name]
    rows = list(ws.iter_rows(values_only=True))

    funds = []
    cur_group = {"name": None, "mc": 1.0, "me": 1.0, "ma": 1.0}

    for row in rows:
        name = row[COL_A]
        if not name or not isinstance(name, str):
            continue
        name = name.strip()

        if name in GROUP_NAMES:
            mc = row[COL_C] if isinstance(row[COL_C], (int, float)) else 1.0
            me = row[COL_G] if isinstance(row[COL_G], (int, float)) else 1.0
            ma = row[COL_K] if isinstance(row[COL_K], (int, float)) else 1.0
            cur_group = {"name": name, "mc": float(mc), "me": float(me), "ma": float(ma)}
            continue

        # Fund row — must start with "AZ" and have a valid R-weight
        if name.startswith("AZ") and cur_group["name"]:
            cat   = row[COL_B] if isinstance(row[COL_B], str) else ""
            az    = float(row[COL_O]) if isinstance(row[COL_O], (int, float)) else 0.5
            rw    = row[COL_R]
            if not isinstance(rw, (int, float)) or rw <= 0:
                continue
            funds.append({
                "nome":     name,
                "categoria": cat,
                "gruppo":   cur_group["name"],
                "az_pct":  min(1.0, max(0.0, az)),
                "obb_pct": min(1.0, max(0.0, 1.0 - az)),
                "r_weight": float(rw),
                "mc": cur_group["mc"],
                "me": cur_group["me"],
                "ma": cur_group["ma"],
            })

    if not funds:
        return pd.DataFrame()

    df = pd.DataFrame(funds)

    # Compute profile weights = group_multiplier × r_weight, then normalize
    for wcol, mcol in [("w_cons", "mc"), ("w_equil", "me"), ("w_accr", "ma")]:
        raw = df["r_weight"] * df[mcol]
        total = raw.sum()
        df[wcol] = raw / total if total > 0 else raw

    # CENTRALE: direct r_weight normalized
    tot_c = df["r_weight"].sum()
    df["w_centrale"] = df["r_weight"] / tot_c if tot_c > 0 else df["r_weight"]

    df["macro_cat"] = df["categoria"].apply(get_macro)
    df = assign_colors(df)
    return df


def _parse_fida(wb) -> pd.DataFrame:
    ws = wb["FIDA"]
    rows = list(ws.iter_rows(values_only=True))
    funds = []
    for row in rows[1:]:          # skip header
        nome = row[0]
        isin = row[1]
        cat  = row[2]
        if not nome or not isinstance(nome, str):
            continue
        nome = nome.strip().replace("\xa0", "")
        cat  = (cat or "").strip()
        funds.append({"nome": nome, "isin": isin or "", "categoria": cat, "macro_cat": get_macro(cat)})
    df = pd.DataFrame(funds).drop_duplicates(subset=["nome"])
    return df


# ── CHARTS ──────────────────────────────────────────────────

def make_fund_pie(df: pd.DataFrame, wcol: str, profile: str) -> go.Figure:
    d = df[df[wcol] > 0.005].copy()
    d["pct"] = d[wcol] * 100

    labels = d["nome"].apply(lambda x: (x[:38] + "…") if len(x) > 38 else x)

    fig = go.Figure(go.Pie(
        labels=labels,
        values=d["pct"],
        marker=dict(colors=d["color"].tolist(), line=dict(color="#fff", width=2.5)),
        hovertemplate="<b>%{label}</b><br>Peso: <b>%{value:.1f}%</b><extra></extra>",
        textinfo="percent",
        textfont=dict(size=10, family="DM Sans"),
        hole=0.40,
        pull=[0.04 if v == d["pct"].max() else 0 for v in d["pct"]],
        sort=False,
        direction="clockwise",
    ))
    fig.update_layout(
        margin=dict(t=10, b=10, l=10, r=180),
        showlegend=True,
        legend=dict(
            x=1.01, y=0.5, orientation="v",
            font=dict(size=9.5, family="DM Sans"),
            bgcolor="rgba(0,0,0,0)",
            itemclick=False, itemdoubleclick=False,
        ),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        height=430,
        annotations=[dict(
            text=f"<b>{profile[:4]}</b>",
            x=0.5, y=0.5, showarrow=False,
            font=dict(size=17, color="#0d1b2a", family="Cormorant Garamond"),
        )],
    )
    return fig


def make_macro_bar(df: pd.DataFrame, wcol: str) -> go.Figure:
    agg = (
        df[df[wcol] > 0.001]
        .groupby("macro_cat")[wcol].sum()
        .reset_index()
        .sort_values(wcol, ascending=True)
    )
    agg["pct"] = agg[wcol] * 100
    agg["color"] = agg["macro_cat"].map(MACRO_COLORS)

    fig = go.Figure(go.Bar(
        x=agg["pct"],
        y=agg["macro_cat"],
        orientation="h",
        marker=dict(color=agg["color"].tolist(), line=dict(color="#fff", width=1)),
        hovertemplate="<b>%{y}</b>: %{x:.1f}%<extra></extra>",
        text=agg["pct"].apply(lambda v: f"{v:.1f}%"),
        textposition="inside",
        insidetextanchor="middle",
        textfont=dict(color="#fff", size=11, family="DM Sans"),
    ))
    fig.update_layout(
        margin=dict(t=10, b=10, l=10, r=10),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(showgrid=False, showticklabels=False, range=[0, 105]),
        yaxis=dict(showgrid=False, tickfont=dict(size=11, family="DM Sans")),
        height=max(160, len(agg) * 48),
        bargap=0.28,
    )
    return fig


# ── PDF GENERATION ──────────────────────────────────────────

def _make_mpl_pie(df: pd.DataFrame, wcol: str, profile: str) -> io.BytesIO:
    """Matplotlib donut chart for PDF embedding."""
    d = df[df[wcol] > 0.005].sort_values(wcol, ascending=False)

    fig, (ax_pie, ax_leg) = plt.subplots(1, 2, figsize=(11, 6),
                                          gridspec_kw={"width_ratios": [1.3, 1]})

    wedges, _, autotexts = ax_pie.pie(
        d[wcol],
        colors=d["color"].tolist(),
        autopct=lambda p: f"{p:.1f}%" if p > 3.5 else "",
        pctdistance=0.72,
        wedgeprops=dict(width=0.58, edgecolor="white", linewidth=2),
        startangle=90,
    )
    for at in autotexts:
        at.set_fontsize(8); at.set_color("white"); at.set_fontweight("bold")
    ax_pie.text(0, 0, profile[:4], ha="center", va="center",
                fontsize=15, fontweight="bold", color="#0D1B2A")

    ax_leg.axis("off")
    handles = [
        mpatches.Patch(
            color=row["color"],
            label=f"{row['nome'][:34]}{'…' if len(row['nome'])>34 else ''}  {row[wcol]*100:.1f}%"
        )
        for _, row in d.iterrows()
    ]
    ax_leg.legend(handles=handles, loc="center left", frameon=False,
                  fontsize=7.8, labelspacing=0.85, handlelength=1.2)

    fig.patch.set_facecolor("#FFFFFF")
    plt.tight_layout(pad=1.5)

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    buf.seek(0)
    return buf


def generate_pdf(df: pd.DataFrame, wcol: str, profile: str, ptf_name: str) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            leftMargin=2*cm, rightMargin=2*cm,
                            topMargin=2.2*cm, bottomMargin=2.2*cm)

    ss = getSampleStyleSheet()
    S = lambda name, **kw: ParagraphStyle(name, parent=ss["Normal"], **kw)

    title_s   = S("T", fontName="Helvetica-Bold",   fontSize=22, textColor=rl_colors.HexColor("#0D1B2A"), spaceAfter=4, leading=28)
    eyebrow_s = S("E", fontName="Helvetica",         fontSize=8,  textColor=rl_colors.HexColor("#94A3B8"), spaceAfter=4, letterSpacing=1.5)
    sub_s     = S("U", fontName="Helvetica",         fontSize=10, textColor=rl_colors.HexColor("#64748B"), spaceAfter=4)
    sec_s     = S("H", fontName="Helvetica-Bold",    fontSize=11, textColor=rl_colors.HexColor("#0D1B2A"), spaceBefore=18, spaceAfter=8)
    body_s    = S("B", fontName="Helvetica",         fontSize=8.5, textColor=rl_colors.HexColor("#1E293B"), leading=13)
    foot_s    = S("F", fontName="Helvetica-Oblique", fontSize=7,  textColor=rl_colors.HexColor("#94A3B8"), leading=10)

    story = []

    # ── Top accent bar
    story.append(Table([[""]], colWidths=[17*cm], rowHeights=[10],
        style=TableStyle([
            ("BACKGROUND", (0,0), (-1,-1), rl_colors.HexColor("#0D1B2A")),
            ("LINEBELOW",  (0,0), (-1,-1), 3, rl_colors.HexColor("#C9A84C")),
        ])))
    story.append(Spacer(1, 14))

    # ── Title block
    story.append(Paragraph("AZIMUT INVESTMENTS  ·  PORTAFOGLI MODELLO", eyebrow_s))
    story.append(Spacer(1, 4))
    story.append(Paragraph(f"Portafoglio {ptf_name}", title_s))
    story.append(Paragraph(
        f"{PROFILE_ICONS.get(profile,'●')} Profilo {profile.title()}  ·  "
        f"Generato il {datetime.date.today().strftime('%d %B %Y')}", sub_s))
    story.append(HRFlowable(width="100%", thickness=0.8,
                             color=rl_colors.HexColor("#E2E8F0"), spaceAfter=14))

    # ── KPI row
    d_act = df[df[wcol] > 0.001].copy()
    n      = len(d_act)
    w_az   = (d_act[wcol] * d_act["az_pct"]).sum() * 100
    w_obb  = (d_act[wcol] * d_act["obb_pct"]).sum() * 100

    def kpi_cell(val, label):
        return Paragraph(f'<font size="18"><b>{val}</b></font><br/>'
                         f'<font size="8" color="#64748B">{label}</font>', body_s)

    kpi_tbl = Table(
        [[kpi_cell(str(n), "Fondi"),
          kpi_cell(f"{w_az:.1f}%", "Quota Azionaria"),
          kpi_cell(f"{w_obb:.1f}%", "Quota Obbligazionaria"),
          kpi_cell(datetime.date.today().strftime("%m/%Y"), "Data Report")]],
        colWidths=[4.25*cm]*4
    )
    kpi_tbl.setStyle(TableStyle([
        ("BOX",       (0,0), (-1,-1), 0.8, rl_colors.HexColor("#E2E8F0")),
        ("INNERGRID", (0,0), (-1,-1), 0.8, rl_colors.HexColor("#E2E8F0")),
        ("BACKGROUND",(0,0), (-1,-1), rl_colors.HexColor("#F8FAFC")),
        ("PADDING",   (0,0), (-1,-1), 12),
        ("ALIGN",     (0,0), (-1,-1), "CENTER"),
        ("VALIGN",    (0,0), (-1,-1), "MIDDLE"),
    ]))
    story.append(kpi_tbl)

    # ── Pie chart
    story.append(Paragraph("Distribuzione del Portafoglio", sec_s))
    pie_buf = _make_mpl_pie(d_act, wcol, profile)
    story.append(RLImage(pie_buf, width=14.5*cm, height=8.5*cm))
    story.append(Spacer(1, 10))

    # ── Fund table
    story.append(Paragraph("Composizione Analitica", sec_s))

    d_sorted = d_act.sort_values(wcol, ascending=False).copy()
    d_sorted["pct_s"]  = (d_sorted[wcol]*100).map(lambda v: f"{v:.1f}%")
    d_sorted["az_s"]   = (d_sorted["az_pct"]*100).map(lambda v: f"{v:.0f}%")
    d_sorted["obb_s"]  = (d_sorted["obb_pct"]*100).map(lambda v: f"{v:.0f}%")

    hdr = [Paragraph(f"<b>{t}</b>", body_s)
           for t in ["Fondo", "Categoria", "Macro", "Peso", "AZ%", "OBB%"]]
    tdata = [hdr]

    for _, r in d_sorted.iterrows():
        cat_str = (r["categoria"][:38] + "…") if len(r["categoria"]) > 38 else r["categoria"]
        tdata.append([
            Paragraph(r["nome"][:50], body_s),
            Paragraph(cat_str or "—", body_s),
            Paragraph(r["macro_cat"], body_s),
            Paragraph(f"<b>{r['pct_s']}</b>", body_s),
            Paragraph(r["az_s"],  body_s),
            Paragraph(r["obb_s"], body_s),
        ])

    ftbl = Table(tdata, colWidths=[5.4*cm, 4.5*cm, 2.7*cm, 1.4*cm, 1.2*cm, 1.2*cm], repeatRows=1)
    ts = [
        ("BACKGROUND", (0,0), (-1,0), rl_colors.HexColor("#0D1B2A")),
        ("TEXTCOLOR",  (0,0), (-1,0), rl_colors.white),
        ("FONTNAME",   (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE",   (0,0), (-1,-1), 8),
        ("PADDING",    (0,0), (-1,-1), 6),
        ("ROWBACKGROUNDS",(0,1),(-1,-1), [rl_colors.white, rl_colors.HexColor("#F8FAFC")]),
        ("LINEBELOW",  (0,0), (-1,-1), 0.4, rl_colors.HexColor("#E2E8F0")),
        ("ALIGN",      (3,0), (-1,-1), "CENTER"),
        ("VALIGN",     (0,0), (-1,-1), "MIDDLE"),
    ]
    # Color-code the peso column
    for i, (_, r) in enumerate(d_sorted.iterrows(), 1):
        w = r[wcol]
        if w >= 0.12:
            ts += [("BACKGROUND",(3,i),(3,i), rl_colors.HexColor("#1B4FBB")),
                   ("TEXTCOLOR", (3,i),(3,i), rl_colors.white)]
        elif w >= 0.07:
            ts += [("BACKGROUND",(3,i),(3,i), rl_colors.HexColor("#DBEAFE")),
                   ("TEXTCOLOR", (3,i),(3,i), rl_colors.HexColor("#1B4FBB"))]
    ftbl.setStyle(TableStyle(ts))
    story.append(ftbl)

    # ── Footer
    story.append(Spacer(1, 16))
    story.append(HRFlowable(width="100%", thickness=0.5,
                             color=rl_colors.HexColor("#E2E8F0"), spaceAfter=8))
    story.append(Paragraph(
        "Documento generato automaticamente a scopo illustrativo. I pesi indicati sono riferiti al portafoglio "
        "modello e non costituiscono offerta o consulenza di investimento. Dati soggetti a variazione. "
        "© Azimut Group — uso interno.", foot_s))

    doc.build(story)
    buf.seek(0)
    return buf.read()


# ── FREE PORTFOLIO BUILDER ───────────────────────────────────

def free_portfolio_ui(data: dict) -> pd.DataFrame | None:
    fida = data.get("FIDA", pd.DataFrame())
    if fida.empty:
        st.error("❌ Foglio FIDA non trovato nel file. È necessario per il portafoglio libero.")
        return None

    # Combine fund names from FIDA and PTF sheets for richer AZ% lookup
    az_lookup = {}
    for sname in ["PTF FULL", "PTF SHORT"]:
        if sname in data and not data[sname].empty:
            for _, r in data[sname].iterrows():
                az_lookup[r["nome"]] = r["az_pct"]

    if "free_ptf" not in st.session_state:
        st.session_state.free_ptf = []

    st.markdown('<p class="sec-title">Costruttore Portafoglio Libero</p>', unsafe_allow_html=True)

    options = fida.apply(
        lambda r: f"{r['nome']}  [{r['macro_cat']}]" if r['macro_cat'] != 'Altro' else r['nome'],
        axis=1
    ).tolist()

    c1, c2, c3 = st.columns([3.5, 1, 0.8])
    with c1:
        sel = st.selectbox("Seleziona fondo:", options, key="sel_fund")
    with c2:
        w = st.number_input("Peso %", 0.1, 100.0, 10.0, 0.5, key="sel_w")
    with c3:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("➕ Aggiungi", use_container_width=True):
            fname = sel.split("  [")[0].strip()
            if any(f["nome"] == fname for f in st.session_state.free_ptf):
                st.toast("⚠️ Fondo già presente!", icon="⚠️")
            else:
                fd = fida[fida["nome"] == fname].iloc[0] if not fida[fida["nome"] == fname].empty else None
                mc = fd["macro_cat"] if fd is not None else "Altro"
                az = az_lookup.get(fname, DEFAULT_AZ.get(mc, 0.5))
                st.session_state.free_ptf.append({
                    "nome": fname,
                    "categoria": fd["categoria"] if fd is not None else "",
                    "macro_cat": mc,
                    "az_pct": az,
                    "w_input": w,
                })
                st.rerun()

    if not st.session_state.free_ptf:
        st.info("☝️ Aggiungi almeno un fondo usando il selettore sopra.")
        return None

    st.markdown("**Fondi nel portafoglio:**")
    total_w = 0.0
    for i, fund in enumerate(st.session_state.free_ptf):
        r1, r2, r3 = st.columns([4, 1.5, 0.6])
        with r1:
            st.markdown(f"**{fund['nome']}** <span style='color:#64748b;font-size:.8rem;'>— {fund['macro_cat']}</span>",
                        unsafe_allow_html=True)
        with r2:
            nw = st.number_input("Peso", 0.0, 100.0, float(fund["w_input"]), 0.5,
                                  key=f"fw_{i}", label_visibility="collapsed")
            st.session_state.free_ptf[i]["w_input"] = nw
        with r3:
            if st.button("🗑️", key=f"del_{i}", use_container_width=True):
                st.session_state.free_ptf.pop(i)
                st.rerun()
        total_w += st.session_state.free_ptf[i]["w_input"]

    diff = abs(total_w - 100.0)
    if diff < 0.05:
        st.markdown(f'<div class="w-ok">✅ Somma pesi: <b>{total_w:.1f}%</b> — Portafoglio bilanciato!</div>',
                    unsafe_allow_html=True)
    else:
        st.markdown(f'<div class="w-warn">⚠️ Somma pesi: <b>{total_w:.1f}%</b> — Deve essere 100% (mancano {100-total_w:+.1f}%)</div>',
                    unsafe_allow_html=True)

    if diff > 0.5:
        return None

    records = []
    for fund in st.session_state.free_ptf:
        az = fund.get("az_pct", DEFAULT_AZ.get(fund["macro_cat"], 0.5))
        records.append({
            "nome":      fund["nome"],
            "categoria": fund["categoria"],
            "gruppo":    fund["macro_cat"],
            "macro_cat": fund["macro_cat"],
            "az_pct":    az,
            "obb_pct":   1.0 - az,
            "r_weight":  fund["w_input"] / 100.0,
            "w_cons":    fund["w_input"] / 100.0,
            "w_equil":   fund["w_input"] / 100.0,
            "w_accr":    fund["w_input"] / 100.0,
            "w_centrale": fund["w_input"] / 100.0,
        })

    df = pd.DataFrame(records)
    df = assign_colors(df)
    # Normalize
    for wc in ["w_cons", "w_equil", "w_accr", "w_centrale"]:
        t = df[wc].sum()
        df[wc] = df[wc] / t if t > 0 else df[wc]
    return df


# ── MAIN APP ─────────────────────────────────────────────────

def main():

    # ── SIDEBAR ──────────────────────────────────────────────
    with st.sidebar:
        st.markdown("""
        <div style='padding:1.4rem 0 .8rem 0;'>
          <div style='font-size:.6rem;letter-spacing:.22em;color:#3a5a78;text-transform:uppercase;font-weight:700;'>
            Strumento di Analisi
          </div>
          <div style='font-family:"Cormorant Garamond",serif;font-size:1.6rem;color:#dde8f5;font-weight:700;margin-top:4px;line-height:1.2;'>
            Portfolio<br>Analyzer
          </div>
          <div style='width:32px;height:3px;background:#C9A84C;border-radius:2px;margin-top:10px;'></div>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("---")

        uploaded = st.file_uploader(
            "FILE EXCEL (PTF FULL + PTF SHORT + FIDA)",
            type=["xlsx", "xls"],
        )
        st.markdown("---")

        ptf_choice = st.radio(
            "TIPO PORTAFOGLIO",
            options=["📋  PTF FULL", "⚡  PTF SHORT", "🎨  LIBERO"],
        )
        st.markdown("---")

        profile = st.selectbox("PROFILO DI RISCHIO", PROFILES, index=3)

        # Reset free portfolio if switching away from LIBERO
        if "LIBERO" not in ptf_choice and "free_ptf" in st.session_state:
            del st.session_state["free_ptf"]

    # ── HEADER ───────────────────────────────────────────────
    ptf_label = ptf_choice.split("  ", 1)[1] if "  " in ptf_choice else ptf_choice

    st.markdown(f"""
    <div class="az-header">
      <div class="az-eyebrow">AZIMUT INVESTMENTS · PORTAFOGLI MODELLO</div>
      <div class="az-rule"></div>
      <div class="az-title">{ptf_label}</div>
      <div class="az-meta">
        {PROFILE_ICONS.get(profile,'●')} Profilo {profile.title()}
        &nbsp;·&nbsp; {datetime.date.today().strftime('%d %B %Y')}
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── ONBOARDING ───────────────────────────────────────────
    if uploaded is None:
        st.info("⬅️ **Carica il file Excel** nella barra laterale per iniziare.")
        with st.expander("📖 Struttura richiesta del file Excel", expanded=True):
            st.markdown("""
| Foglio | Contenuto |
|--------|-----------|
| **PTF FULL** | Portafoglio completo (colonne: Nome fondo, Categoria, pesi per profilo, % azionaria) |
| **PTF SHORT** | Versione semplificata con meno fondi |
| **FIDA** | Elenco di tutti i fondi con ISIN e categoria (per portafoglio libero) |
            """)
        return

    # ── LOAD DATA ────────────────────────────────────────────
    with st.spinner("⏳ Elaborazione file in corso…"):
        raw = parse_excel(uploaded.read())

    # ── SELECT PORTFOLIO ─────────────────────────────────────
    if "LIBERO" in ptf_choice:
        df = free_portfolio_ui(raw)
    else:
        key = "PTF FULL" if "FULL" in ptf_choice else "PTF SHORT"
        if key not in raw or raw[key].empty:
            st.error(f"❌ Foglio '{key}' non trovato o vuoto nel file Excel.")
            return
        df = raw[key]

    if df is None or df.empty:
        return

    wcol   = PROFILE_W_COL[profile]
    df_act = df[df[wcol] > 0.001].copy()

    # ── KPI ROW ──────────────────────────────────────────────
    n_fondi = len(df_act)
    w_az    = (df_act[wcol] * df_act["az_pct"]).sum()  * 100
    w_obb   = (df_act[wcol] * df_act["obb_pct"]).sum() * 100
    srri    = max(1, min(7, round(w_az / 100 * 6 + 1)))

    c1, c2, c3, c4 = st.columns(4)
    for col, val, lbl, sub in [
        (c1, str(n_fondi),          "Fondi in Portafoglio",       f"{df_act['gruppo'].nunique()} gruppi"),
        (c2, f"{w_az:.1f}%",        "Quota Azionaria",            "ponderata per peso"),
        (c3, f"{w_obb:.1f}%",       "Quota Obbligazionaria",      "ponderata per peso"),
        (c4, f"{srri} / 7",         "Risk Score (SRRI proxy)",    "basato su quota azionaria"),
    ]:
        col.markdown(
            f'<div class="kpi"><div class="kpi-label">{lbl}</div>'
            f'<div class="kpi-value">{val}</div>'
            f'<div class="kpi-sub">{sub}</div></div>',
            unsafe_allow_html=True,
        )

    st.markdown("<br>", unsafe_allow_html=True)

    # ── CHARTS + FUND LIST ────────────────────────────────────
    col_l, col_r = st.columns([1.15, 0.85], gap="large")

    with col_l:
        st.markdown('<p class="sec-title">Allocazione per Fondo</p>', unsafe_allow_html=True)
        st.plotly_chart(make_fund_pie(df_act, wcol, profile),
                        use_container_width=True, config={"displayModeBar": False})

        st.markdown('<p class="sec-title">Allocazione per Macro-Categoria</p>', unsafe_allow_html=True)
        st.plotly_chart(make_macro_bar(df_act, wcol),
                        use_container_width=True, config={"displayModeBar": False})

    with col_r:
        st.markdown('<p class="sec-title">Composizione del Portafoglio</p>', unsafe_allow_html=True)

        for gruppo in df_act["gruppo"].unique():
            sub = df_act[df_act["gruppo"] == gruppo].sort_values(wcol, ascending=False)
            rows_html = "".join([
                f"""<div class="fund-row">
                  <div class="fund-dot" style="background:{row['color']};"></div>
                  <div style="flex:1;min-width:0;">
                    <div class="fund-name">{row['nome']}</div>
                    <div class="fund-cat">{row['categoria'][:48] + '…' if row['categoria'] and len(row['categoria'])>48 else (row['categoria'] or '—')}</div>
                  </div>
                  <div class="fund-pct">{row[wcol]*100:.1f}%</div>
                </div>"""
                for _, row in sub.iterrows()
            ])
            st.markdown(
                f'<div style="background:#fff;border:1px solid #e2e8f0;border-radius:10px;margin-bottom:12px;overflow:hidden;">'
                f'<div class="fund-group-hdr">{gruppo}</div>{rows_html}</div>',
                unsafe_allow_html=True,
            )

    # ── DOWNLOAD ─────────────────────────────────────────────
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<p class="sec-title">Esporta Report PDF</p>', unsafe_allow_html=True)

    col_btn, col_inf = st.columns([1, 2])
    with col_btn:
        try:
            pdf_data = generate_pdf(df_act, wcol, profile, ptf_label)
            fname = (f"Azimut_{ptf_label.replace(' ', '_')}_"
                     f"{profile}_{datetime.date.today().strftime('%Y%m%d')}.pdf")
            st.download_button(
                label="📥   Scarica Report PDF",
                data=pdf_data,
                file_name=fname,
                mime="application/pdf",
                use_container_width=True,
            )
        except Exception as e:
            st.error(f"Errore generazione PDF: {e}")

    with col_inf:
        st.markdown(f"""
        <div style='background:#eff6ff;border:1px solid #bfdbfe;border-radius:10px;padding:1rem 1.25rem;'>
          <div style='font-size:.8rem;color:#1d4ed8;font-weight:600;margin-bottom:.4rem;'>Il report PDF contiene:</div>
          <div style='font-size:.82rem;color:#1e40af;line-height:1.9;'>
            ✓ Intestazione professionale con data e profilo<br>
            ✓ Riepilogo KPI (fondi, quota az./obb., data)<br>
            ✓ Grafico a torta con legenda colorata<br>
            ✓ Tabella completa con pesi e breakdown AZ/OBB<br>
            ✓ Footer con disclaimer legale
          </div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<br><br>", unsafe_allow_html=True)


if __name__ == "__main__":
    main()
