# ============================================================
# AZIMUT PORTFOLIO BUILDER v1.0 — app2.py
# App separata da app.py — NON sovrascrive il Portfolio Analyzer
# ============================================================
import re, io, json, datetime
import streamlit as st
import pandas as pd
import numpy as np
import requests
from bs4 import BeautifulSoup
from concurrent.futures import ThreadPoolExecutor, as_completed
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.ticker as mticker
from reportlab.lib.pagesizes import A4
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    Image as RLImage, HRFlowable, PageBreak, KeepTogether,
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors as rl_colors
from reportlab.lib.units import cm

try:
    import anthropic as _anthropic_mod
    _HAS_ANTHROPIC = True
except ImportError:
    _HAS_ANTHROPIC = False

try:
    from pypdf import PdfReader, PdfWriter
    _HAS_PYPDF = True
except ImportError:
    _HAS_PYPDF = False

# ── PAGE CONFIG ──────────────────────────────────────────────
st.set_page_config(
    page_title="Azimut | Portfolio Builder",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── COSTANTI ─────────────────────────────────────────────────
MACRO_COLORS = {
    "Azionari": "#1B4FBB", "Bilanciati/Flessibili": "#C9A84C",
    "Obbligazionari": "#2D9D78", "Alternativi": "#8B5CF6",
    "Monetari": "#F59E0B", "Altro": "#94A3B8",
}
SHADES = {
    "Azionari":             ["#0D3080","#1B4FBB","#2563EB","#3B82F6","#60A5FA","#93C5FD"],
    "Bilanciati/Flessibili":["#92650A","#B8860B","#C9A84C","#D4B572","#DFC298","#E9CEB4"],
    "Obbligazionari":       ["#065F46","#14855F","#2D9D78","#34B98A","#6DE5BC","#9AEFD2"],
    "Alternativi":          ["#5B21B6","#7C3AED","#8B5CF6","#A78BFA","#C4B5FD"],
    "Monetari":             ["#D97706","#F59E0B","#FCD34D"],
    "Altro":                ["#475569","#64748B","#94A3B8","#CBD5E1"],
}
MS_HDR = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120",
    "Accept": "application/json, text/html",
    "Accept-Language": "it-IT,it;q=0.9,en;q=0.8",
}
FD_HDR = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120",
    "Accept": "text/html,application/xhtml+xml",
    "Accept-Language": "en-US,en;q=0.9,it;q=0.8",
}

# ── CSS ──────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@600;700&family=DM+Sans:opsz,wght@9..40,300;9..40,400;9..40,500;9..40,600&display=swap');
html,body,[class*="css"]{font-family:'DM Sans',sans-serif;}
h1,h2,h3{font-family:'Cormorant Garamond',serif !important;}
[data-testid="stSidebar"]{background:linear-gradient(170deg,#06101e 0%,#0d1f3c 55%,#0a1628 100%);border-right:1px solid #1a3050;}
[data-testid="stSidebar"] label{color:#4a6582 !important;font-size:.68rem !important;letter-spacing:.12em !important;text-transform:uppercase !important;font-weight:600 !important;}
.main{background:#f6f8fb !important;}.block-container{padding-top:1.8rem !important;max-width:1300px;}
.az-header{background:linear-gradient(130deg,#081420 0%,#0f2644 50%,#162e52 100%);border-radius:16px;padding:2rem 2.5rem;margin-bottom:1.8rem;box-shadow:0 8px 32px rgba(0,0,0,.15);}
.az-eyebrow{font-size:.65rem;letter-spacing:.2em;color:#4a7098;text-transform:uppercase;font-weight:600;}
.az-title{font-family:'Cormorant Garamond',serif;font-size:2.1rem;font-weight:700;color:#f0f6ff;margin:.2rem 0 .4rem;line-height:1.1;}
.az-rule{width:38px;height:3px;background:#C9A84C;border-radius:2px;margin:.6rem 0;}
.az-meta{font-size:.88rem;color:#6b8fb0;}
.kpi{background:#fff;border:1px solid #e4eaf3;border-radius:12px;padding:1.2rem 1.4rem;box-shadow:0 1px 4px rgba(0,0,0,.05);}
.kpi-label{font-size:.65rem;text-transform:uppercase;letter-spacing:.1em;color:#94a3b8;font-weight:500;margin-bottom:.3rem;}
.kpi-value{font-size:1.9rem;font-weight:700;color:#0d1b2a;font-family:'Cormorant Garamond',serif;line-height:1;}
.kpi-sub{font-size:.75rem;color:#64748b;margin-top:.3rem;}
.sec-title{font-family:'Cormorant Garamond',serif;font-size:1.25rem;font-weight:600;color:#0d1b2a;border-bottom:2px solid #c9a84c;display:inline-block;padding-bottom:.4rem;margin-bottom:.9rem;}
[data-testid="stDownloadButton"]>button{background:linear-gradient(135deg,#0f2d6b 0%,#1b4fbb 100%) !important;color:#fff !important;font-size:1.05rem !important;font-weight:600 !important;padding:.9rem 2rem !important;border-radius:10px !important;border:none !important;width:100% !important;}
</style>
""", unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════
# HELPERS
# ════════════════════════════════════════════════════════════

def get_macro(cat: str) -> str:
    if not cat or cat == "-": return "Altro"
    c = cat.lower()
    if "azionari" in c or "equity" in c: return "Azionari"
    if any(x in c for x in ["obbligazionari","bond","credit","debt","reddito"]): return "Obbligazionari"
    if any(x in c for x in ["bilanciati","allocation","flessibili","balanced","flexible"]): return "Bilanciati/Flessibili"
    if any(x in c for x in ["alternativi","alternative","commodity"]): return "Alternativi"
    if any(x in c for x in ["monetari","money market","liquidità"]): return "Monetari"
    return "Altro"


def assign_colors(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy(); counter = {k: 0 for k in SHADES}; cols = []
    for _, row in df.iterrows():
        mc = row.get("macro_cat", "Altro")
        sh = SHADES.get(mc, SHADES["Altro"])
        cols.append(sh[counter.get(mc,0) % len(sh)])
        counter[mc] = counter.get(mc,0)+1
    df["color"] = cols; return df


def _norm_w(lst: list) -> list:
    tot = sum(f.get("peso",0) for f in lst)
    if tot <= 0:
        n = len(lst)
        return [{**f,"peso":round(100/n,1)} for f in lst]
    return [{**f,"peso":round(f.get("peso",0)*100/tot,1)} for f in lst]


# ════════════════════════════════════════════════════════════
# FILE PARSERS
# ════════════════════════════════════════════════════════════

def _extract_text_from_pdf(fb: bytes) -> str:
    """Estrae testo grezzo da un PDF usando pypdf."""
    if not _HAS_PYPDF:
        return ""
    try:
        reader = PdfReader(io.BytesIO(fb))
        return "\n".join(p.extract_text() or "" for p in reader.pages)
    except Exception:
        return ""


def _parse_funds_from_text(text: str) -> pd.DataFrame:
    """
    Dato testo grezzo (da PDF), estrae una lista fondi cercando pattern ISIN.
    Per ogni ISIN trovato prende le righe vicine come candidato nome.
    """
    ISIN_RE = re.compile(r'\b([A-Z]{2}[A-Z0-9]{10})\b')
    lines = [l.strip() for l in text.split("\n") if l.strip()]
    records = []
    seen_isin = set()

    for i, line in enumerate(lines):
        for m in ISIN_RE.finditer(line):
            isin = m.group(1)
            if isin in seen_isin:
                continue
            seen_isin.add(isin)

            # Cerca il nome nelle righe circostanti (prima dell'ISIN di solito)
            name_candidates = []
            for offset in [-1, -2, 1, -3, 2]:
                idx = i + offset
                if 0 <= idx < len(lines):
                    cand = lines[idx].strip()
                    # Scarta righe che contengono altri ISIN, numeri puri, o sono troppo corte
                    if (cand and len(cand) > 6
                            and not ISIN_RE.search(cand)
                            and not re.match(r'^[\d\s\.\,\%\-\+]+$', cand)
                            and not re.match(r'^\d{1,2}/\d{1,2}/\d{2,4}$', cand)):
                        name_candidates.append(cand)
            # Prendi il candidato più lungo come nome (solitamente il più descrittivo)
            nome = max(name_candidates, key=len) if name_candidates else isin

            # Inferisci categoria dal contesto
            ctx = " ".join(lines[max(0,i-4):i+5]).lower()
            cat = ""
            if any(w in ctx for w in ["azionari","equity","azioni","stock"]): cat = "Azionari"
            elif any(w in ctx for w in ["obbligazion","bond","credit","reddito fisso","fixed income"]): cat = "Obbligazionari"
            elif any(w in ctx for w in ["bilanc","flessib","balanced","allocation","flexible"]): cat = "Bilanciati/Flessibili"
            elif any(w in ctx for w in ["monetar","money market","liquidit"]): cat = "Monetari"
            elif any(w in ctx for w in ["alternativ","commodity","real asset"]): cat = "Alternativi"

            # Inferisci MIFID/SRRI dal contesto
            mifid = 4
            srri_m = re.search(r'(?:srri|mifid|rischio)[^\d]{0,10}([1-7])', ctx)
            if not srri_m:
                srri_m = re.search(r'\b([1-7])\s*/\s*7\b', ctx)
            if srri_m:
                try: mifid = int(srri_m.group(1))
                except: pass

            records.append({"nome": nome, "isin": isin, "categoria": cat,
                            "descrizione": "", "az_pct": 0.5, "mifid": mifid})

    if not records:
        return pd.DataFrame()
    df = pd.DataFrame(records).drop_duplicates("isin").reset_index(drop=True)
    df["macro_cat"] = df["categoria"].apply(get_macro)
    return assign_colors(df)


@st.cache_data(show_spinner=False)
def parse_funds_file(fb: bytes, fname: str) -> pd.DataFrame:
    # ── PDF: estrai ISIN + metadati dal testo ──────────────────
    if fname.lower().endswith(".pdf"):
        text = _extract_text_from_pdf(fb)
        if not text.strip():
            st.error("PDF vuoto o non leggibile con pypdf."); return pd.DataFrame()
        df = _parse_funds_from_text(text)
        if df.empty:
            st.error("Nessun ISIN trovato nel PDF. Carica un Excel/CSV con la lista fondi.")
        else:
            st.success(f"📄 PDF: estratti {len(df)} fondi tramite ISIN.")
        return df

    # ── Excel / CSV ───────────────────────────────────────────
    try:
        df = pd.read_csv(io.BytesIO(fb)) if fname.lower().endswith(".csv") else pd.read_excel(io.BytesIO(fb))
    except Exception as e:
        st.error(f"Errore lettura fondi: {e}"); return pd.DataFrame()
    df.columns = [str(c).strip().lower() for c in df.columns]
    col_map = {
        "nome":       ["nome","fund name","nome fondo","name","fondo"],
        "isin":       ["isin","codice isin","isin code"],
        "categoria":  ["categoria","category","asset class","tipo","classificazione"],
        "descrizione":["descrizione","description","note","obiettivo"],
        "az_pct":     ["az_pct","quota azionaria","azionario","equity %","equity pct"],
        "mifid":      ["mifid","srri","rischio","risk score","profilo rischio"],
    }
    rename = {}
    for std, variants in col_map.items():
        for v in variants:
            if v in df.columns: rename[v] = std; break
    df = df.rename(columns=rename)
    if "nome" not in df.columns:
        st.error("Colonna 'nome' non trovata."); return pd.DataFrame()
    df["nome"] = df["nome"].astype(str).str.strip()
    df["isin"] = df.get("isin", pd.Series("", index=df.index)).astype(str).str.strip().str.upper()
    df["isin"] = df["isin"].apply(lambda x: x if re.match(r'^[A-Z]{2}[A-Z0-9]{10}$',x) else "")
    df["categoria"]   = df.get("categoria",   pd.Series("", index=df.index)).fillna("").astype(str)
    df["descrizione"] = df.get("descrizione", pd.Series("", index=df.index)).fillna("").astype(str)
    df["az_pct"]  = pd.to_numeric(df.get("az_pct",  pd.Series(0.5, index=df.index)), errors="coerce").fillna(0.5).clip(0,1)
    df["mifid"]   = pd.to_numeric(df.get("mifid",   pd.Series(4,   index=df.index)), errors="coerce").fillna(4).clip(1,7).astype(int)
    df["macro_cat"] = df["categoria"].apply(get_macro)
    df = df[df["nome"].notna() & (df["nome"] != "") & (df["nome"] != "nan")].drop_duplicates("nome").reset_index(drop=True)
    return assign_colors(df)


@st.cache_data(show_spinner=False)
def parse_mifid_file(fb: bytes, fname: str) -> pd.DataFrame:
    # ── PDF: estrai ISIN + SRRI/MIFID dal testo ───────────────
    if fname.lower().endswith(".pdf"):
        text = _extract_text_from_pdf(fb)
        if not text.strip():
            st.warning("MIFID PDF vuoto o non leggibile."); return pd.DataFrame()
        ISIN_RE = re.compile(r'\b([A-Z]{2}[A-Z0-9]{10})\b')
        lines = [l.strip() for l in text.split("\n") if l.strip()]
        records = []; seen = set()
        for i, line in enumerate(lines):
            for m in ISIN_RE.finditer(line):
                isin = m.group(1)
                if isin in seen: continue
                seen.add(isin)
                ctx = " ".join(lines[max(0,i-3):i+4]).lower()
                mifid = 4
                srri_m = re.search(r'(?:srri|mifid|rischio|risk)[^\d]{0,12}([1-7])', ctx)
                if not srri_m:
                    srri_m = re.search(r'\b([1-7])\s*/\s*7\b', ctx)
                if srri_m:
                    try: mifid = int(srri_m.group(1))
                    except: pass
                records.append({"isin": isin, "mifid": mifid})
        if not records:
            st.warning("Nessun ISIN trovato nel PDF MIFID."); return pd.DataFrame()
        st.success(f"📄 MIFID PDF: {len(records)} fondi estratti.")
        return pd.DataFrame(records).drop_duplicates("isin").reset_index(drop=True)

    # ── Excel / CSV ───────────────────────────────────────────
    try:
        df = pd.read_csv(io.BytesIO(fb)) if fname.lower().endswith(".csv") else pd.read_excel(io.BytesIO(fb))
    except Exception as e:
        st.error(f"Errore MIFID: {e}"); return pd.DataFrame()
    df.columns = [str(c).strip().lower() for c in df.columns]
    name_col  = next((c for c in df.columns if any(x in c for x in ["nome","name","fondo","fund"])), None)
    isin_col  = next((c for c in df.columns if "isin" in c), None)
    score_col = next((c for c in df.columns if any(x in c for x in ["mifid","srri","score","rischio","risk"])), None)
    if not (name_col or isin_col) or not score_col:
        st.warning("File MIFID: colonne non riconosciute."); return pd.DataFrame()
    out = pd.DataFrame()
    if name_col: out["nome"] = df[name_col].astype(str).str.strip()
    if isin_col: out["isin"] = df[isin_col].astype(str).str.strip().str.upper()
    out["mifid"] = pd.to_numeric(df[score_col], errors="coerce").fillna(4).clip(1,7).astype(int)
    return out


@st.cache_data(show_spinner=False)
def parse_unp_catalog(fb: bytes, fname: str) -> pd.DataFrame:
    """
    Legge il catalogo prodotti UNP/IUNP Azimut.
    Accetta Excel, CSV o PDF.
    Restituisce DataFrame con colonne: isin, nome, unp, iunp
    UNP = Utile Netto di Portafoglio (% annuo commissione netta advisor)
    """
    ISIN_RE = re.compile(r'\b([A-Z]{2}[A-Z0-9]{10})\b')

    # ── PDF ───────────────────────────────────────────────────
    if fname.lower().endswith(".pdf"):
        text = _extract_text_from_pdf(fb)
        if not text.strip():
            st.warning("UNP PDF: testo non estraibile."); return pd.DataFrame()
        lines  = [l.strip() for l in text.split("\n") if l.strip()]
        # Regex percentuale: cattura "0,75" / "0.75" / "75" (poi normalizziamo)
        PCT_RE = re.compile(r'(\d{1,3}[,\.]\d{1,4})')
        records = []; seen = set()
        for i, line in enumerate(lines):
            for m in ISIN_RE.finditer(line):
                isin = m.group(1)
                if isin in seen: continue
                seen.add(isin)
                # cerca UNP e IUNP nelle 3 righe successive
                ctx = " ".join(lines[i:i+4])
                nums = PCT_RE.findall(ctx)
                nums_f = []
                for n in nums:
                    try:
                        v = float(n.replace(",","."))
                        if 0 < v < 10: nums_f.append(round(v, 4))   # già in %
                        elif 10 <= v <= 500: nums_f.append(round(v/100, 4))  # basis points
                    except: pass
                # nome = prima stringa lunga non-ISIN attorno alla riga
                nome = ""
                for j in range(max(0,i-1), min(len(lines),i+3)):
                    if not ISIN_RE.search(lines[j]) and len(lines[j]) > 8:
                        nome = lines[j][:80]; break
                records.append({
                    "isin":  isin,
                    "nome":  nome,
                    "unp":   nums_f[0] if len(nums_f) > 0 else None,
                    "iunp":  nums_f[1] if len(nums_f) > 1 else None,
                })
        if not records:
            st.warning("UNP PDF: nessun ISIN trovato."); return pd.DataFrame()
        st.success(f"📊 UNP catalog: {len(records)} fondi estratti dal PDF.")
        return pd.DataFrame(records)

    # ── Excel / CSV ───────────────────────────────────────────
    try:
        df = pd.read_csv(io.BytesIO(fb)) if fname.lower().endswith(".csv") \
             else pd.read_excel(io.BytesIO(fb))
    except Exception as e:
        st.error(f"Errore lettura UNP: {e}"); return pd.DataFrame()
    df.columns = [str(c).strip().lower() for c in df.columns]
    isin_col = next((c for c in df.columns if "isin" in c), None)
    nome_col = next((c for c in df.columns if any(x in c for x in
                     ["nome","name","fondo","fund","prodotto"])), None)
    unp_col  = next((c for c in df.columns if "unp"  in c and "i" not in c[:2]), None)
    iunp_col = next((c for c in df.columns if "iunp" in c), None)
    if not isin_col or not unp_col:
        # fallback: prendi le prime colonne numeriche dopo ISIN
        if isin_col:
            num_cols = [c for c in df.columns
                        if pd.to_numeric(df[c], errors="coerce").notna().sum() > len(df)*0.5]
            if num_cols: unp_col  = num_cols[0]
            if len(num_cols)>1: iunp_col = num_cols[1]
    if not isin_col:
        st.warning("Catalogo UNP: colonna ISIN non trovata."); return pd.DataFrame()
    out = pd.DataFrame()
    out["isin"] = df[isin_col].astype(str).str.strip().str.upper()
    out["isin"] = out["isin"].apply(lambda x: x if re.match(r'^[A-Z]{2}[A-Z0-9]{10}$',x) else "")
    if nome_col: out["nome"] = df[nome_col].astype(str).str.strip()
    else:        out["nome"] = ""
    if unp_col:
        out["unp"] = pd.to_numeric(
            df[unp_col].astype(str).str.replace("%","").str.replace(",","."),
            errors="coerce")
        # normalizza: se valori > 10 sono basis points, converti in %
        if out["unp"].median(skipna=True) > 5:
            out["unp"] = out["unp"] / 100
    else:
        out["unp"] = None
    if iunp_col:
        out["iunp"] = pd.to_numeric(
            df[iunp_col].astype(str).str.replace("%","").str.replace(",","."),
            errors="coerce")
        if out["iunp"].median(skipna=True) > 5:
            out["iunp"] = out["iunp"] / 100
    else:
        out["iunp"] = None
    out = out[out["isin"] != ""].drop_duplicates("isin").reset_index(drop=True)
    if out.empty:
        st.warning("Catalogo UNP: nessun ISIN valido trovato."); return pd.DataFrame()
    st.success(f"📊 UNP catalog: {len(out)} fondi caricati.")
    return out


def read_text_file(fb: bytes, fname: str) -> str:
    if fname.lower().endswith((".txt",".md")):
        return fb.decode("utf-8", errors="ignore")
    if fname.lower().endswith(".pdf"):
        if _HAS_PYPDF:
            try:
                reader = PdfReader(io.BytesIO(fb))
                return "\n".join(p.extract_text() or "" for p in reader.pages)
            except Exception: pass
        return "[PDF: installa pypdf per estrarre testo]"
    return fb.decode("utf-8", errors="ignore")


def _parse_schede_alloc(pdf_bytes_list: list, funds_df: pd.DataFrame) -> dict:
    """
    Legge i PDF delle schede prodotto e tenta di estrarre l'asset allocation
    per ogni fondo, riconoscendolo tramite ISIN.
    Returns: {isin: {"Azionari": float%, "Obbligazionari": float%, ...}}
    """
    if not pdf_bytes_list or not _HAS_PYPDF:
        return {}
    ISIN_RE  = re.compile(r'\b([A-Z]{2}[A-Z0-9]{10})\b')
    known    = set(funds_df["isin"].dropna().astype(str).unique())
    result   = {}

    # Pattern per ogni macro-categoria (ricerca su testo lowercase)
    CAT_PAT = {
        "Azionari": [
            r"azion\w*\s*[\:\-]?\s*([\d]{1,3}(?:[,\.]\d+)?)\s*%",
            r"equity\s*[\:\-]?\s*([\d]{1,3}(?:[,\.]\d+)?)\s*%",
            r"([\d]{1,3}(?:[,\.]\d+)?)\s*%\s*(?:azion\w*|equity)",
        ],
        "Obbligazionari": [
            r"obbligazion\w*\s*[\:\-]?\s*([\d]{1,3}(?:[,\.]\d+)?)\s*%",
            r"(?:bond|reddito\s+fisso|fixed\s+income)\s*[\:\-]?\s*([\d]{1,3}(?:[,\.]\d+)?)\s*%",
            r"([\d]{1,3}(?:[,\.]\d+)?)\s*%\s*(?:obbligazion\w*|bond)",
        ],
        "Monetari": [
            r"monetar\w*\s*[\:\-]?\s*([\d]{1,3}(?:[,\.]\d+)?)\s*%",
            r"(?:liquidit[àa]|money\s+market)\s*[\:\-]?\s*([\d]{1,3}(?:[,\.]\d+)?)\s*%",
        ],
        "Alternativi": [
            r"alternativ\w*\s*[\:\-]?\s*([\d]{1,3}(?:[,\.]\d+)?)\s*%",
            r"commodity\s*[\:\-]?\s*([\d]{1,3}(?:[,\.]\d+)?)\s*%",
        ],
    }

    for fb in pdf_bytes_list:
        text = _extract_text_from_pdf(fb)
        if not text:
            continue
        # Trova ISIN noti in questo PDF
        found = [i for i in ISIN_RE.findall(text) if i in known]
        if not found:
            continue
        isin = found[0]  # Usa il primo ISIN noto trovato
        t    = text.lower()
        alloc = {}
        for cat, patterns in CAT_PAT.items():
            for pat in patterns:
                m = re.search(pat, t)
                if m:
                    try:
                        v = float(m.group(1).replace(",", "."))
                        if 0 < v <= 100:
                            alloc[cat] = v
                            break
                    except Exception:
                        pass
        if alloc:
            # Normalizza se la somma supera 100
            tot = sum(alloc.values())
            if tot > 102:
                alloc = {k: round(v * 100 / tot, 1) for k, v in alloc.items()}
            result[isin] = alloc

    return result


# ════════════════════════════════════════════════════════════
# DATA FETCHERS — Morningstar + FondiDoc
# ════════════════════════════════════════════════════════════

@st.cache_data(ttl=86400, show_spinner=False)
def fetch_morningstar(isin: str) -> dict:
    if not isin or len(isin) != 12: return {}
    try:
        r = requests.get(
            "https://www.morningstar.it/it/util/SecuritySearch.ashx",
            params={"q":isin,"t":"FO","limit":"1","langId":"it","currencyId":"EUR"},
            headers=MS_HDR, timeout=6)
        if r.status_code == 200 and r.text.strip().startswith("["):
            d = r.json()
            if d and isinstance(d, list):
                sid = d[0].get("id","")
                return {
                    "ms_name": d[0].get("name",""),
                    "ms_id":   sid,
                    "ms_url":  f"https://www.morningstar.it/it/funds/snapshot/snapshot.aspx?id={sid}" if sid else "",
                    "ms_cat":  d[0].get("universe",""),
                }
    except Exception: pass
    return {}


def _fd_en(url):
    return url if "/en/" in url else url.replace("fondidoc.it/d/","fondidoc.it/en/d/")

def _fd_ana(url):
    return url.replace("/d/Index/","/d/Ana/").replace("/en/d/Index/","/en/d/Ana/")

def _fd_fetch(url):
    try:
        r = requests.get(_fd_en(url), headers=FD_HDR, timeout=8)
        return r.text if r.status_code == 200 else None
    except: return None

def _fd_overview(html):
    soup = BeautifulSoup(html,"lxml")
    lines = [l.strip() for l in soup.get_text(separator="\n").split("\n") if l.strip()]
    d = {}
    for i,line in enumerate(lines):
        nxt = lines[i+1] if i+1<len(lines) else ""
        if line=="SRRI (risk value)":        d["srri"]=nxt
        elif line=="Start date":             d["start_date"]=nxt
        elif line=="Assogestioni category":  d["cat_assog"]=nxt
        elif line=="Management Fee":         d["mgmt_fee"]=nxt
        elif line=="Performance Fee":        d["perf_fee"]=nxt
        elif line=="Rating" and "fida_rating" not in d: d["fida_rating"]=nxt
        elif line=="Score":                  d["fida_score"]=nxt
        elif line=="Category" and "fida_cat" not in d: d["fida_cat"]=nxt
    return d

def _fd_analysis(html):
    soup = BeautifulSoup(html,"lxml")
    lines = [l.strip() for l in soup.get_text(separator="\n").split("\n") if l.strip()]
    d = {}
    for i,line in enumerate(lines):
        nxt = lines[i+1] if i+1<len(lines) else ""
        if line=="NAV": d["nav"]=nxt
        elif line=="Last update": d["last_update"]=nxt
    for table in soup.find_all("table"):
        rows = table.find_all("tr")
        if not rows: continue
        hdr = [td.get_text(strip=True) for td in rows[0].find_all(["th","td"])]
        if "YTD" in hdr and "1 year" in hdr:
            def _sg(cells,key,h):
                try: return cells[h.index(key)] if h.index(key)<len(cells) else "—"
                except: return "—"
            for row in rows[1:]:
                cells=[td.get_text(strip=True) for td in row.find_all("td")]
                if not cells: continue
                lbl=cells[0].lower()
                if "performance" in lbl:
                    d["ytd"]=_sg(cells,"YTD",hdr); d["perf_1y"]=_sg(cells,"1 year",hdr)
                    d["perf_3y"]=_sg(cells,"3 years",hdr); d["perf_5y"]=_sg(cells,"5 years",hdr)
        elif "1 year" in hdr and "YTD" not in hdr and len(hdr)>=4:
            for row in rows[1:]:
                cells=[td.get_text(strip=True) for td in row.find_all("td")]
                if not cells: continue
                lbl=cells[0].lower()
                def gv(i): return cells[i] if i<len(cells) else "—"
                if "volatility" in lbl and "negative" not in lbl:
                    d["vol_1y"],d["vol_3y"],d["vol_5y"]=gv(1),gv(2),gv(3)
                elif "negative" in lbl:
                    d["neg_vol_1y"]=gv(1)
                    if len(hdr)>2: d["neg_vol_3y"]=gv(2)
                    if len(hdr)>3: d["neg_vol_5y"]=gv(3)
                elif "sharpe" in lbl:
                    d["sharpe_1y"],d["sharpe_3y"],d["sharpe_5y"]=gv(1),gv(2),gv(3)
                elif "sortino" in lbl:
                    d["sortino_1y"]=gv(1)
                elif "var" in lbl or "value at risk" in lbl:
                    d["var_1y"]=gv(1)
                    if len(hdr)>2: d["var_3y"]=gv(2)
        elif any(h.isdigit() and len(h)==4 for h in hdr):
            years=[h for h in hdr if h.isdigit() and len(h)==4]
            for row in rows[1:]:
                cells=[td.get_text(strip=True) for td in row.find_all("td")]
                if len(cells)<2: continue
                if any("%"in c for c in cells) and "annual_perf" not in d:
                    ann={}
                    for yr in years:
                        try: idx=hdr.index(yr); ann[yr]=cells[idx] if idx<len(cells) else "—"
                        except: pass
                    if ann: d["annual_perf"]=ann
                    break
    return d


@st.cache_data(ttl=3600, show_spinner=False)
def fetch_fondidoc(isin: str) -> dict:
    if not isin: return {}
    # Build search URL from ISIN
    search = f"https://www.fondidoc.it/en/search?q={isin}"
    try:
        r = requests.get(_fd_en(search), headers=FD_HDR, timeout=8)
        if r.status_code != 200: return {}
        soup = BeautifulSoup(r.text,"lxml")
        link = soup.find("a", href=re.compile(r'/d/Index/'))
        if not link: return {}
        fund_url = "https://www.fondidoc.it" + link["href"]
        result = {"url": fund_url}
        m = re.search(r'/([A-Z]{2}[A-Z0-9]{10})[_/]', fund_url)
        if m: result["isin"] = m.group(1)
        h_ov = _fd_fetch(fund_url)
        if h_ov: result["overview"] = _fd_overview(h_ov)
        h_an = _fd_fetch(_fd_ana(fund_url))
        if h_an: result["analysis"] = _fd_analysis(h_an)
        return result
    except Exception: return {}


def fetch_all(funds_df: pd.DataFrame, progress_cb=None) -> dict:
    results = {}; total = len(funds_df); done = 0
    def _one(row):
        isin = row.get("isin",""); nome = row["nome"]; r = {}
        if isin:
            r["morningstar"] = fetch_morningstar(isin)
            r["fondidoc"]    = fetch_fondidoc(isin)
        return nome, r
    with ThreadPoolExecutor(max_workers=6) as pool:
        futures = {pool.submit(_one, row): row["nome"] for _,row in funds_df.iterrows()}
        for fut in as_completed(futures):
            try: nome, data = fut.result(); results[nome] = data
            except: results[futures[fut]] = {}
            done += 1
            if progress_cb: progress_cb(done/total)
    return results


# ════════════════════════════════════════════════════════════
# PORTFOLIO CONSTRUCTION
# ════════════════════════════════════════════════════════════

def _extract_market_view(text: str) -> dict:
    """
    Analisi strutturata del documento di contesto mercati.
    Cerca sezioni 'Contest View', 'Asset Allocation', 'View' ecc.
    Restituisce segnali per scoring + dati per il Quadro di Mercato nel PDF.
    """
    if not text:
        return {"signals":{"eq":0,"bd":0,"regions":[],"risk":"neutral"},
                "asset_views":[],"geo_views":[],"themes":[],"summary":"","risk":"neutral"}
    t = text.lower()

    # ── cerca la sezione "contest view" / "market view" ─────
    # Isola la porzione rilevante del documento se esiste
    VIEW_MARKERS = ["contest view","market view","view di mercato","asset allocation view",
                    "investment view","our view","tactical view","posizionamento"]
    view_text = text
    for marker in VIEW_MARKERS:
        idx = t.find(marker)
        if idx >= 0:
            # Prendi i 3000 caratteri dopo il marker
            view_text = text[idx:idx+3000]
            break
    vt = view_text.lower()

    # ── asset class: vista +/=/- ─────────────────────────────
    AC_DEF = {
        "Azionario":      (["azionari","equity","azioni","stock"],
                           ["sovrappeso","overweight","positiv","costruttiv","rialzist","preferit","favorit","aumenta"],
                           ["sottopeso","underweight","negativ","ribassist","ridurr","cauto","evitar"]),
        "Obbligazionario":(["obbligazion","bond","reddito fisso","fixed income","governo","governativ","corporate","credito","duration"],
                           ["sovrappeso","overweight","positiv","costruttiv","duration lunga","tassi in calo"],
                           ["sottopeso","underweight","negativ","duration corta","short duration","tassi in rialzo"]),
        "Monetario/Cash": (["monetari","money market","liquidità","liquidita","cash"],
                           ["sovrappeso","overweight","favorit","privilegia"],
                           ["sottopeso","underweight","ridurr"]),
        "Alternativi":    (["alternativ","commodity","commodit","real asset","oro","gold","infrastructure"],
                           ["sovrappeso","overweight","positiv","costruttiv"],
                           ["sottopeso","underweight","negativ"]),
    }
    asset_views = []
    eq_score = 0; bd_score = 0
    for asset, (keys, pos_w, neg_w) in AC_DEF.items():
        pos = neg = 0
        for key in keys:
            for m in re.finditer(re.escape(key), vt):
                ctx = vt[max(0,m.start()-90):m.end()+90]
                pos += sum(1 for w in pos_w if w in ctx)
                neg += sum(1 for w in neg_w if w in ctx)
        if pos == 0 and neg == 0: continue
        if pos > neg:   view, col = "Sovrappeso (+)", "#1A7A4A"
        elif neg > pos: view, col = "Sottopeso (−)",  "#C0392B"
        else:           view, col = "Neutrale (=)",   "#475569"
        asset_views.append({"asset":asset,"view":view,"color":col,"pos":pos,"neg":neg})
        if "Azionario"      in asset: eq_score = pos - neg
        if "Obbligazionario" in asset: bd_score = pos - neg

    # ── preferenze geografiche ───────────────────────────────
    GEO = {
        "Europa":    ["europa","eurozona","eurozone","bce","ecb","dax","euro stoxx"],
        "USA":       ["usa","stati uniti","america","s&p","fed ","federal reserve","wall street"],
        "Emergenti": ["emergenti","emerging market","em ","cina","india","brasile","latam"],
        "Asia":      ["asia ex","asia-pacific","apac"],
        "Giappone":  ["giappone","japan","nikkei","boj","topix"],
        "UK":        ["regno unito","uk ","gran bretagna","ftse"],
    }
    geo_views = []
    regions = []
    for geo, keys in GEO.items():
        cnt = sum(vt.count(k) for k in keys)
        if cnt >= 1:
            regions.append(geo)
            geo_views.append({"region": geo, "cnt": cnt})
    geo_views = sorted(geo_views, key=lambda x: x["cnt"], reverse=True)

    # ── propensione al rischio ────────────────────────────────
    low_w  = ["risk off","risk-off","cautela","prudente","difensiv","quality","qualità","bassa volatil","protettiv"]
    high_w = ["risk on","risk-on","aggressiv","pro-ciclic","high beta","risk appetite","crescita"]
    rl = sum(1 for w in low_w  if w in vt)
    rh = sum(1 for w in high_w if w in vt)
    risk = "low" if rl > rh else ("high" if rh > rl else "neutral")

    # ── temi di mercato ───────────────────────────────────────
    THEMES = [
        (["inflazion","inflation","cpi","pce"],                         "Inflazione"),
        (["tassi","rate","banca centr","bce","fed "],                   "Banche centrali / Tassi"),
        (["crescita","gdp","pil","recessione","recession","soft land"], "Crescita / Ciclo"),
        (["geopolit","guerra","conflitt","ucraina","russia","taiwan"],  "Rischi geopolitici"),
        (["credito","spread","high yield","investment grade"],          "Credito"),
        (["dollaro","euro/dollar","eur/usd","valuta","forex","fx "],    "Valute"),
        (["tecnologi","ai ","intelligenza artificiale","semi"],         "Tecnologia / AI"),
        (["energia","petrolio","oil","gas","materie prime","commodity"],"Energia / Commodities"),
        (["immobil","reit","real estate"],                              "Immobiliare"),
    ]
    themes = [lbl for keys, lbl in THEMES if any(k in vt for k in keys)]

    # ── sommario: prime frasi significative della sezione view ─
    sum_lines = [l.strip() for l in view_text.split("\n") if len(l.strip()) > 40]
    summary = " ".join(sum_lines[:6])[:1800]

    return {
        "signals":     {"eq": eq_score, "bd": bd_score, "regions": regions, "risk": risk},
        "asset_views": asset_views,
        "geo_views":   geo_views,
        "themes":      themes,
        "summary":     summary,
        "risk":        risk,
    }


def _market_signals(text: str) -> dict:
    """Wrapper per compatibilità — restituisce solo i segnali di scoring."""
    return _extract_market_view(text)["signals"]


def _score_funds(funds_df, signals, mifid_df=None):
    df = funds_df.copy()
    if mifid_df is not None and not mifid_df.empty:
        key = "isin" if ("isin" in mifid_df.columns and "isin" in df.columns) else "nome" if "nome" in mifid_df.columns else None
        if key:
            df = df.merge(mifid_df[[key,"mifid"]].rename(columns={"mifid":"_m2"}), on=key, how="left")
            df["mifid"] = df["_m2"].fillna(df["mifid"]).astype(int); df = df.drop(columns=["_m2"])
    scores = []
    for _, row in df.iterrows():
        s = 0.0; mc = row.get("macro_cat","Altro")
        n = str(row.get("nome","")).lower(); c = str(row.get("categoria","")).lower()
        if mc=="Azionari":              s += signals["eq"]*2
        elif mc=="Obbligazionari":      s += signals["bd"]*2
        elif mc=="Bilanciati/Flessibili": s += (signals["eq"]+signals["bd"])*0.5
        for reg in signals.get("regions",[]):
            if reg.lower() in n or reg.lower() in c: s += 1.5
        risk = signals.get("risk","neutral")
        if risk=="low"  and mc in ("Obbligazionari","Monetari","Alternativi"): s += 1
        if risk=="high" and mc in ("Azionari","Bilanciati/Flessibili"): s += 1
        mf = int(row.get("mifid",4))
        if 2<=mf<=5: s += 0.3
        scores.append(s)
    df["_score"] = scores
    return df.sort_values("_score", ascending=False).reset_index(drop=True)


def construct_portfolios(funds_df, market_text, mifid_df=None, api_key="", libero_override=None):
    if api_key and _HAS_ANTHROPIC:
        try:
            result = _construct_ai(funds_df, market_text, mifid_df, api_key)
        except Exception as e:
            st.warning(f"AI fallita ({e}) — uso rules-based.")
            result = _construct_rules(funds_df, market_text, mifid_df)
    else:
        result = _construct_rules(funds_df, market_text, mifid_df)
    # Sovrascrive il Portafoglio Libero con la selezione manuale dell'utente
    if libero_override and libero_override.get("funds"):
        result["libero"] = libero_override
    return result


def _construct_rules(funds_df, market_text, mifid_df=None):
    signals = _market_signals(market_text)
    scored  = _score_funds(funds_df, signals, mifid_df)

    def _fe(r):
        d = {"nome":r["nome"],"isin":r.get("isin",""),
             "macro_cat":r.get("macro_cat","Altro"),"color":r.get("color","#94A3B8")}
        if "unp"  in r.index and r["unp"]  is not None and str(r["unp"])  not in ("nan","None",""):
            d["unp"]  = r["unp"]
        if "iunp" in r.index and r["iunp"] is not None and str(r["iunp"]) not in ("nan","None",""):
            d["iunp"] = r["iunp"]
        return d

    # ── Articolato: 12 fondi, bilanciati per macro-categoria ──
    cats = [c for c in scored["macro_cat"].unique() if c != "Altro"]
    if "Altro" in scored["macro_cat"].values: cats.append("Altro")
    art = []; seen_art = set()
    per_cat = max(1, 12 // max(len(cats), 1))
    # Prima passa: prendi i migliori per categoria
    for cat in cats:
        for _, r in scored[scored["macro_cat"] == cat].head(per_cat).iterrows():
            if len(art) >= 12: break
            if r["nome"] not in seen_art:
                art.append({**_fe(r), "peso": 1.0}); seen_art.add(r["nome"])
    # Seconda passa: riempi fino a 12 con i migliori rimasti
    for _, r in scored.iterrows():
        if len(art) >= 12: break
        if r["nome"] not in seen_art:
            art.append({**_fe(r), "peso": 1.0}); seen_art.add(r["nome"])
    art = _norm_w(art)

    # ── Short: top 6 per score ────────────────────────────────
    short = _norm_w([{**_fe(r), "peso": 1.0} for _, r in scored.head(6).iterrows()])

    # ── Libero: top 25 per score, pesi uguali ─────────────────
    # (non tutti i fondi: sarebbe inutile e il PDF esploderebbe)
    MAX_LIBERO = 25
    libero = _norm_w([{**_fe(r), "peso": 1.0} for _, r in scored.head(MAX_LIBERO).iterrows()])

    eq_label = ("Sovrappeso azionario" if signals["eq"] > 0
                else "Sovrappeso obbligazionario" if signals["bd"] > 0 else "Allocazione bilanciata")
    univ = len(funds_df)
    return {
        "articolato": {
            "funds": art,
            "rationale": (f"Portafoglio diversificato su {len(art)} fondi — {eq_label}. "
                          f"Regioni preferite: {', '.join(signals['regions']) or 'globale'}. "
                          f"Selezionati dall'universo di {univ} fondi.")
        },
        "short": {
            "funds": short,
            "rationale": (f"Alta convinzione: i {len(short)} fondi più allineati alla view di mercato "
                          f"({eq_label.lower()}).")
        },
        "libero": {
            "funds": libero,
            "rationale": (f"Top {len(libero)} fondi per punteggio di coerenza con la view di mercato, "
                          f"pesi uguali — personalizzabile. Universo totale: {univ} fondi.")
        },
    }


def _construct_ai(funds_df, market_text, mifid_df, api_key):
    client = _anthropic_mod.Anthropic(api_key=api_key)
    records = [{"nome":r["nome"],"isin":r.get("isin",""),"categoria":r.get("categoria",""),
                "macro_cat":r.get("macro_cat",""),"mifid":int(r.get("mifid",4)),
                "descrizione":str(r.get("descrizione",""))[:150]}
               for _,r in funds_df.iterrows()]
    prompt = (f"Sei un gestore di portafogli professionista Azimut.\n\n"
              f"CONTESTO MERCATO:\n{market_text[:3000]}\n\n"
              f"UNIVERSO FONDI ({len(records)}):\n{json.dumps(records,ensure_ascii=False)[:4000]}\n\n"
              f"Costruisci 3 portafogli:\n"
              f"1. ARTICOLATO: 8-12 fondi diversificati, pesi variabili\n"
              f"2. SHORT: 4-6 fondi, alta convinzione\n"
              f"3. LIBERO: tutti i fondi, pesi proporzionali alla view\n"
              f"Pesi somma 100 per ciascun portafoglio.\n"
              f"Rispondi SOLO in JSON:\n"
              f'{{"articolato":{{"funds":[{{"nome":"...","isin":"...","peso":15.0}}],"rationale":"..."}},'
              f'"short":{{...}},"libero":{{...}}}}')
    resp = client.messages.create(model="claude-opus-4-5", max_tokens=4000,
                                   messages=[{"role":"user","content":prompt}])
    text = resp.content[0].text.strip()
    m = re.search(r'\{[\s\S]*\}', text)
    if not m: raise ValueError("No JSON in response")
    data = json.loads(m.group())
    lookup = {r["nome"]:r for _,r in funds_df.iterrows()}
    for k in ["articolato","short","libero"]:
        if k not in data: continue
        enriched = []
        for f in data[k].get("funds",[]):
            fd = lookup.get(f["nome"],{})
            enriched.append({"nome":f["nome"],"isin":f.get("isin",fd.get("isin","")),
                              "peso":float(f.get("peso",0)),
                              "macro_cat":fd.get("macro_cat","Altro"),"color":fd.get("color","#94A3B8")})
        data[k]["funds"] = _norm_w(enriched)
    return data


# ════════════════════════════════════════════════════════════
# METRICS
# ════════════════════════════════════════════════════════════

def calc_metrics(ptf_funds: list, fund_data: dict, schede_alloc: dict = None) -> dict:
    keys = ["ytd","perf_1y","perf_3y","perf_5y","vol_1y","vol_3y","var_1y","sharpe_3y","neg_vol_1y","sortino_1y"]
    tot = {k:0.0 for k in keys}; wt = {k:0.0 for k in keys}; macro_alloc = {}
    unp_tot = 0.0; unp_wt = 0.0
    iunp_tot = 0.0; iunp_wt = 0.0
    for f in ptf_funds:
        w    = f["peso"] / 100.0
        nome = f["nome"]
        isin = f.get("isin","")
        ana  = fund_data.get(nome,{}).get("fondidoc",{}).get("analysis",{})
        for k in keys:
            try:
                num = float(str(ana.get(k,"")).replace("%","").replace(",",".").strip())
                tot[k] += num*w; wt[k] += w
            except: pass
        # UNP / IUNP ponderati
        unp_v  = f.get("unp");  iunp_v = f.get("iunp")
        if unp_v  is not None:
            try: unp_tot  += float(unp_v)  * w; unp_wt  += w
            except: pass
        if iunp_v is not None:
            try: iunp_tot += float(iunp_v) * w; iunp_wt += w
            except: pass
        # Asset allocation: usa scheda prodotto se disponibile, altrimenti macro_cat
        if schede_alloc and isin and isin in schede_alloc:
            for cat, pct in schede_alloc[isin].items():
                macro_alloc[cat] = macro_alloc.get(cat, 0.0) + f["peso"] * pct / 100.0
            allocated = sum(schede_alloc[isin].values())
            if allocated < 98:
                macro_alloc["Altro"] = macro_alloc.get("Altro", 0.0) + f["peso"] * (100 - allocated) / 100.0
        else:
            mc = f.get("macro_cat","Altro")
            macro_alloc[mc] = macro_alloc.get(mc, 0.0) + f["peso"]
    result = {k: (f"{tot[k]/wt[k]:+.2f}%" if wt[k]>0.01 else "N/D") for k in keys}
    result["macro_alloc"] = macro_alloc
    result["unp_w"]  = round(unp_tot  / unp_wt  * 100, 3) if unp_wt  > 0.01 else None
    result["iunp_w"] = round(iunp_tot / iunp_wt * 100, 3) if iunp_wt > 0.01 else None
    return result


# ════════════════════════════════════════════════════════════
# CHARTS
# ════════════════════════════════════════════════════════════

def _pie_funds(ptf_funds, title):
    # Ordina per peso decrescente per il grafico
    ptf_sorted = sorted(ptf_funds, key=lambda f: f["peso"], reverse=True)
    names   = [f["nome"][:30]+("…" if len(f["nome"])>30 else "") for f in ptf_sorted]
    weights = [f["peso"] for f in ptf_sorted]
    colors  = [f.get("color","#94A3B8") for f in ptf_sorted]
    n = len(ptf_sorted)

    fig,(ax1,ax2) = plt.subplots(1,2,figsize=(11,5),gridspec_kw={"width_ratios":[1.3,1]})
    wedges,_,ats = ax1.pie(weights,colors=colors,
                            autopct=lambda p:f"{p:.1f}%" if p>3 else "",
                            pctdistance=0.72,
                            wedgeprops=dict(width=0.58,edgecolor="white",linewidth=1.5),
                            startangle=90)
    for at in ats: at.set_fontsize(7); at.set_color("white"); at.set_fontweight("bold")
    ax1.text(0,0,title[:8],ha="center",va="center",fontsize=10,fontweight="bold",color="#0D1B2A")
    ax2.axis("off")

    # Limita la legenda ai top 18 per peso; raggruppa il resto
    MAX_LEG = 18
    if n > MAX_LEG:
        top_handles = [mpatches.Patch(color=colors[i]) for i in range(MAX_LEG)]
        top_labels  = [f"{names[i]}  {weights[i]:.1f}%" for i in range(MAX_LEG)]
        altri_w = sum(weights[MAX_LEG:])
        top_handles.append(mpatches.Patch(color="#CBD5E1"))
        top_labels.append(f"altri {n-MAX_LEG} fondi  {altri_w:.1f}%")
    else:
        top_handles = [mpatches.Patch(color=colors[i]) for i in range(n)]
        top_labels  = [f"{names[i]}  {weights[i]:.1f}%" for i in range(n)]

    ax2.legend(top_handles, top_labels,
               loc="center left", frameon=False,
               fontsize=7.5, labelspacing=0.85, handlelength=1.2)
    fig.patch.set_facecolor("#FFFFFF"); plt.tight_layout(pad=1.5)
    buf=io.BytesIO(); fig.savefig(buf,format="png",dpi=150,bbox_inches="tight",facecolor="white")
    plt.close(fig); buf.seek(0); return buf


def _pie_macro(macro_alloc):
    filtered = {k:v for k,v in macro_alloc.items() if v>0.5}
    if not filtered: return None
    keys=list(filtered.keys()); vals=list(filtered.values())
    colors=[MACRO_COLORS.get(k,"#94A3B8") for k in keys]
    fig,(ax1,ax2)=plt.subplots(1,2,figsize=(10,4),gridspec_kw={"width_ratios":[1.2,1]})
    wedges,_,ats=ax1.pie(vals,colors=colors,autopct=lambda p:f"{p:.1f}%" if p>=5 else "",
                          pctdistance=0.70,wedgeprops=dict(width=0.58,edgecolor="white",linewidth=2.5),startangle=90)
    for at in ats: at.set_fontsize(9.5); at.set_color("white"); at.set_fontweight("bold")
    ax1.text(0,0,"Asset\nAlloc.",ha="center",va="center",fontsize=10,fontweight="bold",color="#0D1B2A")
    ax2.axis("off")
    handles = [mpatches.Patch(color=colors[i]) for i in range(len(keys))]
    labels  = [f"{keys[i]}  {vals[i]:.1f}%" for i in range(len(keys))]
    ax2.legend(handles, labels,
               loc="center left", frameon=False, fontsize=9, labelspacing=1.1, handlelength=1.3)
    fig.patch.set_facecolor("#FFFFFF"); plt.tight_layout(pad=1.2)
    buf=io.BytesIO(); fig.savefig(buf,format="png",dpi=150,bbox_inches="tight",facecolor="white")
    plt.close(fig); buf.seek(0); return buf


# ════════════════════════════════════════════════════════════
# PDF GENERATOR
# ════════════════════════════════════════════════════════════

def generate_pdf(portfolios, funds_df, fund_data, market_text, fund_sheets=None, schede_alloc=None, market_view=None, advisor_note=""):
    # ── STILI ────────────────────────────────────────────────
    NAV="#0D1B2A"; GLD="#C9A84C"; MED="#64748B"; LGT="#94A3B8"; SLT="#1E293B"
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            leftMargin=1.8*cm, rightMargin=1.8*cm,
                            topMargin=1.8*cm,  bottomMargin=1.8*cm)
    ss = getSampleStyleSheet()
    def S(n,**kw): return ParagraphStyle(n, parent=ss["Normal"], **kw)
    T1  = S("g_T1", fontName="Helvetica-Bold",    fontSize=20,  textColor=rl_colors.HexColor(NAV), spaceAfter=3,  leading=24)
    T2  = S("g_T2", fontName="Helvetica-Bold",    fontSize=12,  textColor=rl_colors.HexColor(NAV), spaceBefore=10,spaceAfter=4)
    EY  = S("g_EY", fontName="Helvetica",         fontSize=7,   textColor=rl_colors.HexColor(LGT), spaceAfter=2,  letterSpacing=2)
    SU  = S("g_SU", fontName="Helvetica",         fontSize=9,   textColor=rl_colors.HexColor(MED), spaceAfter=3)
    BD  = S("g_BD", fontName="Helvetica",         fontSize=8.5, textColor=rl_colors.HexColor(SLT), leading=13)
    SM  = S("g_SM", fontName="Helvetica",         fontSize=7,   textColor=rl_colors.HexColor(SLT), leading=10)
    IT  = S("g_IT", fontName="Helvetica-Oblique", fontSize=8,   textColor=rl_colors.HexColor("#475569"), leading=12)
    HDR = S("g_HR", fontName="Helvetica-Bold",    fontSize=7,   textColor=rl_colors.white,          leading=10)
    FS  = S("g_FS", fontName="Helvetica-Bold",    fontSize=8.5, textColor=rl_colors.HexColor(NAV),  spaceAfter=1, leading=12)
    FK  = S("g_FK", fontName="Helvetica",         fontSize=7,   textColor=rl_colors.HexColor(MED),  spaceAfter=1, leading=10)
    NOTE= S("g_NT", fontName="Helvetica-Oblique", fontSize=6,   textColor=rl_colors.HexColor(LGT),  leading=8)
    FT  = S("g_FT", fontName="Helvetica-Oblique", fontSize=6,   textColor=rl_colors.HexColor(LGT),  leading=9)
    LK  = S("g_LK", fontName="Helvetica",         fontSize=7,   textColor=rl_colors.HexColor("#1B4FBB"), spaceAfter=1)

    story = []
    W = 17.4*cm   # larghezza utile

    # ── helper ───────────────────────────────────────────────
    def accent_bar(color=NAV):
        return Table([[""]], colWidths=[W], rowHeights=[8],
                     style=TableStyle([("BACKGROUND",(0,0),(-1,-1),rl_colors.HexColor(color)),
                                       ("LINEBELOW",(0,0),(-1,-1),3,rl_colors.HexColor(GLD))]))
    def color_bar(color):
        return Table([[""]], colWidths=[W], rowHeights=[3],
                     style=TableStyle([("BACKGROUND",(0,0),(-1,-1),rl_colors.HexColor(color))]))
    def light_rule():
        return HRFlowable(width="100%",thickness=0.5,color=rl_colors.HexColor("#E2E8F0"),spaceAfter=6)
    def kpi_cell(v, l):
        return Paragraph(f'<font size="13"><b>{v}</b></font><br/>'
                         f'<font size="7" color="{MED}">{l}</font>', BD)
    def pv(val):
        try:
            n = float(str(val).replace("%","").replace(",",".").replace("+","").strip())
            c = "#1A7A4A" if n>0 else ("#C0392B" if n<0 else "#475569")
            sign = "+" if n>0 else ""
            return Paragraph(f'<font color="{c}"><b>{sign}{n:.1f}%</b></font>', SM)
        except: return Paragraph(str(val) if val else "—", SM)
    def txt(v):
        s = str(v) if v else ""
        if not s or s in ("—","nan","None","N/D"):
            return Paragraph(f'<font color="{LGT}">—</font>', SM)
        return Paragraph(s, SM)
    def kv(v):  # KPI value: mostra n.d. più discreto se mancante
        if not v or v in ("N/D","—","nan","None"):
            return f'<font color="{LGT}" size="9">n.d.</font>'
        return f'<font size="13"><b>{v}</b></font>'

    # ════════════════════════════════════════════════════════
    # PAG 1 — COPERTINA
    # ════════════════════════════════════════════════════════
    n_f    = len(funds_df)
    macros = funds_df["macro_cat"].value_counts()
    top_m  = macros.index[0] if len(macros)>0 else "—"
    n_art  = len(portfolios.get("articolato",{}).get("funds",[]))
    n_sh   = len(portfolios.get("short",{}).get("funds",[]))
    n_lib  = len(portfolios.get("libero",{}).get("funds",[]))

    story += [
        accent_bar(), Spacer(1,14),
        Paragraph("AZIMUT INVESTMENTS  ·  PORTFOLIO BUILDER", EY), Spacer(1,4),
        Paragraph("Proposta di Portafoglio", T1),
        Paragraph(datetime.date.today().strftime("%d %B %Y"), SU),
        HRFlowable(width="100%",thickness=0.8,color=rl_colors.HexColor("#E2E8F0"),spaceAfter=10),
    ]

    # KPI copertina
    kpi_cov = Table([[kpi_cell(str(n_f),"Fondi universo"),
                      kpi_cell(str(len(macros)),"Macro categorie"),
                      kpi_cell(top_m[:15],"Categoria principale"),
                      kpi_cell(datetime.date.today().strftime("%m/%Y"),"Data report")]],
                    colWidths=[W/4]*4)
    kpi_cov.setStyle(TableStyle([
        ("BOX",(0,0),(-1,-1),0.8,rl_colors.HexColor("#E2E8F0")),
        ("INNERGRID",(0,0),(-1,-1),0.8,rl_colors.HexColor("#E2E8F0")),
        ("BACKGROUND",(0,0),(-1,-1),rl_colors.HexColor("#F8FAFC")),
        ("PADDING",(0,0),(-1,-1),10),("ALIGN",(0,0),(-1,-1),"CENTER"),("VALIGN",(0,0),(-1,-1),"MIDDLE"),
    ]))
    story += [kpi_cov, Spacer(1,10)]

    # Sommario 3 portafogli
    ptf_sum_data = [
        [Paragraph("<b>Portafoglio</b>",HDR), Paragraph("<b>Fondi</b>",HDR), Paragraph("<b>Strategia</b>",HDR)],
        [Paragraph("Articolato",SM), Paragraph(str(n_art),SM), Paragraph("Diversificato per macro-categoria",SM)],
        [Paragraph("Short",SM),      Paragraph(str(n_sh), SM), Paragraph("Alta convinzione — top score",SM)],
        [Paragraph("Libero",SM),     Paragraph(str(n_lib),SM), Paragraph("Selezione manuale con pesi personalizzati",SM)],
    ]
    ptf_sum = Table(ptf_sum_data, colWidths=[3.5*cm,2*cm,11.9*cm])
    ptf_sum.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),rl_colors.HexColor(NAV)),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[rl_colors.white,rl_colors.HexColor("#F8FAFC")]),
        ("FONTSIZE",(0,0),(-1,-1),7),("PADDING",(0,0),(-1,-1),5),
        ("LINEBELOW",(0,0),(-1,-1),0.4,rl_colors.HexColor("#E2E8F0")),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
    ]))
    story += [ptf_sum, Spacer(1,10)]

    if market_text and market_text.strip():
        story += [Paragraph("Contesto di Mercato", T2), Spacer(1,4)]
        # Trova la sezione contest view (o usa tutto il testo)
        _mt = market_text.strip()
        _sec = _mt
        for _mk in ["contest view","market view","view di mercato","posizionamento tattico",
                    "posizionamento","asset allocation view","investment view"]:
            _ix = _mt.lower().find(_mk)
            if _ix >= 0:
                _sec = _mt[_ix:_ix+2400].strip()
                break
        else:
            _sec = _mt[:2400]
        # Formatta in paragrafi leggibili
        _paras = [p.strip() for p in re.split(r'\n{1,}', _sec) if len(p.strip()) > 12]
        for _p in _paras[:12]:
            story.append(Paragraph(
                _p[:350].replace("&","&amp;").replace("<","&lt;").replace(">","&gt;"), IT))
            story.append(Spacer(1,3))
        story.append(Spacer(1,6))

    if advisor_note and advisor_note.strip():
        note_style = S("g_note_box", fontName="Helvetica", fontSize=8,
                       textColor=rl_colors.HexColor("#1E293B"), leading=12,
                       leftIndent=10, rightIndent=10, spaceBefore=4, spaceAfter=4)
        note_lbl   = S("g_note_lbl", fontName="Helvetica-Bold", fontSize=7,
                       textColor=rl_colors.HexColor("#7C3AED"), leading=10,
                       leftIndent=10, spaceAfter=2)
        note_rows = [[Paragraph("✏️  NOTE DEL CONSULENTE", note_lbl)],
                     [Paragraph(advisor_note.strip().replace("\n","<br/>"), note_style)]]
        note_tbl  = Table(note_rows, colWidths=[W])
        note_tbl.setStyle(TableStyle([
            ("BOX",    (0,0),(-1,-1), 1.2, rl_colors.HexColor("#7C3AED")),
            ("LINEABOVE",(0,0),(-1,0), 3,  rl_colors.HexColor("#7C3AED")),
            ("BACKGROUND",(0,0),(-1,-1), rl_colors.HexColor("#F5F3FF")),
            ("TOPPADDING",(0,0),(-1,-1), 6),
            ("BOTTOMPADDING",(0,0),(-1,-1), 6),
        ]))
        story += [note_tbl, Spacer(1,6)]

    story.append(PageBreak())

    # ════════════════════════════════════════════════════════
    # PAG 2 — QUADRO DI MERCATO (sempre se c'è testo mercato)
    # ════════════════════════════════════════════════════════
    mv = market_view or {}
    if market_text and market_text.strip():
        story += [
            accent_bar(), Spacer(1,8),
            Paragraph("AZIMUT INVESTMENTS  ·  PORTFOLIO BUILDER", EY), Spacer(1,3),
            Paragraph("Quadro di Mercato", T1),
            Paragraph(f"Analisi del documento di contesto  ·  {datetime.date.today().strftime('%d %B %Y')}", SU),
            color_bar("#C9A84C"), Spacer(1,10),
        ]

        from reportlab.platypus import KeepInFrame
        left_parts  = []
        right_parts = []

        # ── COLONNA SINISTRA: testo completo del documento ──────
        # Trova e mostra la sezione "contest view" (o tutto il testo)
        _full = market_text.strip()
        _mv_sec = _full
        for _mk2 in ["contest view","market view","view di mercato","posizionamento tattico",
                     "posizionamento","asset allocation view","investment view"]:
            _ix2 = _full.lower().find(_mk2)
            if _ix2 >= 0:
                _mv_sec = _full[_ix2:_ix2+4500].strip()
                break
        else:
            _mv_sec = _full[:4500]

        left_parts.append(Paragraph("<b>Documento di Contesto</b>", T2))
        left_parts.append(Spacer(1,6))
        _paras2 = [p.strip() for p in re.split(r'\n{1,}', _mv_sec) if len(p.strip()) > 12]
        for _p2 in _paras2[:20]:
            left_parts.append(Paragraph(
                _p2[:400].replace("&","&amp;").replace("<","&lt;").replace(">","&gt;"), SM))
            left_parts.append(Spacer(1,4))

        # ── COLONNA DESTRA: analisi strutturata ──────────────────
        # Posizionamento asset class
        if mv.get("asset_views"):
            right_parts.append(Paragraph("<b>Posizionamento Asset Class</b>", T2))
            right_parts.append(Spacer(1,4))
            ac_hdr = [Paragraph("<b>Asset Class</b>",HDR),
                      Paragraph("<b>Vista</b>",HDR),
                      Paragraph("<b>Forza</b>",HDR)]
            ac_rows = [ac_hdr]
            for av in mv["asset_views"]:
                dots = min(av.get("pos",0) + av.get("neg",0), 5)
                sig  = "●"*dots + "○"*(5-dots)
                ac_rows.append([
                    Paragraph(av["asset"], SM),
                    Paragraph(f'<font color="{av["color"]}"><b>{av["view"]}</b></font>', SM),
                    Paragraph(f'<font color="{av["color"]}">{sig}</font>', SM),
                ])
            ac_tbl = Table(ac_rows, colWidths=[2.8*cm, 3.0*cm, 1.6*cm])
            ac_tbl.setStyle(TableStyle([
                ("BACKGROUND",(0,0),(-1,0),rl_colors.HexColor(NAV)),
                ("ROWBACKGROUNDS",(0,1),(-1,-1),[rl_colors.white,rl_colors.HexColor("#F8FAFC")]),
                ("FONTSIZE",(0,0),(-1,-1),7),("PADDING",(0,0),(-1,-1),4),
                ("LINEBELOW",(0,0),(-1,-1),0.4,rl_colors.HexColor("#E2E8F0")),
                ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
            ]))
            right_parts.append(ac_tbl)
            right_parts.append(Spacer(1,10))

        # Propensione rischio
        risk_lbl = {"low":"Risk-Off / Difensivo","high":"Risk-On / Aggressivo","neutral":"Bilanciato / Neutrale"}
        risk_col = {"low":"#C0392B","high":"#1A7A4A","neutral":"#475569"}
        risk_val = mv.get("risk","neutral")
        right_parts.append(Paragraph("<b>Propensione al Rischio</b>", T2))
        right_parts.append(Spacer(1,4))
        risk_box = Table([[Paragraph(
            f'<font color="{risk_col[risk_val]}" size="11"><b>{risk_lbl[risk_val]}</b></font>', BD)]],
            colWidths=[7.2*cm])
        risk_box.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,-1),rl_colors.HexColor("#F8FAFC")),
            ("BOX",(0,0),(-1,-1),2,rl_colors.HexColor(risk_col[risk_val])),
            ("PADDING",(0,0),(-1,-1),9),("ALIGN",(0,0),(-1,-1),"CENTER"),
        ]))
        right_parts += [risk_box, Spacer(1,10)]

        # Temi identificati
        if mv.get("themes"):
            right_parts.append(Paragraph("<b>Temi Chiave</b>", T2))
            right_parts.append(Spacer(1,4))
            for th in mv["themes"]:
                right_parts.append(Paragraph(f"▸  {th}", SM))
                right_parts.append(Spacer(1,3))
            right_parts.append(Spacer(1,8))

        # Aree geografiche
        if mv.get("geo_views"):
            right_parts.append(Paragraph("<b>Regioni citate</b>", T2))
            right_parts.append(Spacer(1,4))
            geo_rows = [[Paragraph("<b>Regione</b>",HDR), Paragraph("<b>Rilevanza</b>",HDR)]]
            for gv_item in mv["geo_views"][:6]:
                geo_rows.append([
                    Paragraph(gv_item["region"], SM),
                    Paragraph("●" * min(gv_item["cnt"],5) + "○"*(5-min(gv_item["cnt"],5)), SM),
                ])
            geo_tbl = Table(geo_rows, colWidths=[3*cm, 4.2*cm])
            geo_tbl.setStyle(TableStyle([
                ("BACKGROUND",(0,0),(-1,0),rl_colors.HexColor(NAV)),
                ("ROWBACKGROUNDS",(0,1),(-1,-1),[rl_colors.white,rl_colors.HexColor("#F8FAFC")]),
                ("FONTSIZE",(0,0),(-1,-1),7),("PADDING",(0,0),(-1,-1),4),
                ("LINEBELOW",(0,0),(-1,-1),0.4,rl_colors.HexColor("#E2E8F0")),
            ]))
            right_parts.append(geo_tbl)

        # Layout a due colonne (sinistra più larga per il testo)
        left_frame  = KeepInFrame(9.8*cm, 24*cm, left_parts,  mode="shrink")
        right_frame = KeepInFrame(7.2*cm, 24*cm, right_parts, mode="shrink")
        two_col = Table([[left_frame, right_frame]], colWidths=[10.0*cm, 7.4*cm])
        two_col.setStyle(TableStyle([
            ("VALIGN",(0,0),(-1,-1),"TOP"),
            ("LEFTPADDING",(1,0),(1,0),12),
            ("LINEAFTER",(0,0),(0,-1),0.5,rl_colors.HexColor("#E2E8F0")),
        ]))
        story.append(two_col)
        story.append(PageBreak())

    # ════════════════════════════════════════════════════════
    # PAG 3-5 (o 2-4 senza quadro) — UN PORTAFOGLIO PER PAGINA
    # tabella fondi + torta asset allocation affiancati
    # ════════════════════════════════════════════════════════
    PTF_COLORS = {"articolato":"#1B4FBB","short":"#065F46","libero":"#7C3AED"}

    for ptf_key, ptf_label, ptf_color in [
        ("articolato","Portafoglio Articolato","#1B4FBB"),
        ("short",     "Portafoglio Short",     "#065F46"),
        ("libero",    "Portafoglio Libero",     "#7C3AED"),
    ]:
        ptf = portfolios.get(ptf_key,{})
        if not ptf or not ptf.get("funds"): continue
        funds_list = ptf["funds"]
        rationale  = ptf.get("rationale","")
        metrics    = calc_metrics(funds_list, fund_data, schede_alloc)

        # Header pagina
        story += [
            accent_bar(), Spacer(1,8),
            Paragraph("AZIMUT INVESTMENTS  ·  PORTFOLIO BUILDER", EY), Spacer(1,3),
            Paragraph(ptf_label, T1),
            Paragraph(f"{len(funds_list)} fondi  ·  {datetime.date.today().strftime('%d %B %Y')}", SU),
            color_bar(ptf_color), Spacer(1,6),
        ]

        # Rationale (max 1 riga compatta)
        if rationale:
            rat = Table([[Paragraph(rationale[:220], IT)]], colWidths=[W])
            rat.setStyle(TableStyle([
                ("BACKGROUND",(0,0),(-1,-1),rl_colors.HexColor("#F8FAFC")),
                ("PADDING",(0,0),(-1,-1),7),
                ("BOX",(0,0),(-1,-1),0.5,rl_colors.HexColor("#E2E8F0")),
                ("LINEBELOW",(0,0),(-1,-1),2,rl_colors.HexColor(ptf_color)),
            ]))
            story += [rat, Spacer(1,7)]

        # Strip KPI — aggiunge UNP/IUNP ponderato se disponibile
        def _kpi(k, lbl):
            v = metrics.get(k,"")
            return Paragraph(f'{kv(v)}<br/><font size="7" color="{MED}">{lbl}</font>', BD)
        has_unp  = metrics.get("unp_w")  is not None
        has_iunp = metrics.get("iunp_w") is not None
        def _unp_kpi(val, lbl, color="#7C3AED"):
            if val is None:
                return Paragraph(f'<font color="{LGT}" size="9">n.d.</font>'
                                 f'<br/><font size="7" color="{MED}">{lbl}</font>', BD)
            return Paragraph(
                f'<font color="{color}" size="13"><b>{val:.3f}%</b></font>'
                f'<br/><font size="7" color="{MED}">{lbl}</font>', BD)

        if has_unp or has_iunp:
            # 9 celle: 7 metriche + UNP + IUNP
            kpi_cells = [_kpi("ytd","YTD"), _kpi("perf_1y","1A"), _kpi("perf_3y","3A"),
                         _kpi("perf_5y","5A"), _kpi("var_1y","VaR 1A"),
                         _kpi("sharpe_3y","Sharpe"), _kpi("neg_vol_1y","Vol.Neg."),
                         _unp_kpi(metrics.get("unp_w"),  "UNP %",  "#7C3AED"),
                         _unp_kpi(metrics.get("iunp_w"), "IUNP %", "#5B21B6")]
            kpi_widths = [W/9]*9
        else:
            kpi_cells  = [_kpi("ytd","YTD"), _kpi("perf_1y","1 Anno"), _kpi("perf_3y","3 Anni"),
                          _kpi("perf_5y","5 Anni"), _kpi("var_1y","VaR 1A"),
                          _kpi("sharpe_3y","Sharpe 3A"), _kpi("neg_vol_1y","Vol.Neg. 1A")]
            kpi_widths = [W/7]*7

        mk = Table([kpi_cells], colWidths=kpi_widths)
        mk.setStyle(TableStyle([
            ("BOX",(0,0),(-1,-1),0.8,rl_colors.HexColor("#E2E8F0")),
            ("INNERGRID",(0,0),(-1,-1),0.8,rl_colors.HexColor("#E2E8F0")),
            ("BACKGROUND",(0,0),(-1,-1),rl_colors.HexColor("#F8FAFC")),
            # Evidenzia UNP/IUNP se presenti
            *([("BACKGROUND",(-2,0),(-1,0),rl_colors.HexColor("#F5F3FF"))] if has_unp or has_iunp else []),
            ("PADDING",(0,0),(-1,-1),8),("ALIGN",(0,0),(-1,-1),"CENTER"),("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ]))
        story += [mk, Spacer(1,8)]

        # Tabella fondi + torta affiancati
        # Determina se mostrare colonna UNP nella tabella fondi
        _show_unp = any(f.get("unp") is not None for f in funds_list)
        if _show_unp:
            hdr = [Paragraph(f"<b>{t}</b>",HDR) for t in
                   ["Fondo","ISIN","Cat.","Peso","UNP","YTD","1A","3A","5A","VaR","Sharpe"]]
        else:
            hdr = [Paragraph(f"<b>{t}</b>",HDR) for t in
                   ["Fondo","ISIN","Categoria","Peso","YTD","1A","3A","5A","VaR 1A","Sharpe"]]
        rows = [hdr]
        for f in funds_list:
            nome = f["nome"]
            ana  = fund_data.get(nome,{}).get("fondidoc",{}).get("analysis",{})
            def gf(k): return ana.get(k,"—")
            unp_v  = f.get("unp")
            unp_s  = f"{unp_v:.3f}%" if unp_v is not None else "—"
            if _show_unp:
                rows.append([
                    txt(nome[:30]),           txt(f.get("isin","—")),
                    txt(f.get("macro_cat","—")[:12]),
                    Paragraph(f"<b>{f['peso']:.1f}%</b>",SM),
                    Paragraph(f'<font color="#7C3AED"><b>{unp_s}</b></font>',SM),
                    pv(gf("ytd")),            pv(gf("perf_1y")),
                    pv(gf("perf_3y")),        pv(gf("perf_5y")),
                    txt(gf("var_1y")),        txt(gf("sharpe_3y")),
                ])
            else:
                rows.append([
                    txt(nome[:36]),       txt(f.get("isin","—")),
                    txt(f.get("macro_cat","—")[:16]),
                    Paragraph(f"<b>{f['peso']:.1f}%</b>",SM),
                    pv(gf("ytd")),        pv(gf("perf_1y")),
                    pv(gf("perf_3y")),    pv(gf("perf_5y")),
                    txt(gf("var_1y")),    txt(gf("sharpe_3y")),
                ])
        _fund_tbl_style = TableStyle([
            ("BACKGROUND",(0,0),(-1,0),rl_colors.HexColor(NAV)),
            ("FONTSIZE",(0,0),(-1,-1),6.8),("PADDING",(0,0),(-1,-1),3),
            ("ROWBACKGROUNDS",(0,1),(-1,-1),[rl_colors.white,rl_colors.HexColor("#F8FAFC")]),
            ("LINEBELOW",(0,0),(-1,-1),0.3,rl_colors.HexColor("#E2E8F0")),
            ("ALIGN",(3,0),(-1,-1),"CENTER"),("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ])
        if _show_unp:
            ft = Table(rows,
                       colWidths=[2.9*cm,1.8*cm,1.5*cm,0.9*cm,0.9*cm,
                                  0.9*cm,0.85*cm,0.85*cm,0.85*cm,0.85*cm,0.85*cm],
                       repeatRows=1)
        else:
            ft = Table(rows,
                       colWidths=[3.5*cm,1.85*cm,1.85*cm,1.0*cm,1.0*cm,
                                  1.0*cm,1.0*cm,1.0*cm,1.0*cm,1.0*cm],
                       repeatRows=1)
        ft.setStyle(_fund_tbl_style)

        mb = _pie_macro(metrics.get("macro_alloc",{}))
        if mb:
            pie_img = RLImage(mb, width=5.8*cm, height=4.2*cm)
            tbl_w = W - 6.0*cm
            if _show_unp:
                ft2 = Table(rows,
                            colWidths=[2.4*cm,1.6*cm,1.2*cm,0.8*cm,0.8*cm,
                                       0.8*cm,0.75*cm,0.75*cm,0.75*cm,0.75*cm,0.75*cm],
                            repeatRows=1)
            else:
                ft2 = Table(rows,
                            colWidths=[3.0*cm,1.6*cm,1.5*cm,0.9*cm,0.9*cm,
                                       0.9*cm,0.9*cm,0.9*cm,0.9*cm,0.9*cm],
                            repeatRows=1)
            ft2.setStyle(_fund_tbl_style)
            side = Table([[ft2, pie_img]], colWidths=[tbl_w, 6.0*cm])
            side.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP"),
                                      ("LEFTPADDING",(1,0),(1,0),6)]))
            story.append(side)
        else:
            story.append(ft)

        story += [Spacer(1,4),
                  Paragraph("◆ Metriche = media ponderata sui fondi con dati FondiDoc. A titolo indicativo.",NOTE),
                  PageBreak()]

    # ════════════════════════════════════════════════════════
    # SCHEDE ANALITICHE — solo Articolato + Short (max ~15 fondi)
    # Formato compatto: 4 righe per fondo, ~6 fondi per pagina
    # ════════════════════════════════════════════════════════
    art_nomi   = {f["nome"] for f in portfolios.get("articolato",{}).get("funds",[])}
    short_nomi = {f["nome"] for f in portfolios.get("short",{}).get("funds",[])}
    key_nomi   = art_nomi | short_nomi
    key_df     = funds_df[funds_df["nome"].isin(key_nomi)].reset_index(drop=True)

    if not key_df.empty:
        story += [
            accent_bar(), Spacer(1,8),
            Paragraph("AZIMUT INVESTMENTS  ·  PORTFOLIO BUILDER", EY), Spacer(1,3),
            Paragraph("Schede Analitiche — Portafogli Articolato + Short", T1),
            Paragraph(f"Fonte: FondiDoc + Morningstar  ·  {datetime.date.today().strftime('%d %B %Y')}", SU),
            light_rule(),
            Paragraph('🔍 <link href="https://www.morningstar.it/it/funds/SecuritySearchResults.aspx">'
                      '<u>Motore di ricerca Morningstar</u></link>', LK),
            Spacer(1,6),
        ]

        for _, fr in key_df.iterrows():
            nome  = fr["nome"];  isin  = fr.get("isin","")
            fd    = fund_data.get(nome,{})
            ov    = fd.get("fondidoc",{}).get("overview",{})
            ana   = fd.get("fondidoc",{}).get("analysis",{})
            ms    = fd.get("morningstar",{})
            mifid = fr.get("mifid","—");  mc = fr.get("macro_cat","")

            def gv(k, src=None):
                v = (src or ana).get(k,""); return v if v else "—"

            def pvs(v):
                try:
                    n = float(str(v).replace("%","").replace(",",".").replace("+","").strip())
                    c = "#1A7A4A" if n>0 else ("#C0392B" if n<0 else "#475569")
                    return f'<font color="{c}"><b>{"+" if n>0 else ""}{n:.1f}%</b></font>'
                except: return str(v) if v and v!="—" else "—"

            # Colore accento (blu=articolato, verde=short, entrambi=blu)
            acc = "#065F46" if nome in short_nomi and nome not in art_nomi else "#1B4FBB"

            srri_s = f"SRRI {gv('srri',ov)}/7" if gv('srri',ov) not in ("—","") else ""
            nav_s  = f"NAV {gv('nav')} €"       if gv('nav')    not in ("—","") else ""
            rat_s  = f"FIDArating {gv('fida_rating',ov)}" if gv('fida_rating',ov) not in ("—","") else ""
            meta   = "  ·  ".join(x for x in [srri_s,rat_s,nav_s] if x) or "—"
            fee_s  = f"Fee gest.: {gv('mgmt_fee',ov)}" if gv('mgmt_fee',ov) not in ("—","") else ""
            ms_s   = f"MS: {ms.get('ms_cat','—')}" if ms.get('ms_cat') else ""

            # UNP / IUNP dal funds_df (già mergiato con catalogo)
            _fr_row = funds_df[funds_df["nome"]==nome]
            unp_val  = _fr_row["unp"].values[0]  if not _fr_row.empty and "unp"  in _fr_row.columns else None
            iunp_val = _fr_row["iunp"].values[0] if not _fr_row.empty and "iunp" in _fr_row.columns else None
            unp_s2   = (f'  ·  UNP: <font color="#7C3AED"><b>{float(unp_val):.3f}%</b></font>'
                        if unp_val is not None and str(unp_val) not in ("nan","None","") else "")
            iunp_s2  = (f'  ·  IUNP: <font color="#5B21B6"><b>{float(iunp_val):.3f}%</b></font>'
                        if iunp_val is not None and str(iunp_val) not in ("nan","None","") else "")

            r1 = f"<b>{nome}</b>"
            r2 = (f"ISIN: <b>{isin or '—'}</b>  ·  {mc}  ·  MIFID: <b>{mifid}/7</b>"
                  + (f"  ·  {meta}" if meta != "—" else "")
                  + unp_s2 + iunp_s2)
            r3 = (f"YTD: {pvs(gv('ytd'))}  ·  "
                  f"1A: {pvs(gv('perf_1y'))}  ·  "
                  f"3A: {pvs(gv('perf_3y'))}  ·  "
                  f"5A: {pvs(gv('perf_5y'))}  ·  "
                  f"Vol.1A: <b>{gv('vol_1y')}</b>  ·  "
                  f"Vol.Neg.: <b>{gv('neg_vol_1y')}</b>  ·  "
                  f"Sharpe 3A: <b>{gv('sharpe_3y')}</b>  ·  "
                  f"Sortino 1A: <b>{gv('sortino_1y')}</b>")
            det_parts = [x for x in [
                f"Avvio: {gv('start_date',ov)}", fee_s,
                f"Fee perf.: {gv('perf_fee',ov)}" if gv('perf_fee',ov) not in ("—","") else "",
                f"Cat. Assog.: {gv('cat_assog',ov)}", ms_s,
            ] if x]
            r4 = "  ·  ".join(det_parts) or "—"

            card = Table(
                [[Paragraph(r1,FS)],[Paragraph(r2,FK)],
                 [Paragraph(r3,SM)],[Paragraph(r4,FK)]],
                colWidths=[W])
            card.setStyle(TableStyle([
                ("BOX",(0,0),(-1,-1),0.8,rl_colors.HexColor("#E2E8F0")),
                ("LEFTPADDING",(0,0),(-1,-1),8),("RIGHTPADDING",(0,0),(-1,-1),8),
                ("TOPPADDING",(0,0),(0,0),7),("BOTTOMPADDING",(0,-1),(-1,-1),7),
                ("TOPPADDING",(0,1),(-1,-1),2),("BOTTOMPADDING",(0,0),(-1,-2),2),
                ("LINEABOVE",(0,0),(-1,0),2,rl_colors.HexColor(acc)),
                ("BACKGROUND",(0,0),(-1,-1),rl_colors.HexColor("#FAFBFC")),
            ]))
            story.append(KeepTogether([card, Spacer(1,5)]))

    # ── DISCLAIMER ──────────────────────────────────────────
    story += [light_rule(),
              Paragraph("Documento generato automaticamente a scopo illustrativo. Dati da FIDA FondiDoc e Morningstar. "
                        "I portafogli sono costruiti tramite analisi rules-based/AI del contesto di mercato e non "
                        "costituiscono offerta o consulenza di investimento. Rendimenti passati non garantiscono "
                        "risultati futuri. © Azimut Group — uso interno.", FT)]

    doc.build(story)
    main_pdf = buf.getvalue()

    if fund_sheets and _HAS_PYPDF:
        try:
            writer = PdfWriter()
            for pdf_b in [main_pdf] + fund_sheets:
                for page in PdfReader(io.BytesIO(pdf_b)).pages:
                    writer.add_page(page)
            out = io.BytesIO(); writer.write(out); return out.getvalue()
        except Exception: pass

    return main_pdf


# ════════════════════════════════════════════════════════════
# STREAMLIT MAIN
# ════════════════════════════════════════════════════════════

def main():
    # ── SIDEBAR ─────────────────────────────────────────────
    with st.sidebar:
        st.markdown("""<div style='padding:1.4rem 0 .8rem 0;'>
          <div style='font-size:.6rem;letter-spacing:.22em;color:#3a5a78;text-transform:uppercase;font-weight:700;'>Portfolio Builder</div>
          <div style='font-family:"Cormorant Garamond",serif;font-size:1.6rem;color:#dde8f5;font-weight:700;margin-top:4px;line-height:1.2;'>AI-Assisted<br>Portfolio</div>
          <div style='width:32px;height:3px;background:#C9A84C;border-radius:2px;margin-top:10px;'></div>
        </div>""", unsafe_allow_html=True)
        st.markdown("---")
        st.markdown("<span style='color:#4a6582;font-size:.68rem;letter-spacing:.12em;text-transform:uppercase;font-weight:600;'>FONDI POTENZIALI (mensile)</span>", unsafe_allow_html=True)
        file_fondi = st.file_uploader("Excel/CSV/PDF — nome, ISIN, categoria, descrizione, MIFID",
                                       type=["xlsx","xls","csv","pdf"], key="u_fondi", label_visibility="collapsed")
        st.markdown("---")
        st.markdown("<span style='color:#4a6582;font-size:.68rem;letter-spacing:.12em;text-transform:uppercase;font-weight:600;'>CONTESTO MERCATI</span>", unsafe_allow_html=True)
        file_mercato = st.file_uploader("TXT/PDF con view di mercato",
                                         type=["txt","pdf","md"], key="u_mercato", label_visibility="collapsed")
        st.markdown("---")
        st.markdown("<span style='color:#4a6582;font-size:.68rem;letter-spacing:.12em;text-transform:uppercase;font-weight:600;'>CATALOGO UNP / IUNP</span>", unsafe_allow_html=True)
        st.caption("Excel/CSV/PDF con ISIN + UNP + IUNP (catalogo prodotti Azimut)")
        file_unp = st.file_uploader("Catalogo UNP",
                                     type=["xlsx","xls","csv","pdf"], key="u_unp",
                                     label_visibility="collapsed")
        st.markdown("---")
        st.markdown("<span style='color:#4a6582;font-size:.68rem;letter-spacing:.12em;text-transform:uppercase;font-weight:600;'>SCHEDE PRODOTTO (PDF allegati)</span>", unsafe_allow_html=True)
        files_schede = st.file_uploader("PDF schede singoli fondi (multipli)",
                                         type=["pdf"], accept_multiple_files=True,
                                         key="u_schede", label_visibility="collapsed")
        st.markdown("---")
        if _HAS_ANTHROPIC:
            st.markdown("<span style='color:#4a6582;font-size:.68rem;letter-spacing:.12em;text-transform:uppercase;font-weight:600;'>CLAUDE API KEY (opzionale)</span>", unsafe_allow_html=True)
            api_key = st.text_input("Claude API key per portafogli AI", type="password",
                                     key="u_api", label_visibility="collapsed", placeholder="sk-ant-…")
        else:
            api_key = ""
            st.caption("💡 Installa `anthropic` per portafogli AI-powered.")
        st.markdown("---")
        st.markdown("<span style='color:#4a6582;font-size:.68rem;letter-spacing:.12em;text-transform:uppercase;font-weight:600;'>NOTE PER IL REPORT</span>", unsafe_allow_html=True)
        st.caption("Testo libero incluso nella copertina del PDF.")
        advisor_note = st.text_area(
            "Note",
            key="u_note",
            label_visibility="collapsed",
            height=130,
            placeholder=(
                "Esempi di cosa scrivere:\n"
                "• Profilo del cliente: es. «Investitore prudente, orizzonte 3 anni»\n"
                "• Obiettivo: es. «Protezione del capitale con crescita moderata»\n"
                "• Vincoli: es. «Escludere settore energia fossile, max 20% azionario USA»\n"
                "• Contesto specifico: es. «Proposta in seguito a ribilanciamento Q2 2026»\n"
                "• Avvertenze: es. «Il cliente ha già esposizione immobiliare diretta»"
            ),
        )

    # ── HEADER ──────────────────────────────────────────────
    st.markdown(f"""<div class="az-header">
      <div class="az-eyebrow">AZIMUT INVESTMENTS · PORTFOLIO BUILDER</div>
      <div class="az-rule"></div>
      <div class="az-title">AI-Assisted Portfolio Construction</div>
      <div class="az-meta">Costruzione portafogli da universo fondi · {datetime.date.today().strftime('%d %B %Y')}</div>
    </div>""", unsafe_allow_html=True)

    if not file_fondi:
        st.info("⬅️ **Carica il file fondi potenziali** per iniziare.")
        st.markdown("""
**Flusso di lavoro:**
1. 📁 **Fondi potenziali** *(mensile)* — Excel/CSV o PDF: gli ISIN vengono estratti automaticamente
2. 📰 **Contesto mercati** — view PDF/TXT con la contest view → guida la costruzione dei portafogli
3. 📊 **Catalogo UNP** *(opzionale)* — Excel/CSV/PDF con ISIN + UNP + IUNP: mostra redditività per fondo e portafoglio
4. 📎 **Schede prodotto** *(opzionale)* — PDF singoli fondi: asset allocation granulare estratta automaticamente
5. ✏️ **Note** *(opzionale)* — testo libero incluso nella copertina (profilo cliente, obiettivi, vincoli)
6. 🚀 **Genera** → 3 portafogli con rendimenti, asset allocation, VaR, Sharpe e **UNP ponderato**
        """)
        return

    # ── PARSE FILE ───────────────────────────────────────────
    with st.spinner("Lettura file fondi…"):
        funds_df = parse_funds_file(file_fondi.read(), file_fondi.name)
    if funds_df.empty:
        st.error("❌ Nessun fondo trovato."); return

    market_text = ""
    if file_mercato:
        with st.spinner("Lettura contesto mercati…"):
            market_text = read_text_file(file_mercato.read(), file_mercato.name)

    mifid_df = pd.DataFrame()

    fund_sheets = [f.read() for f in files_schede] if files_schede else []

    # ── PARSE CATALOGO UNP ──────────────────────────────────
    unp_df = pd.DataFrame()
    if file_unp:
        with st.spinner("Lettura catalogo UNP…"):
            unp_df = parse_unp_catalog(file_unp.read(), file_unp.name)
    # Merge UNP in funds_df per ISIN
    if not unp_df.empty and "isin" in unp_df.columns:
        merge_cols = ["isin"] + [c for c in ["unp","iunp"] if c in unp_df.columns]
        funds_df = funds_df.merge(unp_df[merge_cols], on="isin", how="left")
        if "unp"  not in funds_df.columns: funds_df["unp"]  = None
        if "iunp" not in funds_df.columns: funds_df["iunp"] = None
    else:
        funds_df["unp"]  = None
        funds_df["iunp"] = None

    n_unp = int(funds_df["unp"].notna().sum())

    # ── OVERVIEW UNIVERSO ───────────────────────────────────
    n_f = len(funds_df); n_isin = (funds_df["isin"]!="").sum()
    c1,c2,c3,c4 = st.columns(4)
    for col,v,l,s in [
        (c1, str(n_f),    "Fondi universo",  f"{funds_df['macro_cat'].nunique()} categorie"),
        (c2, str(n_isin), "Con ISIN",        "dati da FondiDoc + Morningstar"),
        (c3, f"{n_unp}" if n_unp else "—",   "Con UNP",   "da catalogo prodotti" if n_unp else "carica catalogo UNP"),
        (c4, "✓" if market_text else "—",    "View mercato", "caricata" if market_text else "non caricata"),
    ]: col.markdown(f'<div class="kpi"><div class="kpi-label">{l}</div><div class="kpi-value">{v}</div><div class="kpi-sub">{s}</div></div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    col_l, col_r = st.columns([1.2, 0.8], gap="large")
    with col_l:
        st.markdown('<p class="sec-title">Universo Fondi</p>', unsafe_allow_html=True)
        dc = [c for c in ["nome","isin","categoria","macro_cat","unp","iunp"] if c in funds_df.columns]
        rename_map = {"nome":"Fondo","isin":"ISIN","categoria":"Categoria",
                      "macro_cat":"Macro","unp":"UNP %","iunp":"IUNP %"}
        st.dataframe(funds_df[dc].rename(columns=rename_map),
                     height=320, use_container_width=True, hide_index=True)
    with col_r:
        if market_text:
            st.markdown('<p class="sec-title">Contesto Mercato</p>', unsafe_allow_html=True)
            st.markdown(f"<div style='background:#f0f7ff;border:1px solid #bfdbfe;border-radius:10px;"
                        f"padding:1rem;font-size:.82rem;color:#1e40af;max-height:320px;overflow-y:auto;"
                        f"line-height:1.7;'>{market_text[:700].replace(chr(10),'<br/>')}</div>",
                        unsafe_allow_html=True)
        else:
            st.markdown('<p class="sec-title">Distribuzione Categorie</p>', unsafe_allow_html=True)
            mc_df = funds_df["macro_cat"].value_counts().reset_index()
            mc_df.columns = ["Categoria","Fondi"]
            st.dataframe(mc_df, height=320, use_container_width=True, hide_index=True)

    # ── CONFIGURA PORTAFOGLIO LIBERO ────────────────────────────
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<p class="sec-title">⚙️ Portafoglio Libero — selezione e pesi</p>', unsafe_allow_html=True)

    # Calcola top-25 di default per la selezione iniziale
    _sig_def   = _market_signals(market_text)
    _scored_def = _score_funds(funds_df, _sig_def, mifid_df if not mifid_df.empty else None)
    _default25  = _scored_def.head(25)["nome"].tolist()
    all_nomi    = funds_df["nome"].tolist()

    # Inizializza session state
    if "libero_sel" not in st.session_state:
        st.session_state["libero_sel"] = _default25
    if "libero_w_df" not in st.session_state:
        st.session_state["libero_w_df"] = None

    with st.expander("Seleziona fondi e imposta i pesi per il Portafoglio Libero", expanded=True):
        sel = st.multiselect(
            "Cerca e seleziona fondi dall'universo:",
            options=all_nomi,
            default=[n for n in st.session_state["libero_sel"] if n in all_nomi],
            key="libero_ms",
            help="Puoi cercare per nome. La selezione default è il top-25 per score rispetto al contesto mercati."
        )
        st.session_state["libero_sel"] = sel

        if sel:
            n_sel = len(sel)
            # Ricostruisci il DataFrame pesi se la selezione è cambiata
            prev_df = st.session_state.get("libero_w_df")
            if prev_df is None or set(prev_df["Fondo"].tolist()) != set(sel):
                eq_w = round(100.0 / n_sel, 2)
                # Mantieni pesi esistenti dove possibile
                prev_map = dict(zip(prev_df["Fondo"], prev_df["Peso (%)"])) if prev_df is not None else {}
                rows = [{"Fondo": n, "Peso (%)": prev_map.get(n, eq_w)} for n in sel]
                st.session_state["libero_w_df"] = pd.DataFrame(rows)

            edited = st.data_editor(
                st.session_state["libero_w_df"],
                key="libero_de",
                hide_index=True,
                use_container_width=True,
                column_config={
                    "Fondo":    st.column_config.TextColumn("Fondo", disabled=True, width="large"),
                    "Peso (%)": st.column_config.NumberColumn(
                        "Peso (%)", min_value=0.0, max_value=100.0, step=0.5, format="%.1f"),
                }
            )
            st.session_state["libero_w_df"] = edited

            total_w = float(edited["Peso (%)"].sum())
            ca, cb, cc = st.columns([1, 1, 3])
            diff = total_w - 100.0
            ca.metric("Totale pesi", f"{total_w:.1f}%",
                      delta=f"{diff:+.1f}%" if abs(diff) > 0.1 else None,
                      delta_color="inverse")
            if abs(diff) > 0.5:
                cb.warning("⚠️ Non somma 100%")
            else:
                cb.success("✅ Pesi OK")
            if cc.button("⚖️  Normalizza a 100%", key="btn_norm_lib"):
                if total_w > 0:
                    norm = edited.copy()
                    norm["Peso (%)"] = (norm["Peso (%)"] / total_w * 100).round(2)
                    st.session_state["libero_w_df"] = norm
                    st.rerun()
        else:
            st.info("Seleziona almeno un fondo per configurare il Portafoglio Libero.")

    # ── GENERA ──────────────────────────────────────────────────
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<p class="sec-title">Generazione Portafogli</p>', unsafe_allow_html=True)

    btn_col, info_col = st.columns([1, 2])
    with info_col:
        engine = "AI (Claude)" if (api_key and _HAS_ANTHROPIC) else "Rules-based"
        n_sch  = len(fund_sheets)
        n_lib  = len(st.session_state.get("libero_sel") or [])
        has_alloc = n_sch > 0
        st.markdown(f"""<div style='background:#eff6ff;border:1px solid #bfdbfe;border-radius:10px;padding:1rem 1.25rem;'>
          <div style='font-size:.8rem;color:#1d4ed8;font-weight:600;margin-bottom:.4rem;'>Il report PDF include:</div>
          <div style='font-size:.82rem;color:#1e40af;line-height:1.9;'>
            ✓ Copertina con contesto di mercato<br>
            ✓ <b>Portafoglio Articolato</b> — 12 fondi diversificati + metriche<br>
            ✓ <b>Portafoglio Short</b> — 6 fondi alta convinzione<br>
            ✓ <b>Portafoglio Libero</b> — {n_lib} fondi selezionati manualmente<br>
            ✓ Schede analitiche (FondiDoc + Morningstar)<br>
            {'✓ <b>Asset allocation da schede prodotto (' + str(n_sch) + ' PDF)</b>' if has_alloc else '○ Asset allocation da macro-categoria (nessuna scheda caricata)'}<br>
            {'✓ <b>Allegati: ' + str(n_sch) + ' schede prodotto PDF</b>' if n_sch else ''}<br>
            <span style='color:#3b82f6;'>⚙️ Motore: <b>{engine}</b></span>
          </div></div>""", unsafe_allow_html=True)

    with btn_col:
        if st.button("🚀  Genera Portafogli + PDF", use_container_width=True, type="primary"):

            # 1. Costruisci override Portafoglio Libero dalla selezione manuale
            libero_override = None
            w_df = st.session_state.get("libero_w_df")
            if w_df is not None and not w_df.empty:
                lib_funds = []
                for _, rw in w_df.iterrows():
                    nome = rw["Fondo"]; peso = float(rw["Peso (%)"])
                    if peso <= 0: continue
                    fd_r = funds_df[funds_df["nome"] == nome]
                    if fd_r.empty: continue
                    r = fd_r.iloc[0]
                    lf = {"nome": nome, "isin": r.get("isin",""),
                          "macro_cat": r.get("macro_cat","Altro"),
                          "color": r.get("color","#94A3B8"), "peso": peso}
                    if "unp"  in r.index and str(r.get("unp",""))  not in ("","nan","None"):
                        lf["unp"]  = r["unp"]
                    if "iunp" in r.index and str(r.get("iunp","")) not in ("","nan","None"):
                        lf["iunp"] = r["iunp"]
                    lib_funds.append(lf)
                if lib_funds:
                    tot_w = sum(f["peso"] for f in lib_funds)
                    if tot_w > 0:
                        lib_funds = [{**f, "peso": round(f["peso"]*100/tot_w, 2)} for f in lib_funds]
                    libero_override = {
                        "funds": lib_funds,
                        "rationale": (f"Portafoglio Libero: {len(lib_funds)} fondi selezionati manualmente "
                                      f"con pesi personalizzati.")
                    }

            # 2. Costruisci portafogli (scoring + libero override)
            with st.spinner("🧠 Costruzione portafogli…"):
                portfolios = construct_portfolios(funds_df, market_text, mifid_df,
                                                  api_key, libero_override)

            # 3. Raccogli fondi unici nei portafogli
            ptf_nomi = set()
            for pk in ["articolato","short","libero"]:
                for f in portfolios.get(pk,{}).get("funds",[]):
                    ptf_nomi.add(f["nome"])
            ptf_df = funds_df[funds_df["nome"].isin(ptf_nomi)].reset_index(drop=True)
            n_fetch = len(ptf_df)

            # 4. Scarica dati FondiDoc/Morningstar solo per i fondi selezionati
            prog = st.progress(0, text=f"Recupero dati per {n_fetch} fondi…")
            fund_data = fetch_all(ptf_df, lambda v: prog.progress(v, text=f"Fetch {int(v*100)}%…"))
            prog.progress(1.0, text=f"✅ Dati recuperati ({n_fetch} fondi)")

            # 5. Leggi schede prodotto per asset allocation granulare
            schede_alloc = {}
            if fund_sheets:
                with st.spinner(f"📄 Lettura {len(fund_sheets)} schede prodotto per asset allocation…"):
                    schede_alloc = _parse_schede_alloc(fund_sheets, ptf_df)
                if schede_alloc:
                    st.success(f"✅ Asset allocation estratta da {len(schede_alloc)} schede prodotto")
                else:
                    st.info("ℹ️ Nessun dato di asset allocation trovato nelle schede PDF — uso macro-categoria")

            with st.spinner("📄 Generazione PDF…"):
                try:
                    market_view = _extract_market_view(market_text) if market_text else None
                    pdf_bytes = generate_pdf(portfolios, ptf_df, fund_data, market_text,
                                             fund_sheets, schede_alloc, market_view,
                                             advisor_note=advisor_note)
                    prog.empty()

                    # Anteprima quadro mercato
                    if market_view and (market_view.get("asset_views") or market_view.get("themes")):
                        st.markdown("---")
                        st.markdown('<p class="sec-title">📊 Quadro di Mercato estratto</p>', unsafe_allow_html=True)
                        qa, qb = st.columns(2)
                        with qa:
                            if market_view.get("asset_views"):
                                st.markdown("**Posizionamento Asset Class**")
                                for av in market_view["asset_views"]:
                                    color_map = {"Sovrappeso (+)":"🟢","Sottopeso (−)":"🔴","Neutrale (=)":"🟡"}
                                    icon = color_map.get(av["view"],"⚪")
                                    st.write(f"{icon} **{av['asset']}**: {av['view']}")
                        with qb:
                            risk_icon = {"low":"🔴 Risk-Off","high":"🟢 Risk-On","neutral":"🟡 Neutrale"}
                            st.markdown(f"**Propensione rischio:** {risk_icon.get(market_view.get('risk','neutral'),'—')}")
                            if market_view.get("themes"):
                                st.markdown("**Temi chiave:**")
                                for th in market_view["themes"]:
                                    st.write(f"▸ {th}")
                            if market_view.get("geo_views"):
                                st.markdown("**Regioni citate:**")
                                st.write("  ·  ".join(g["region"] for g in market_view["geo_views"]))

                    # Anteprima portafogli
                    st.markdown("---")
                    st.markdown('<p class="sec-title">Anteprima Portafogli</p>', unsafe_allow_html=True)
                    tabs = st.tabs(["📊 Articolato","⚡ Short","🎨 Libero"])
                    for tab, pk in zip(tabs, ["articolato","short","libero"]):
                        with tab:
                            ptf = portfolios.get(pk, {})
                            if not ptf or not ptf.get("funds"): st.info("Nessun dato."); continue
                            if ptf.get("rationale"): st.info(ptf["rationale"])
                            m = calc_metrics(ptf["funds"], fund_data, schede_alloc)
                            # Mostra UNP/IUNP ponderato se disponibili
                            unp_w  = m.get("unp_w")
                            iunp_w = m.get("iunp_w")
                            n_kpi  = 6 + (1 if unp_w is not None else 0) + (1 if iunp_w is not None else 0)
                            mc2 = st.columns(n_kpi)
                            kpi_list = [("ytd","YTD"),("perf_1y","1A"),("perf_3y","3A"),
                                        ("perf_5y","5A"),("var_1y","VaR 1A"),("sharpe_3y","Sharpe 3A")]
                            if unp_w  is not None: kpi_list.append(("unp_w",  "UNP %"))
                            if iunp_w is not None: kpi_list.append(("iunp_w", "IUNP %"))
                            for col2,(k2,lb) in zip(mc2, kpi_list):
                                v2 = m.get(k2)
                                if k2 in ("unp_w","iunp_w"):
                                    col2.metric(lb, f"{v2:.3f}%" if v2 is not None else "n.d.")
                                else:
                                    col2.metric(lb, v2 if v2 else "n.d.")
                            disp_cols = [c for c in ["nome","isin","macro_cat","peso","unp","iunp"]
                                         if c in pd.DataFrame(ptf["funds"]).columns]
                            st.dataframe(
                                pd.DataFrame(ptf["funds"])[disp_cols].rename(
                                    columns={"nome":"Fondo","isin":"ISIN","macro_cat":"Categoria",
                                             "peso":"Peso %","unp":"UNP %","iunp":"IUNP %"}),
                                hide_index=True, use_container_width=True)

                    st.markdown("---")
                    st.download_button("📥   Scarica Report PDF Completo", data=pdf_bytes,
                                       file_name=f"Azimut_Builder_{datetime.date.today().strftime('%Y%m%d')}.pdf",
                                       mime="application/pdf", use_container_width=True)
                    st.success(f"✅ PDF pronto — {n_fetch} fondi selezionati, 3 portafogli"
                               f"{', asset alloc da schede' if schede_alloc else ''}"
                               f"{', ' + str(len(fund_sheets)) + ' allegati' if fund_sheets else ''}")
                except Exception as e:
                    import traceback
                    st.error(f"Errore PDF: {e}")
                    st.code(traceback.format_exc())


if __name__ == "__main__":
    main()
