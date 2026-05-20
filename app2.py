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

def _market_signals(text: str) -> dict:
    if not text: return {"eq":0,"bd":0,"regions":[],"risk":"neutral"}
    t = text.lower()
    eq = sum(1 for w in ["rialzo","bullish","azionario positivo","overweight equity","sovrappeso azionario"] if w in t) \
       - sum(1 for w in ["ribasso","bearish","azionario negativo","sottopeso azionario","sell equity"] if w in t)
    bd = sum(1 for w in ["obbligazionario positivo","overweight bond","tassi in calo","duration lunga"] if w in t) \
       - sum(1 for w in ["obbligazionario negativo","underweight bond","tassi in rialzo","duration corta"] if w in t)
    regions = ([r for r in ["Europa","USA","Emergenti","Globale","Asia","Giappone"]
                if r.lower() in t or (r=="USA" and ("usa" in t or "america" in t or "s&p" in t))])
    risk = "low" if any(w in t for w in ["risk off","cautela","prudente","difensivo","bassa volatilità"]) \
      else "high" if any(w in t for w in ["risk on","aggressivo","growth","high risk"]) else "neutral"
    return {"eq":eq,"bd":bd,"regions":regions,"risk":risk}


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


def construct_portfolios(funds_df, market_text, mifid_df=None, api_key=""):
    if api_key and _HAS_ANTHROPIC:
        try: return _construct_ai(funds_df, market_text, mifid_df, api_key)
        except Exception as e: st.warning(f"AI fallita ({e}) — uso rules-based.")
    return _construct_rules(funds_df, market_text, mifid_df)


def _construct_rules(funds_df, market_text, mifid_df=None):
    signals = _market_signals(market_text)
    scored  = _score_funds(funds_df, signals, mifid_df)

    def _fund_entry(r):
        return {"nome":r["nome"],"isin":r.get("isin",""),
                "macro_cat":r.get("macro_cat","Altro"),"color":r.get("color","#94A3B8")}

    # Articolato: max 12 fondi, diversificato per categoria
    cats = scored["macro_cat"].unique(); art = []
    per_cat = max(2, 12 // max(len(cats),1))
    for cat in cats:
        for _,r in scored[scored["macro_cat"]==cat].head(per_cat).iterrows():
            if len(art) >= 12: break
            art.append({**_fund_entry(r),"peso":round(100/min(12,len(scored)),1)})
    art = _norm_w(art[:12])

    # Short: top 5 per score
    short = [_norm_w([{**_fund_entry(r),"peso":20.0} for _,r in scored.head(5).iterrows()])]
    short = _norm_w([{**_fund_entry(r),"peso":20.0} for _,r in scored.head(5).iterrows()])

    # Libero: tutti, pesi uguali
    n = len(funds_df)
    libero = _norm_w([{**_fund_entry(r),"peso":round(100/max(n,1),1)} for _,r in scored.iterrows()])

    eq_label = ("Sovrappeso azionario" if signals["eq"]>0
                else "Sovrappeso obbligazionario" if signals["bd"]>0 else "Allocazione bilanciata")
    return {
        "articolato": {"funds":art,   "rationale":f"Portafoglio su {len(art)} fondi. {eq_label}. "
                        f"Regioni preferite: {', '.join(signals['regions']) or 'globale'}."},
        "short":      {"funds":short, "rationale":f"Alta convinzione: {len(short)} fondi più allineati alla view di mercato."},
        "libero":     {"funds":libero,"rationale":f"Tutti i {len(libero)} fondi dell'universo a pesi uguali — personalizzare."},
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

def calc_metrics(ptf_funds: list, fund_data: dict) -> dict:
    keys = ["ytd","perf_1y","perf_3y","perf_5y","vol_1y","vol_3y","var_1y","sharpe_3y","neg_vol_1y","sortino_1y"]
    tot = {k:0.0 for k in keys}; wt = {k:0.0 for k in keys}; macro_alloc = {}
    for f in ptf_funds:
        w = f["peso"]/100.0
        ana = fund_data.get(f["nome"],{}).get("fondidoc",{}).get("analysis",{})
        for k in keys:
            try:
                num = float(str(ana.get(k,"")).replace("%","").replace(",",".").strip())
                tot[k]+=num*w; wt[k]+=w
            except: pass
        mc = f.get("macro_cat","Altro")
        macro_alloc[mc] = macro_alloc.get(mc,0.0)+f["peso"]
    result = {k: (f"{tot[k]/wt[k]:+.2f}%" if wt[k]>0.01 else "N/D") for k in keys}
    result["macro_alloc"] = macro_alloc
    return result


# ════════════════════════════════════════════════════════════
# CHARTS
# ════════════════════════════════════════════════════════════

def _pie_funds(ptf_funds, title):
    names   = [f["nome"][:32]+("…" if len(f["nome"])>32 else "") for f in ptf_funds]
    weights = [f["peso"] for f in ptf_funds]
    colors  = [f.get("color","#94A3B8") for f in ptf_funds]
    fig,(ax1,ax2) = plt.subplots(1,2,figsize=(11,5),gridspec_kw={"width_ratios":[1.3,1]})
    wedges,_,ats = ax1.pie(weights,colors=colors,autopct=lambda p:f"{p:.1f}%" if p>4 else "",
                            pctdistance=0.72,wedgeprops=dict(width=0.58,edgecolor="white",linewidth=2),startangle=90)
    for at in ats: at.set_fontsize(8); at.set_color("white"); at.set_fontweight("bold")
    ax1.text(0,0,title[:8],ha="center",va="center",fontsize=11,fontweight="bold",color="#0D1B2A")
    ax2.axis("off")
    ax2.legend([mpatches.Patch(color=colors[i],label=f"{names[i]}  {weights[i]:.1f}%")
                for i in range(len(ptf_funds))],
               loc="center left",frameon=False,fontsize=7.5,labelspacing=0.9,handlelength=1.2)
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
    ax2.legend([mpatches.Patch(color=colors[i],label=f"{keys[i]}  {vals[i]:.1f}%")
                for i in range(len(keys))],
               loc="center left",frameon=False,fontsize=9,labelspacing=1.1,handlelength=1.3)
    fig.patch.set_facecolor("#FFFFFF"); plt.tight_layout(pad=1.2)
    buf=io.BytesIO(); fig.savefig(buf,format="png",dpi=150,bbox_inches="tight",facecolor="white")
    plt.close(fig); buf.seek(0); return buf


# ════════════════════════════════════════════════════════════
# PDF GENERATOR
# ════════════════════════════════════════════════════════════

def generate_pdf(portfolios, funds_df, fund_data, market_text, fund_sheets=None):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            leftMargin=2*cm, rightMargin=2*cm,
                            topMargin=2.2*cm, bottomMargin=2.2*cm)
    ss = getSampleStyleSheet()
    def S(n,**kw): return ParagraphStyle(n, parent=ss["Normal"], **kw)
    T   = S("T2",  fontName="Helvetica-Bold",  fontSize=22, textColor=rl_colors.HexColor("#0D1B2A"), spaceAfter=4, leading=28)
    EY  = S("EY2", fontName="Helvetica",       fontSize=8,  textColor=rl_colors.HexColor("#94A3B8"), spaceAfter=4, letterSpacing=1.5)
    SU  = S("SU2", fontName="Helvetica",       fontSize=10, textColor=rl_colors.HexColor("#64748B"), spaceAfter=4)
    SC  = S("SC2", fontName="Helvetica-Bold",  fontSize=11, textColor=rl_colors.HexColor("#0D1B2A"), spaceBefore=14, spaceAfter=8)
    BD  = S("BD2", fontName="Helvetica",       fontSize=8.5,textColor=rl_colors.HexColor("#1E293B"), leading=13)
    SM  = S("SM2", fontName="Helvetica",       fontSize=7.5,textColor=rl_colors.HexColor("#1E293B"), leading=11)
    IT  = S("IT2", fontName="Helvetica-Oblique",fontSize=8.5,textColor=rl_colors.HexColor("#475569"), leading=13)
    FT2 = S("FT2", fontName="Helvetica-Oblique",fontSize=7, textColor=rl_colors.HexColor("#94A3B8"), leading=10)
    HDR = S("HDR2",fontName="Helvetica-Bold",  fontSize=7.5,textColor=rl_colors.white, leading=11)
    WH2 = S("WH2", fontName="Helvetica-Bold",  fontSize=8,  textColor=rl_colors.white, leading=11)
    NOTE= S("NT2", fontName="Helvetica-Oblique",fontSize=6.5,textColor=rl_colors.HexColor("#94A3B8"), leading=9)
    LK2 = S("LK2", fontName="Helvetica",       fontSize=7.5,textColor=rl_colors.HexColor("#1B4FBB"), spaceAfter=2)
    FS2 = S("FS2", fontName="Helvetica-Bold",  fontSize=12, textColor=rl_colors.HexColor("#0D1B2A"), spaceBefore=4, spaceAfter=2)
    FK2 = S("FK2", fontName="Helvetica",       fontSize=7.5,textColor=rl_colors.HexColor("#64748B"), spaceAfter=2)

    story = []

    def accent():
        return Table([[""]], colWidths=[17*cm], rowHeights=[10], style=TableStyle([
            ("BACKGROUND",(0,0),(-1,-1),rl_colors.HexColor("#0D1B2A")),
            ("LINEBELOW",(0,0),(-1,-1),3,rl_colors.HexColor("#C9A84C")),
        ]))

    def kpi_cell(v,l):
        return Paragraph(f'<font size="17"><b>{v}</b></font><br/>'
                         f'<font size="7.5" color="#64748B">{l}</font>', BD)

    def pv(val):
        try:
            num = float(str(val).replace("%","").replace(",",".").replace("+",""))
            c = "#1A7A4A" if num>0 else ("#C0392B" if num<0 else "#475569")
            return Paragraph(f'<font color="{c}"><b>{val}</b></font>', SM)
        except: return Paragraph(str(val), SM)

    # ── PAG 1: COPERTINA ────────────────────────────────────
    story += [accent(), Spacer(1,14),
              Paragraph("AZIMUT INVESTMENTS  ·  PORTFOLIO BUILDER", EY), Spacer(1,4),
              Paragraph("Proposta di Portafoglio", T),
              Paragraph(f"AI-Assisted  ·  {datetime.date.today().strftime('%d %B %Y')}", SU),
              HRFlowable(width="100%",thickness=0.8,color=rl_colors.HexColor("#E2E8F0"),spaceAfter=14)]

    n_f = len(funds_df)
    macros = funds_df["macro_cat"].value_counts()
    top_m  = macros.index[0] if len(macros)>0 else "—"
    kpi_row = Table([[kpi_cell(str(n_f),"Fondi universo"),
                      kpi_cell(str(len(macros)),"Macro categorie"),
                      kpi_cell(top_m[:12],"Categoria principale"),
                      kpi_cell(datetime.date.today().strftime("%m/%Y"),"Data report")]],
                    colWidths=[4.25*cm]*4)
    kpi_row.setStyle(TableStyle([
        ("BOX",(0,0),(-1,-1),0.8,rl_colors.HexColor("#E2E8F0")),
        ("INNERGRID",(0,0),(-1,-1),0.8,rl_colors.HexColor("#E2E8F0")),
        ("BACKGROUND",(0,0),(-1,-1),rl_colors.HexColor("#F8FAFC")),
        ("PADDING",(0,0),(-1,-1),12),("ALIGN",(0,0),(-1,-1),"CENTER"),("VALIGN",(0,0),(-1,-1),"MIDDLE"),
    ]))
    story.append(kpi_row)
    story.append(Spacer(1,14))

    if market_text and market_text.strip():
        story.append(Paragraph("Contesto di Mercato", SC))
        excerpt = market_text.strip()[:900]+("…" if len(market_text.strip())>900 else "")
        story.append(Paragraph(excerpt.replace("\n","<br/>"), IT))

    story.append(PageBreak())

    # ── PAG 2+: I TRE PORTAFOGLI ────────────────────────────
    ptf_meta = [
        ("articolato","Portafoglio Articolato","#1B4FBB"),
        ("short",     "Portafoglio Short",     "#065F46"),
        ("libero",    "Portafoglio Libero",     "#7C3AED"),
    ]

    for ptf_key, ptf_label, ptf_color in ptf_meta:
        ptf = portfolios.get(ptf_key,{})
        if not ptf or not ptf.get("funds"): continue
        funds_list = ptf["funds"]; rationale = ptf.get("rationale","")
        metrics = calc_metrics(funds_list, fund_data)

        story += [accent(), Spacer(1,14),
                  Paragraph("AZIMUT INVESTMENTS  ·  PORTFOLIO BUILDER", EY), Spacer(1,4),
                  Paragraph(ptf_label, T),
                  Paragraph(f"{len(funds_list)} fondi  ·  {datetime.date.today().strftime('%d %B %Y')}", SU),
                  HRFlowable(width="100%",thickness=0.8,color=rl_colors.HexColor("#E2E8F0"),spaceAfter=10)]

        # Rationale box
        if rationale:
            rt = Table([[Paragraph(rationale, IT)]], colWidths=[17*cm])
            rt.setStyle(TableStyle([
                ("BACKGROUND",(0,0),(-1,-1),rl_colors.HexColor("#F0F7FF")),
                ("PADDING",(0,0),(-1,-1),10),
                ("LINEBELOW",(0,0),(-1,-1),2,rl_colors.HexColor(ptf_color)),
                ("BOX",(0,0),(-1,-1),0.5,rl_colors.HexColor("#BFDBFE")),
            ]))
            story += [rt, Spacer(1,10)]

        # KPI metriche portafoglio
        mk = Table([[kpi_cell(metrics.get("ytd","N/D"),"YTD"),
                     kpi_cell(metrics.get("perf_1y","N/D"),"1 Anno"),
                     kpi_cell(metrics.get("perf_3y","N/D"),"3 Anni"),
                     kpi_cell(metrics.get("perf_5y","N/D"),"5 Anni"),
                     kpi_cell(metrics.get("var_1y","N/D"),"VaR 1A"),
                     kpi_cell(metrics.get("sharpe_3y","N/D"),"Sharpe 3A"),
                     kpi_cell(metrics.get("neg_vol_1y","N/D"),"Vol.Neg. 1A")]],
                   colWidths=[2.43*cm]*7)
        mk.setStyle(TableStyle([
            ("BOX",(0,0),(-1,-1),0.8,rl_colors.HexColor("#E2E8F0")),
            ("INNERGRID",(0,0),(-1,-1),0.8,rl_colors.HexColor("#E2E8F0")),
            ("BACKGROUND",(0,0),(-1,-1),rl_colors.HexColor("#F8FAFC")),
            ("PADDING",(0,0),(-1,-1),10),("ALIGN",(0,0),(-1,-1),"CENTER"),("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ]))
        story += [mk, Spacer(1,10)]

        # Grafici
        story.append(RLImage(_pie_funds(funds_list, ptf_label), width=15*cm, height=6.5*cm))
        mb = _pie_macro(metrics.get("macro_alloc",{}))
        if mb: story += [Spacer(1,6), RLImage(mb, width=15*cm, height=5*cm)]
        story.append(Spacer(1,10))

        # Tabella fondi
        hdr_row = [Paragraph(f"<b>{t}</b>",HDR) for t in
                   ["Fondo","ISIN","Categoria","Peso","YTD","1A","3A","5A","VaR 1A","Sharpe 3A"]]
        rows_data = [hdr_row]
        for f in funds_list:
            nome = f["nome"]
            ana  = fund_data.get(nome,{}).get("fondidoc",{}).get("analysis",{})
            def gf(k): return ana.get(k,"—")
            rows_data.append([
                Paragraph(nome[:42], SM), Paragraph(f.get("isin","—"), SM),
                Paragraph(f.get("macro_cat","—")[:22], SM),
                Paragraph(f"<b>{f['peso']:.1f}%</b>", SM),
                pv(gf("ytd")), pv(gf("perf_1y")), pv(gf("perf_3y")), pv(gf("perf_5y")),
                pv(gf("var_1y")), Paragraph(gf("sharpe_3y"), SM),
            ])
        ft = Table(rows_data,
                   colWidths=[4.1*cm,2.0*cm,2.2*cm,1.2*cm,1.3*cm,1.3*cm,1.3*cm,1.3*cm,1.3*cm,1.3*cm],
                   repeatRows=1)
        ft.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0),rl_colors.HexColor("#0D1B2A")),
            ("FONTSIZE",(0,0),(-1,-1),7.5),("PADDING",(0,0),(-1,-1),4),
            ("ROWBACKGROUNDS",(0,1),(-1,-1),[rl_colors.white,rl_colors.HexColor("#F8FAFC")]),
            ("LINEBELOW",(0,0),(-1,-1),0.4,rl_colors.HexColor("#E2E8F0")),
            ("ALIGN",(3,0),(-1,-1),"CENTER"),("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ]))
        story.append(KeepTogether([ft]))
        story += [Spacer(1,6),
                  Paragraph("◆ Metriche portafoglio = media ponderata sui fondi con dati disponibili (FondiDoc). "
                             "YTD/1A/3A/5A: rendimenti %. VaR 1A: Value at Risk annuale. Sharpe 3A: indice di Sharpe. "
                             "Dati a titolo indicativo.", NOTE),
                  PageBreak()]

    # ── SCHEDE SINGOLI FONDI ────────────────────────────────
    story += [accent(), Spacer(1,14),
              Paragraph("AZIMUT INVESTMENTS  ·  PORTFOLIO BUILDER", EY), Spacer(1,4),
              Paragraph("Schede Analitiche Fondi", T),
              Paragraph(f"Universo completo  ·  Fonte: FondiDoc + Morningstar  ·  {datetime.date.today().strftime('%d %B %Y')}", SU),
              HRFlowable(width="100%",thickness=0.8,color=rl_colors.HexColor("#E2E8F0"),spaceAfter=6),
              Paragraph('🔍 <link href="https://www.morningstar.it/it/funds/SecuritySearchResults.aspx">'
                        '<u>Motore di ricerca Morningstar</u></link>', LK2),
              Spacer(1,10)]

    for idx, (_, fr) in enumerate(funds_df.iterrows()):
        nome = fr["nome"]; isin = fr.get("isin","")
        fd   = fund_data.get(nome,{})
        ov   = fd.get("fondidoc",{}).get("overview",{})
        ana  = fd.get("fondidoc",{}).get("analysis",{})
        ms   = fd.get("morningstar",{})

        def gv(k, src=None): return (src or ana).get(k,"—")

        srri_s  = f"SRRI {gv('srri',ov)}/7"        if gv('srri',ov)       not in ("—","")  else ""
        nav_s   = f"NAV {gv('nav')} €"              if gv('nav')           not in ("—","")  else ""
        rat_s   = f"FIDArating {gv('fida_rating',ov)}" if gv('fida_rating',ov) not in ("—","")  else ""
        meta    = "  ·  ".join(x for x in [srri_s, rat_s, nav_s] if x)
        isin_s  = f"  ·  ISIN: <b>{isin}</b>" if isin else ""
        mc_s    = fr.get("macro_cat",""); mifid_v = fr.get("mifid","—")

        hdr_rows = [
            [Paragraph(f"<b>{nome}</b>", FS2)],
            [Paragraph(f"MIFID: <b>{mifid_v}/7</b>  ·  {fr.get('categoria','')}{isin_s}", FK2)],
            [Paragraph(meta or "—", FK2)],
        ]
        if ms.get("ms_url"):
            hdr_rows.append([Paragraph(
                f'<link href="{ms["ms_url"]}"><u>↗ {ms.get("ms_name","Morningstar")}</u></link>', FK2)])
        if fr.get("descrizione","").strip():
            hdr_rows.append([Paragraph(str(fr["descrizione"])[:200], FK2)])

        hdr_tbl = Table(hdr_rows, colWidths=[17*cm])
        hdr_tbl.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,-1),rl_colors.HexColor("#F0F4F9")),
            ("LEFTPADDING",(0,0),(-1,-1),10),("RIGHTPADDING",(0,0),(-1,-1),10),
            ("TOPPADDING",(0,0),(-1,0),10),("BOTTOMPADDING",(0,-1),(-1,-1),10),
            ("TOPPADDING",(0,1),(-1,-1),2),("BOTTOMPADDING",(0,0),(-1,-2),2),
            ("LINEBELOW",(0,-1),(-1,-1),2,rl_colors.HexColor("#C9A84C")),
        ]))

        def pval2(v):
            try:
                num=float(str(v).replace("%","").replace(",",".").replace("+",""))
                c="#1A7A4A" if num>0 else ("#C0392B" if num<0 else "#475569")
                return Paragraph(f'<font color="{c}"><b>{v}</b></font>', BD)
            except: return Paragraph(str(v), BD)

        HDR3=S(f"HDR3_{idx}",fontName="Helvetica-Bold",fontSize=7.5,textColor=rl_colors.white,leading=11)
        perf_data=[
            [Paragraph("<b>Metrica</b>",HDR3),Paragraph("<b>YTD</b>",HDR3),
             Paragraph("<b>1 Anno</b>",HDR3),Paragraph("<b>3 Anni</b>",HDR3),Paragraph("<b>5 Anni</b>",HDR3)],
            [Paragraph("Performance",SM),pval2(gv("ytd")),pval2(gv("perf_1y")),pval2(gv("perf_3y")),pval2(gv("perf_5y"))],
            [Paragraph("Volatilità",SM),Paragraph("—",SM),Paragraph(gv("vol_1y"),SM),Paragraph(gv("vol_3y"),SM),Paragraph(gv("vol_5y"),SM)],
            [Paragraph("Vol. Neg.",SM),Paragraph("—",SM),Paragraph(gv("neg_vol_1y"),SM),Paragraph(gv("neg_vol_3y"),SM),Paragraph(gv("neg_vol_5y"),SM)],
            [Paragraph("VaR",SM),Paragraph("—",SM),Paragraph(gv("var_1y"),SM),Paragraph(gv("var_3y"),SM),Paragraph("—",SM)],
            [Paragraph("Sharpe",SM),Paragraph("—",SM),Paragraph("—",SM),Paragraph(gv("sharpe_3y"),SM),Paragraph(gv("sharpe_5y"),SM)],
            [Paragraph("Sortino",SM),Paragraph("—",SM),Paragraph(gv("sortino_1y"),SM),Paragraph("—",SM),Paragraph("—",SM)],
        ]
        perf_tbl=Table(perf_data,colWidths=[2.4*cm,1.5*cm,1.8*cm,1.8*cm,1.8*cm])
        perf_tbl.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0),rl_colors.HexColor("#0D1B2A")),
            ("FONTSIZE",(0,0),(-1,-1),7.5),("PADDING",(0,0),(-1,-1),4),
            ("ROWBACKGROUNDS",(0,1),(-1,-1),[rl_colors.white,rl_colors.HexColor("#F8FAFC")]),
            ("LINEBELOW",(0,0),(-1,-1),0.4,rl_colors.HexColor("#E2E8F0")),
            ("ALIGN",(1,0),(-1,-1),"CENTER"),("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ]))

        det_data=[
            [Paragraph("<b>Dettagli</b>",BD)],
            [Paragraph(f"Avvio: {gv('start_date',ov)}",SM)],
            [Paragraph(f"Categoria Assog.: {gv('cat_assog',ov)}",SM)],
            [Paragraph(f"Gest.: {gv('mgmt_fee',ov)}  |  Perf.: {gv('perf_fee',ov)}",SM)],
            [Paragraph(f"FIDArating: {gv('fida_rating',ov)}  |  Score: {gv('fida_score',ov)}",SM)],
            [Paragraph(f"Morningstar cat.: {ms.get('ms_cat','—')}",SM)],
        ]
        det_tbl=Table([[d[0]] for d in det_data],colWidths=[7.3*cm])
        det_tbl.setStyle(TableStyle([
            ("PADDING",(0,0),(-1,-1),3),("TOPPADDING",(0,0),(-1,0),6),
            ("LINEBELOW",(0,0),(0,0),0.8,rl_colors.HexColor("#C9A84C")),
            ("BACKGROUND",(0,0),(0,-1),rl_colors.HexColor("#F8FAFC")),
        ]))

        mid=Table([[perf_tbl,det_tbl]],colWidths=[9.7*cm,7.3*cm])
        mid.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP"),("PADDING",(0,0),(-1,-1),0),
                                  ("LEFTPADDING",(1,0),(1,-1),10)]))

        story.append(KeepTogether([Spacer(1,6),hdr_tbl,Spacer(1,6),mid]))
        if idx < len(funds_df)-1:
            story.append(HRFlowable(width="100%",thickness=0.5,
                                     color=rl_colors.HexColor("#CBD5E1"),spaceBefore=8,spaceAfter=8))

    # ── DISCLAIMER ──────────────────────────────────────────
    story += [PageBreak(),
              HRFlowable(width="100%",thickness=0.5,color=rl_colors.HexColor("#E2E8F0"),spaceAfter=8),
              Paragraph("Documento generato automaticamente a scopo illustrativo. Dati da FIDA FondiDoc e Morningstar. "
                        "I portafogli sono costruiti tramite analisi AI/rules-based del contesto di mercato e non "
                        "costituiscono offerta o consulenza di investimento. Rendimenti passati non garantiscono "
                        "risultati futuri. © Azimut Group — uso interno.", FT2)]

    doc.build(story)
    main_pdf = buf.getvalue()

    # Merge schede prodotto PDF se presenti
    if fund_sheets and _HAS_PYPDF:
        try:
            writer = PdfWriter()
            for pdf_bytes in [main_pdf] + fund_sheets:
                for page in PdfReader(io.BytesIO(pdf_bytes)).pages:
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
        st.markdown("<span style='color:#4a6582;font-size:.68rem;letter-spacing:.12em;text-transform:uppercase;font-weight:600;'>CLASSIFICAZIONE MIFID</span>", unsafe_allow_html=True)
        file_mifid = st.file_uploader("Excel/CSV/PDF — nome/ISIN + punteggio 1-7",
                                       type=["xlsx","xls","csv","pdf"], key="u_mifid", label_visibility="collapsed")
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
1. 📁 **Fondi potenziali** *(mensile)* — Excel/CSV **oppure PDF** (es. factbook): gli ISIN vengono estratti automaticamente
2. 📰 **Contesto mercati** *(più frequente)* — view PDF/TXT → guida la costruzione dei portafogli
3. 📋 **Classificazione MIFID** — Excel/CSV **oppure PDF**: ISIN + punteggio SRRI/MIFID 1-7
4. 📎 **Schede prodotto** — PDF allegati al report finale
5. 🚀 **Genera** → 3 portafogli (Articolato · Short · Libero) con rendimenti, asset allocation, VaR, Sharpe
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
    if file_mifid:
        mifid_df = parse_mifid_file(file_mifid.read(), file_mifid.name)

    fund_sheets = [f.read() for f in files_schede] if files_schede else []

    # ── OVERVIEW UNIVERSO ───────────────────────────────────
    n_f = len(funds_df); n_isin = (funds_df["isin"]!="").sum()
    c1,c2,c3,c4 = st.columns(4)
    for col,v,l,s in [
        (c1, str(n_f),  "Fondi universo",  f"{funds_df['macro_cat'].nunique()} categorie"),
        (c2, str(n_isin),"Con ISIN",        "dati da FondiDoc + Morningstar"),
        (c3, str(len(mifid_df)) if not mifid_df.empty else "—","Fondi MIFID","classificati"),
        (c4, "✓" if market_text else "—","View mercato","caricata" if market_text else "non caricata"),
    ]: col.markdown(f'<div class="kpi"><div class="kpi-label">{l}</div><div class="kpi-value">{v}</div><div class="kpi-sub">{s}</div></div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    col_l, col_r = st.columns([1.2, 0.8], gap="large")
    with col_l:
        st.markdown('<p class="sec-title">Universo Fondi</p>', unsafe_allow_html=True)
        dc = [c for c in ["nome","isin","categoria","macro_cat","mifid"] if c in funds_df.columns]
        st.dataframe(funds_df[dc].rename(columns={"nome":"Fondo","isin":"ISIN","categoria":"Categoria",
                                                    "macro_cat":"Macro","mifid":"MIFID"}),
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

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<p class="sec-title">Generazione Portafogli</p>', unsafe_allow_html=True)

    btn_col, info_col = st.columns([1,2])
    with info_col:
        engine = "AI (Claude)" if (api_key and _HAS_ANTHROPIC) else "Rules-based"
        n_sch = len(fund_sheets)
        st.markdown(f"""<div style='background:#eff6ff;border:1px solid #bfdbfe;border-radius:10px;padding:1rem 1.25rem;'>
          <div style='font-size:.8rem;color:#1d4ed8;font-weight:600;margin-bottom:.4rem;'>Il report PDF include:</div>
          <div style='font-size:.82rem;color:#1e40af;line-height:1.9;'>
            ✓ Copertina con contesto di mercato<br>
            ✓ <b>Portafoglio Articolato</b> — diversificato, grafici + metriche<br>
            ✓ <b>Portafoglio Short</b> — alta convinzione, 4-6 fondi<br>
            ✓ <b>Portafoglio Libero</b> — tutti i {n_f} fondi<br>
            ✓ Schede analitiche (FondiDoc + Morningstar + classificazioni)<br>
            {'✓ <b>Allegati: ' + str(n_sch) + ' schede prodotto PDF</b>' if n_sch else ''}<br>
            <span style='color:#3b82f6;'>⚙️ Motore: <b>{engine}</b></span>
          </div></div>""", unsafe_allow_html=True)

    with btn_col:
        if st.button("🚀  Carica Dati + Genera PDF", use_container_width=True, type="primary"):
            prog = st.progress(0, text="Recupero dati FondiDoc + Morningstar…")
            fund_data = fetch_all(funds_df, lambda v: prog.progress(v, text=f"Fetch: {int(v*100)}%…"))
            prog.progress(1.0, text="✅ Dati recuperati")

            with st.spinner("🧠 Costruzione portafogli…"):
                portfolios = construct_portfolios(funds_df, market_text, mifid_df, api_key)

            with st.spinner("📄 Generazione PDF…"):
                try:
                    pdf_bytes = generate_pdf(portfolios, funds_df, fund_data, market_text, fund_sheets)
                    prog.empty()

                    # Anteprima
                    st.markdown("---")
                    st.markdown('<p class="sec-title">Anteprima Portafogli</p>', unsafe_allow_html=True)
                    tabs = st.tabs(["📊 Articolato","⚡ Short","🎨 Libero"])
                    for tab, pk in zip(tabs, ["articolato","short","libero"]):
                        with tab:
                            ptf = portfolios.get(pk,{});
                            if not ptf or not ptf.get("funds"): st.info("Nessun dato."); continue
                            if ptf.get("rationale"): st.info(ptf["rationale"])
                            m = calc_metrics(ptf["funds"], fund_data)
                            mc2 = st.columns(6)
                            for col2,(k2,lb) in zip(mc2,[("ytd","YTD"),("perf_1y","1A"),("perf_3y","3A"),
                                                          ("perf_5y","5A"),("var_1y","VaR 1A"),("sharpe_3y","Sharpe 3A")]):
                                col2.metric(lb, m.get(k2,"N/D"))
                            st.dataframe(pd.DataFrame(ptf["funds"])[["nome","macro_cat","peso"]].rename(
                                columns={"nome":"Fondo","macro_cat":"Categoria","peso":"Peso %"}),
                                hide_index=True, use_container_width=True)

                    st.markdown("---")
                    st.download_button("📥   Scarica Report PDF Completo", data=pdf_bytes,
                                       file_name=f"Azimut_Builder_{datetime.date.today().strftime('%Y%m%d')}.pdf",
                                       mime="application/pdf", use_container_width=True)
                    st.success(f"✅ PDF pronto — {n_f} fondi, 3 portafogli{', ' + str(len(fund_sheets)) + ' allegati' if fund_sheets else ''}")
                except Exception as e:
                    import traceback
                    st.error(f"Errore PDF: {e}")
                    st.code(traceback.format_exc())


if __name__ == "__main__":
    main()
