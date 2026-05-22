# ============================================================
# SEGRETERIA SASSUOLO — COMPILAZIONE MODULI SUCCESSIONE
# app5.py
# ============================================================

import io
import re
import zipfile

import pandas as pd
import pdfplumber
import streamlit as st
from pypdf import PdfReader, PdfWriter
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas as rl_canvas

# OCR — importazione opzionale (disponibile solo se pytesseract e pdf2image installati)
try:
    import pytesseract
    from pdf2image import convert_from_bytes
    _OCR_AVAILABLE = True
except ImportError:
    _OCR_AVAILABLE = False

# ─────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Segreteria Sassuolo – Successioni",
    page_icon="📋",
    layout="wide",
)

st.markdown("""
<style>
    .main-header {
        font-size: 1.7rem; font-weight: 700; color: #1a3a5c;
        padding: 0.5rem 0; margin-bottom: 0.2rem;
    }
    .section-title {
        font-size: 1.05rem; font-weight: 700; color: #2c5f8a;
        border-bottom: 2px solid #2c5f8a; padding-bottom: 3px; margin: 1rem 0 0.6rem;
    }
    .badge-ok  { color: #28a745; font-weight: 600; }
    .badge-warn{ color: #e67e22; font-weight: 600; }
    div[data-testid="stSidebar"] .stMarkdown h3 { margin-top: 1rem; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────────────────────
_ANA_KEYS = [
    "cognome", "nome", "data_nascita", "luogo_nascita", "codice_fiscale",
    "indirizzo", "comune", "cap", "sesso",
    "numero_documento", "data_rilascio", "data_scadenza", "ente_rilascio",
]

# Map anagrafica keys → widget session-state keys
_WIDGET_KEY = {
    "cognome":          "a_cog",
    "nome":             "a_nom",
    "data_nascita":     "a_dn",
    "luogo_nascita":    "a_ln",
    "codice_fiscale":   "a_cf",
    "indirizzo":        "a_ind",
    "comune":           "a_com",
    "cap":              "a_cap",
    "sesso":            "a_ses",
    "numero_documento": "a_ndoc",
    "data_rilascio":    "a_drl",
    "data_scadenza":    "a_dsc",
    "ente_rilascio":    "a_ent",
}

def _init():
    if "anagrafica" not in st.session_state:
        st.session_state.anagrafica = {k: "" for k in _ANA_KEYS}
    if "fondi" not in st.session_state:
        st.session_state.fondi = []
    if "filled" not in st.session_state:
        st.session_state.filled = {}          # slot_idx -> {bytes, name, values}
    # Byte caches (to survive reruns after file_uploader disappears)
    for k in ["id_bytes", "id_name", "az_bytes", "az_name"]:
        if k not in st.session_state:
            st.session_state[k] = None
    for i in range(5):
        for k in [f"fb_{i}", f"fn_{i}"]:
            if k not in st.session_state:
                st.session_state[k] = None

_init()

# ─────────────────────────────────────────────────────────────
# PDF UTILITIES
# ─────────────────────────────────────────────────────────────

def pdf_text(b: bytes) -> str:
    try:
        with pdfplumber.open(io.BytesIO(b)) as pdf:
            return "\n".join(p.extract_text() or "" for p in pdf.pages)
    except Exception:
        return ""


def pdf_tables(b: bytes):
    out = []
    try:
        with pdfplumber.open(io.BytesIO(b)) as pdf:
            for page in pdf.pages:
                tbls = page.extract_tables()
                if tbls:
                    out.extend(tbls)
    except Exception:
        pass
    return out


def ocr_pdf(b: bytes) -> str:
    """OCR su PDF scansionato: converte le pagine in immagini e legge il testo con Tesseract."""
    if not _OCR_AVAILABLE:
        return ""
    try:
        images = convert_from_bytes(b, dpi=300)
        pages_text = []
        for img in images:
            text = pytesseract.image_to_string(img, lang="ita+eng")
            pages_text.append(text)
        return "\n".join(pages_text)
    except Exception:
        return ""


def pdf_fields(b: bytes) -> dict:
    try:
        return PdfReader(io.BytesIO(b)).get_fields() or {}
    except Exception:
        return {}


def fill_acroform(b: bytes, values: dict) -> bytes:
    try:
        reader = PdfReader(io.BytesIO(b))
        writer = PdfWriter()
        writer.clone_reader_document_root(reader)
        for page in writer.pages:
            writer.update_page_form_field_values(page, values)
        out = io.BytesIO()
        writer.write(out)
        return out.getvalue()
    except Exception as e:
        st.warning(f"Compilazione AcroForm: {e}")
        return b


def overlay_data_page(b: bytes, values: dict) -> bytes:
    """Append a plain-text data sheet to a non-fillable PDF."""
    reader = PdfReader(io.BytesIO(b))
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)

    # Build an extra page with the data
    packet = io.BytesIO()
    c = rl_canvas.Canvas(packet, pagesize=A4)
    w, h = A4
    c.setFont("Helvetica-Bold", 13)
    c.drawString(40, h - 50, "DATI DA INSERIRE NEL MODULO")
    c.setFont("Helvetica", 9)
    c.drawString(40, h - 66, "(foglio allegato — il modulo originale non ha campi compilabili)")
    c.setLineWidth(0.5)
    c.line(40, h - 72, w - 40, h - 72)

    labels = {
        "cognome":         "Cognome",
        "nome":            "Nome",
        "data_nascita":    "Data di nascita",
        "luogo_nascita":   "Luogo di nascita",
        "codice_fiscale":  "Codice Fiscale",
        "indirizzo":       "Indirizzo",
        "comune":          "Comune",
        "cap":             "CAP",
        "sesso":           "Sesso",
        "numero_documento":"N° documento",
        "data_rilascio":   "Data rilascio",
        "data_scadenza":   "Data scadenza",
        "nome_fondo":      "Nome fondo",
        "isin":            "ISIN",
        "quote":           "N° quote",
        "controvalore":    "Controvalore (€)",
    }

    y = h - 95
    c.setFont("Helvetica-Bold", 10)
    c.drawString(40, y, "Dati anagrafici")
    y -= 14
    c.setFont("Helvetica", 10)
    ana_keys = ["cognome","nome","data_nascita","luogo_nascita","codice_fiscale",
                "indirizzo","comune","cap","sesso","numero_documento",
                "data_rilascio","data_scadenza"]
    for k in ana_keys:
        v = values.get(k, "")
        if v:
            c.drawString(55, y, f"{labels.get(k, k)}: {v}")
            y -= 14
    y -= 6
    c.setFont("Helvetica-Bold", 10)
    c.drawString(40, y, "Dati fondo")
    y -= 14
    c.setFont("Helvetica", 10)
    for k in ["nome_fondo","isin","quote","controvalore"]:
        v = values.get(k, "")
        if v:
            c.drawString(55, y, f"{labels.get(k, k)}: {v}")
            y -= 14
    c.save()
    packet.seek(0)

    extra = PdfReader(packet)
    writer.add_page(extra.pages[0])

    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()

# ─────────────────────────────────────────────────────────────
# DATA EXTRACTION
# ─────────────────────────────────────────────────────────────

def _first(patterns: list, text: str) -> str:
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            val = m.group(1).strip()
            val = re.sub(r"\s{2,}", " ", val)
            if len(val) > 1:
                return val
    return ""


def _parse_identity_from_text(text: str) -> dict:
    """Parsa i dati anagrafici da testo grezzo (pdfplumber o OCR)."""
    t = text.upper()
    d = {}

    d["cognome"] = _first([
        r"COGNOME[:\s/]+([A-Z][A-Z\s'\-]{1,30})(?:\n|NOME|$)",
        r"SURNAME[:\s]+([A-Z][A-Z\s'\-]{1,30})(?:\n|NAME|$)",
    ], t)

    d["nome"] = _first([
        r"(?:^|\n|\s)NOME[:\s/]+([A-Z][A-Z\s'\-]{1,30})(?:\n|DATA|LUOGO|SESSO|$)",
        r"GIVEN\s+NAME[S]?[:\s]+([A-Z][A-Z\s'\-]{1,30})(?:\n|$)",
    ], t)

    d["data_nascita"] = _first([
        r"DATA\s+DI\s+NASCITA[:\s]+(\d{2}[/.\-]\d{2}[/.\-]\d{4})",
        r"NATO[/A]*\s+IL[:\s]+(\d{2}[/.\-]\d{2}[/.\-]\d{4})",
        r"DATE\s+OF\s+BIRTH[:\s]+(\d{2}[/.\-]\d{2}[/.\-]\d{4})",
        r"NASCITA[:\s]+(\d{2}[/.\-]\d{2}[/.\-]\d{4})",
        r"\b(\d{2}[/.\-]\d{2}[/.\-]\d{4})\b",   # fallback: prima data trovata
    ], t)

    d["luogo_nascita"] = _first([
        r"LUOGO\s+DI\s+NASCITA[:\s]+([A-Z][A-Z\s\(\)]{2,40})(?:\n|DATA|$)",
        r"COMUNE\s+DI\s+NASCITA[:\s]+([A-Z][A-Z\s]{2,30})(?:\n|PROV|$)",
        r"PLACE\s+OF\s+BIRTH[:\s]+([A-Z][A-Z\s,]{2,40})(?:\n|DATE|$)",
    ], t)

    d["codice_fiscale"] = _first([
        r"CODICE\s+FISCALE[:\s]+([A-Z]{6}\d{2}[A-Z]\d{2}[A-Z]\d{3}[A-Z])",
        r"\bCF[:\s]+([A-Z]{6}\d{2}[A-Z]\d{2}[A-Z]\d{3}[A-Z])",
        r"([A-Z]{6}\d{2}[A-Z]\d{2}[A-Z]\d{3}[A-Z])",
    ], t)

    d["indirizzo"] = _first([
        r"(?:INDIRIZZO|RESIDENZA|DOMICILIO)[:\s]+(.+?)(?:\n|CAP|COMUNE|$)",
        r"VIA\s+(.+?)(?:\n|CAP|$)",
        r"PIAZZA\s+(.+?)(?:\n|CAP|$)",
    ], t)

    d["sesso"] = _first([
        r"\bSESSO[:\s]+([MF])\b",
        r"\bSEX[:\s]+([MF])\b",
    ], t)

    d["numero_documento"] = _first([
        r"N[°\.]?\s*DOCUMENTO[:\s]+([A-Z0-9]{6,12})",
        r"NUMERO\s+DOCUMENTO[:\s]+([A-Z0-9]{6,12})",
        r"DOCUMENT\s+N[O°]?[:\s]+([A-Z0-9]{6,12})",
        r"\bNR[.\s]+([A-Z]{2}\d{5,7})\b",
        r"\b([A-Z]{2}\d{5}[A-Z0-9]{0,3})\b",
    ], t)

    d["data_rilascio"] = _first([
        r"DATA\s+(?:DI\s+)?(?:RILASCIO|EMISSIONE)[:\s]+(\d{2}[/.\-]\d{2}[/.\-]\d{4})",
        r"RILASCIATA?\s+IL[:\s]+(\d{2}[/.\-]\d{2}[/.\-]\d{4})",
        r"DATE\s+OF\s+ISSUE[:\s]+(\d{2}[/.\-]\d{2}[/.\-]\d{4})",
    ], t)

    d["data_scadenza"] = _first([
        r"(?:DATA\s+DI\s+)?SCADENZA[:\s]+(\d{2}[/.\-]\d{2}[/.\-]\d{4})",
        r"VALIDA?\s+FINO\s+AL[:\s]+(\d{2}[/.\-]\d{2}[/.\-]\d{4})",
        r"DATE\s+OF\s+EXPIRY[:\s]+(\d{2}[/.\-]\d{2}[/.\-]\d{4})",
        r"EXPIRY\s+DATE[:\s]+(\d{2}[/.\-]\d{2}[/.\-]\d{4})",
    ], t)

    d["ente_rilascio"] = _first([
        r"RILASCIATA?\s+DA[:\s]+(.+?)(?:\n|IL\s+\d|$)",
        r"ISSUED\s+BY[:\s]+(.+?)(?:\n|ON\s+\d|$)",
    ], t)

    return {k: v for k, v in d.items() if v}


def parse_identity(b: bytes) -> dict:
    return _parse_identity_from_text(pdf_text(b))


def _clean_num(s: str):
    s = s.strip().replace(" ", "")
    # Italian: 1.234,56  →  English: 1234.56
    if re.match(r"^\d{1,3}(\.\d{3})+(,\d+)?$", s):
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return None


def parse_azimut(b: bytes) -> list:
    text = pdf_text(b)
    tables = pdf_tables(b)
    fondi = []

    # ── Try tables first ──────────────────────────────────────
    for table in tables:
        for row in table or []:
            row = [str(c).strip() if c else "" for c in (row or [])]
            row_up = " ".join(row).upper()
            isin = re.search(r"\b([A-Z]{2}[A-Z0-9]{10})\b", row_up)
            is_fund = any(k in row_up for k in [
                "AZIMUT", "AZ FUND", "AZ ", "GESTIONE", "BILANCIATO",
                "OBBLIGAZIONARIO", "AZIONARIO", "FLESSIBILE", "CAPITAL",
                "TREND", "GROWTH", "INCOME", "GLOBAL",
            ])
            if not (isin or is_fund):
                continue

            nums = [_clean_num(c) for c in row if re.match(r"^[\d.,]+$", c.replace(" ", ""))]
            nums = [n for n in nums if n and n > 0]

            name_cands = [c for c in row if len(c) > 6 and not re.match(r"^[\d.,\s%+\-]+$", c)]
            name = max(name_cands, key=len) if name_cands else ""

            if name and nums:
                fondi.append({
                    "nome": name,
                    "isin": isin.group(1) if isin else "",
                    "quote": str(nums[0]) if nums else "",
                    "controvalore": str(nums[-1]) if len(nums) > 1 else "",
                })

    # ── Fallback: line-by-line text ───────────────────────────
    if not fondi:
        cur = None
        for line in text.split("\n"):
            lu = line.upper().strip()
            if not lu:
                continue
            isin = re.search(r"\b([A-Z]{2}[A-Z0-9]{10})\b", lu)
            is_fund = any(k in lu for k in [
                "AZIMUT", "AZ FUND", "GESTIONE", "BILANCIATO",
                "OBBLIGAZIONARIO", "AZIONARIO", "FLESSIBILE",
            ])
            if is_fund and not isin:
                if cur:
                    fondi.append(cur)
                cur = {"nome": line.strip(), "isin": "", "quote": "", "controvalore": ""}
            if isin and cur:
                cur["isin"] = isin.group(1)
            if cur:
                nums = [_clean_num(t) for t in re.findall(r"\b[\d.,]+\b", lu)]
                nums = [n for n in nums if n and n > 0.001]
                if nums and not cur.get("quote"):
                    cur["quote"] = str(nums[0])
                if len(nums) > 1:
                    cur["controvalore"] = str(nums[-1])
        if cur and cur.get("nome"):
            fondi.append(cur)

    # Deduplicate by ISIN / nome
    seen = set()
    unique = []
    for f in fondi:
        key = f.get("isin") or f.get("nome", "")[:30]
        if key and key not in seen:
            seen.add(key)
            unique.append(f)

    return unique


def detect_fund_in_form(b: bytes) -> str:
    text = pdf_text(b).upper()
    isin = re.search(r"\b([A-Z]{2}[A-Z0-9]{10})\b", text)
    if isin:
        return isin.group(1)
    for kw in ["AZIMUT", "AZ FUND", "GESTIONE", "BILANCIATO",
               "OBBLIGAZIONARIO", "AZIONARIO", "FLESSIBILE"]:
        if kw in text:
            for line in text.split("\n"):
                if kw in line and len(line.strip()) > 5:
                    return line.strip()[:60]
    return ""


def build_values(ana: dict, fondo) -> dict:
    nome_completo = f"{ana.get('cognome','')} {ana.get('nome','')}".strip()
    v = {
        **{k: ana.get(k, "") for k in _ANA_KEYS},
        "cognome_nome":   nome_completo,
        "nome_cognome":   nome_completo,
        "intestatario":   nome_completo,
        "cf":             ana.get("codice_fiscale", ""),
        "residenza":      ana.get("indirizzo", ""),
        "documento":      ana.get("numero_documento", ""),
    }
    if fondo:
        v.update({
            "nome_fondo":     fondo.get("nome", ""),
            "fondo":          fondo.get("nome", ""),
            "isin":           fondo.get("isin", ""),
            "quote":          str(fondo.get("quote", "")),
            "numero_quote":   str(fondo.get("quote", "")),
            "controvalore":   str(fondo.get("controvalore", "")),
            "importo":        str(fondo.get("controvalore", "")),
        })
    return v


def smart_fill(b: bytes, values: dict) -> tuple[bytes, str]:
    """Fill a PDF. Returns (filled_bytes, method_description)."""
    fields = pdf_fields(b)
    if fields:
        mapped = {}
        for fname in fields:
            fl = fname.lower()
            for k, v in values.items():
                if k.lower() in fl or fl in k.lower():
                    mapped[fname] = v
                    break
        # also try direct name match
        for fname in fields:
            if fname not in mapped and fname.lower() in values:
                mapped[fname] = values[fname.lower()]
        filled = fill_acroform(b, mapped)
        n_filled = len(mapped)
        return filled, f"AcroForm — {n_filled}/{len(fields)} campi compilati"
    else:
        filled = overlay_data_page(b, values)
        return filled, "Nessun campo AcroForm rilevato — allegato foglio dati"


# ─────────────────────────────────────────────────────────────
# SIDEBAR — UPLOAD TOOLBAR
# ─────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📂 Caricamento Documenti")
    st.caption("Carica i file nei rispettivi slot e i dati verranno estratti automaticamente.")
    st.divider()

    st.markdown("### 1️⃣ Documento d'Identità")
    id_up = st.file_uploader(
        "Carta d'identità / Passaporto / Patente",
        type=["pdf", "jpg", "jpeg", "png"],
        key="id_up",
        help="PDF (o immagine) del documento d'identità del cliente",
    )

    st.markdown("### 2️⃣ Posizione Azimut")
    az_up = st.file_uploader(
        "Estratto conto / Posizione fondi Azimut",
        type=["pdf"],
        key="az_up",
        help="PDF della posizione Azimut con i fondi e le quantità",
    )

    st.markdown("### 3️⃣ – 7️⃣ Moduli da Compilare")
    form_ups = []
    for i in range(5):
        fu = st.file_uploader(
            f"Modulo {i + 1}",
            type=["pdf"],
            key=f"fu_{i}",
            help=f"PDF del modulo successione {i + 1}",
        )
        form_ups.append(fu)

    st.divider()
    if st.button("🔄 Azzera sessione", use_container_width=True):
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.rerun()

# ─────────────────────────────────────────────────────────────
# HELPER: aggiorna anagrafica E i widget key contemporaneamente
# ─────────────────────────────────────────────────────────────
def _set_anagrafica(data: dict):
    """Write extracted values to both anagrafica dict and widget session-state keys."""
    for k, v in data.items():
        if k in _ANA_KEYS:
            st.session_state.anagrafica[k] = v
            wk = _WIDGET_KEY.get(k)
            if wk:
                st.session_state[wk] = v


# ─────────────────────────────────────────────────────────────
# CACHE UPLOADED BYTES & AUTO-EXTRACT
# ─────────────────────────────────────────────────────────────

# Identity document
if id_up:
    if st.session_state.id_name != id_up.name:
        st.session_state.id_bytes = id_up.read()
        st.session_state.id_name  = id_up.name
        if id_up.type == "application/pdf":
            with st.spinner("Estrazione dati anagrafici…"):
                raw_text = pdf_text(st.session_state.id_bytes)

                # Se il PDF non ha testo (scansione), prova OCR
                if not raw_text.strip():
                    if _OCR_AVAILABLE:
                        with st.spinner("PDF scansionato — avvio OCR con Tesseract…"):
                            raw_text = ocr_pdf(st.session_state.id_bytes)
                        if raw_text.strip():
                            st.session_state["_id_scanned"] = False
                        else:
                            st.session_state["_id_scanned"] = True
                    else:
                        st.session_state["_id_scanned"] = True

                if not raw_text.strip():
                    st.toast("⚠️ Impossibile estrarre testo — inserisci i dati manualmente", icon="⚠️")
                else:
                    st.session_state["_id_scanned"] = False
                    extracted = parse_identity(st.session_state.id_bytes)
                    # parse_identity usa pdfplumber internamente; se era OCR, passa il testo diretto
                    if not extracted and raw_text.strip():
                        from io import StringIO
                        extracted = _parse_identity_from_text(raw_text)
                    if extracted:
                        _set_anagrafica(extracted)
                        st.toast(f"✅ Estratti {len(extracted)} campi anagrafici", icon="🪪")
                    else:
                        st.toast("⚠️ Dati non riconosciuti — verifica e correggi manualmente", icon="⚠️")
        else:
            # Immagine JPG/PNG: OCR diretto
            if _OCR_AVAILABLE:
                with st.spinner("Lettura immagine con OCR…"):
                    try:
                        from PIL import Image
                        img = Image.open(io.BytesIO(st.session_state.id_bytes))
                        raw_text = pytesseract.image_to_string(img, lang="ita+eng")
                        extracted = _parse_identity_from_text(raw_text)
                        if extracted:
                            _set_anagrafica(extracted)
                            st.session_state["_id_scanned"] = False
                            st.toast(f"✅ Estratti {len(extracted)} campi dall'immagine", icon="🪪")
                        else:
                            st.session_state["_id_scanned"] = True
                            st.toast("⚠️ OCR completato ma dati non riconosciuti — inserisci manualmente", icon="⚠️")
                    except Exception:
                        st.session_state["_id_scanned"] = True
                        st.toast("⚠️ Errore OCR — inserisci i dati manualmente", icon="⚠️")
            else:
                st.session_state["_id_scanned"] = True
                st.toast("Immagine caricata — inserisci i dati anagrafici manualmente", icon="ℹ️")

# Azimut position
if az_up:
    if st.session_state.az_name != az_up.name:
        st.session_state.az_bytes = az_up.read()
        st.session_state.az_name  = az_up.name
        with st.spinner("Estrazione fondi Azimut…"):
            fondi = parse_azimut(st.session_state.az_bytes)
            st.session_state.fondi = fondi
            if fondi:
                st.toast(f"✅ Trovati {len(fondi)} fondi", icon="💼")
            else:
                st.toast("⚠️ Nessun fondo identificato — inserisci manualmente", icon="⚠️")

# Form PDFs
for i, fu in enumerate(form_ups):
    if fu and st.session_state[f"fn_{i}"] != fu.name:
        st.session_state[f"fb_{i}"] = fu.read()
        st.session_state[f"fn_{i}"] = fu.name
        # Invalidate previous compiled version
        if i in st.session_state.filled:
            del st.session_state.filled[i]

# ─────────────────────────────────────────────────────────────
# MAIN HEADER
# ─────────────────────────────────────────────────────────────
st.markdown(
    '<div class="main-header">📋 Segreteria Sassuolo — Compilazione Moduli Successione</div>',
    unsafe_allow_html=True,
)

loaded_id  = st.session_state.id_bytes is not None
loaded_az  = st.session_state.az_bytes is not None
n_forms    = sum(1 for i in range(5) if st.session_state[f"fb_{i}"] is not None)

c1, c2, c3, c4 = st.columns(4)
c1.metric("Documento identità", "✅ Caricato" if loaded_id else "⬜ Mancante")
c2.metric("Posizione Azimut",   "✅ Caricato" if loaded_az else "⬜ Mancante")
c3.metric("Fondi rilevati",     len(st.session_state.fondi))
c4.metric("Moduli caricati",    f"{n_forms} / 5")

st.divider()

# ─────────────────────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────────────────────
tab_ana, tab_fondi, tab_moduli, tab_dl = st.tabs([
    "👤 Dati Anagrafici",
    "💼 Fondi Azimut",
    "📝 Compilazione Moduli",
    "⬇️ Download",
])

# ══════════════════════════════════════════════════════════════
# TAB 1 — ANAGRAFICA
# ══════════════════════════════════════════════════════════════
with tab_ana:
    st.markdown('<div class="section-title">Dati estratti dal documento d\'identità</div>',
                unsafe_allow_html=True)

    if st.session_state.get("_id_scanned"):
        st.warning(
            "⚠️ Il PDF caricato è una scansione (immagine) — il testo non è estraibile automaticamente. "
            "Compila i campi manualmente oppure carica un PDF testuale."
        )
    else:
        st.caption("Verifica e correggi i campi se necessario prima di compilare i moduli.")

    # Legge i valori dal session_state dei widget (se già impostati) oppure da anagrafica
    def _val(wk, ana_k):
        return st.session_state.get(wk) or st.session_state.anagrafica.get(ana_k, "")

    ana = st.session_state.anagrafica
    r1c1, r1c2, r1c3 = st.columns(3)
    with r1c1:
        ana["cognome"]        = st.text_input("Cognome",          value=_val("a_cog","cognome"),       key="a_cog")
        ana["data_nascita"]   = st.text_input("Data di nascita",  value=_val("a_dn","data_nascita"),   key="a_dn")
        ana["codice_fiscale"] = st.text_input("Codice Fiscale",   value=_val("a_cf","codice_fiscale"), key="a_cf")
    with r1c2:
        ana["nome"]           = st.text_input("Nome",             value=_val("a_nom","nome"),          key="a_nom")
        ana["luogo_nascita"]  = st.text_input("Luogo di nascita", value=_val("a_ln","luogo_nascita"),  key="a_ln")
        ana["sesso"]          = st.text_input("Sesso (M/F)",      value=_val("a_ses","sesso"),         key="a_ses", max_chars=1)
    with r1c3:
        ana["indirizzo"]      = st.text_input("Indirizzo",        value=_val("a_ind","indirizzo"),     key="a_ind")
        ana["comune"]         = st.text_input("Comune",           value=_val("a_com","comune"),        key="a_com")
        ana["cap"]            = st.text_input("CAP",              value=_val("a_cap","cap"),           key="a_cap", max_chars=5)

    st.markdown('<div class="section-title">Documento</div>', unsafe_allow_html=True)
    r2c1, r2c2, r2c3 = st.columns(3)
    with r2c1:
        ana["numero_documento"] = st.text_input("N° Documento",  value=_val("a_ndoc","numero_documento"), key="a_ndoc")
    with r2c2:
        ana["data_rilascio"]    = st.text_input("Data rilascio", value=_val("a_drl","data_rilascio"),     key="a_drl")
        ana["ente_rilascio"]    = st.text_input("Ente rilascio", value=_val("a_ent","ente_rilascio"),     key="a_ent")
    with r2c3:
        ana["data_scadenza"]    = st.text_input("Data scadenza", value=_val("a_dsc","data_scadenza"),     key="a_dsc")

# ══════════════════════════════════════════════════════════════
# TAB 2 — FONDI AZIMUT
# ══════════════════════════════════════════════════════════════
with tab_fondi:
    st.markdown('<div class="section-title">Fondi estratti dalla posizione Azimut</div>',
                unsafe_allow_html=True)
    st.caption("Verifica i dati. Puoi aggiungere fondi manualmente o modificare quelli estratti.")

    if st.session_state.fondi:
        # Editable dataframe
        df = pd.DataFrame(st.session_state.fondi,
                          columns=["nome", "isin", "quote", "controvalore"])
        df.columns = ["Nome fondo", "ISIN", "N° Quote", "Controvalore (€)"]
        edited = st.data_editor(
            df,
            use_container_width=True,
            num_rows="dynamic",
            key="fondi_editor",
        )
        # Sync back
        st.session_state.fondi = edited.rename(columns={
            "Nome fondo": "nome", "ISIN": "isin",
            "N° Quote": "quote", "Controvalore (€)": "controvalore"
        }).to_dict("records")
    else:
        st.info("Nessun fondo estratto. Aggiungili manualmente qui sotto.")
        if st.button("➕ Aggiungi riga fondo"):
            st.session_state.fondi.append({"nome": "", "isin": "", "quote": "", "controvalore": ""})
            st.rerun()

# ══════════════════════════════════════════════════════════════
# TAB 3 — COMPILAZIONE MODULI
# ══════════════════════════════════════════════════════════════
with tab_moduli:
    st.markdown('<div class="section-title">Associa ciascun modulo al fondo corrispondente e compila</div>',
                unsafe_allow_html=True)

    active = [(i, st.session_state[f"fn_{i}"], st.session_state[f"fb_{i}"])
              for i in range(5) if st.session_state[f"fb_{i}"] is not None]

    if not active:
        st.info("Carica i moduli PDF dalla barra laterale (slot 3–7) per iniziare.")
    else:
        fund_labels = ["— nessun fondo —"] + [
            f.get("nome", f"Fondo {j+1}")[:50]
            for j, f in enumerate(st.session_state.fondi)
        ]

        for idx, fname, fbytes in active:
            with st.expander(f"**Modulo {idx+1}** — {fname}", expanded=True):
                fields = pdf_fields(fbytes)
                detected_fund = detect_fund_in_form(fbytes)

                mc1, mc2 = st.columns([2, 1])
                with mc1:
                    if fields:
                        st.markdown(f'<span class="badge-ok">✅ {len(fields)} campi AcroForm rilevati</span>',
                                    unsafe_allow_html=True)
                        with st.expander("Elenco campi", expanded=False):
                            for fn, fv in list(fields.items())[:30]:
                                st.text(f"• {fn}")
                            if len(fields) > 30:
                                st.text(f"… e altri {len(fields)-30}")
                    else:
                        st.markdown('<span class="badge-warn">⚠️ Nessun campo AcroForm — verrà allegato foglio dati</span>',
                                    unsafe_allow_html=True)
                    if detected_fund:
                        st.info(f"🔍 Fondo rilevato nel modulo: **{detected_fund}**")

                with mc2:
                    sel = st.selectbox(
                        "Fondo da associare",
                        options=range(len(fund_labels)),
                        format_func=lambda x: fund_labels[x],
                        key=f"sel_{idx}",
                    )

                if st.button(f"📝 Compila Modulo {idx+1}", key=f"btn_{idx}", use_container_width=True):
                    fondo = st.session_state.fondi[sel - 1] if sel > 0 else None
                    values = build_values(st.session_state.anagrafica, fondo)
                    filled_b, method = smart_fill(fbytes, values)
                    st.session_state.filled[idx] = {
                        "bytes":  filled_b,
                        "name":   fname,
                        "values": values,
                        "method": method,
                    }
                    st.success(f"✅ {method}")

                if idx in st.session_state.filled:
                    info = st.session_state.filled[idx]
                    st.caption(f"Ultimo metodo: {info['method']}")
                    v = info["values"]
                    st.markdown(
                        f"**{v.get('cognome','')} {v.get('nome','')}** &nbsp;|&nbsp; "
                        f"CF: `{v.get('codice_fiscale','')}` &nbsp;|&nbsp; "
                        f"Fondo: {v.get('nome_fondo','—')} &nbsp;|&nbsp; "
                        f"Quote: {v.get('quote','—')} &nbsp;|&nbsp; "
                        f"Ctv: {v.get('controvalore','—')} €"
                    )

        st.divider()
        if st.button("🚀 Compila TUTTI i moduli caricati", type="primary",
                     use_container_width=True):
            n_ok = 0
            for idx, fname, fbytes in active:
                if idx in st.session_state.filled:
                    continue
                sel = st.session_state.get(f"sel_{idx}", 0)
                fondo = st.session_state.fondi[sel - 1] if sel > 0 else None
                values = build_values(st.session_state.anagrafica, fondo)
                filled_b, method = smart_fill(fbytes, values)
                st.session_state.filled[idx] = {
                    "bytes": filled_b, "name": fname,
                    "values": values, "method": method,
                }
                n_ok += 1
            st.success(f"✅ Compilati {n_ok} moduli (più {len(st.session_state.filled)-n_ok} già compilati)")
            st.rerun()

# ══════════════════════════════════════════════════════════════
# TAB 4 — DOWNLOAD
# ══════════════════════════════════════════════════════════════
with tab_dl:
    st.markdown('<div class="section-title">Scarica i moduli compilati</div>',
                unsafe_allow_html=True)

    if not st.session_state.filled:
        st.info("Nessun modulo compilato. Vai alla tab 'Compilazione Moduli'.")
    else:
        ana = st.session_state.anagrafica
        cliente = f"{ana.get('cognome','')}_{ana.get('nome','')}".strip("_") or "cliente"

        # Riepilogo
        st.markdown(f"""
**Cliente:** {ana.get('cognome','')} {ana.get('nome','')}
**Codice Fiscale:** {ana.get('codice_fiscale','')}
**Data di nascita:** {ana.get('data_nascita','')} — {ana.get('luogo_nascita','')}
**Documento:** {ana.get('numero_documento','')} &nbsp; rilasciato il {ana.get('data_rilascio','')} &nbsp; scade il {ana.get('data_scadenza','')}
        """)

        if st.session_state.fondi:
            df_dl = pd.DataFrame(st.session_state.fondi)
            df_dl.columns = [c.replace("nome","Nome fondo").replace("isin","ISIN")
                              .replace("quote","N° Quote").replace("controvalore","Controvalore €")
                              for c in df_dl.columns]
            st.dataframe(df_dl, use_container_width=True, hide_index=True)

        st.divider()
        st.markdown("#### Download singoli")
        for idx, info in sorted(st.session_state.filled.items()):
            out_name = info["name"].replace(".pdf", "_compilato.pdf")
            st.download_button(
                label=f"⬇️ {out_name}",
                data=info["bytes"],
                file_name=out_name,
                mime="application/pdf",
                key=f"dl_{idx}",
                use_container_width=True,
            )

        st.divider()
        st.markdown("#### Download ZIP — tutti i moduli")
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for idx, info in st.session_state.filled.items():
                zf.writestr(info["name"].replace(".pdf", "_compilato.pdf"), info["bytes"])
        st.download_button(
            label=f"📦 Scarica ZIP ({cliente})",
            data=zip_buf.getvalue(),
            file_name=f"moduli_successione_{cliente}.zip",
            mime="application/zip",
            type="primary",
            use_container_width=True,
        )
