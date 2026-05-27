"""
build_quantalys_clean.py
------------------------
Ricostruisce data/quantalys_cache.json in modo affidabile:

Per ogni ISIN:
  1. Se esiste già un URL nel cache: verifica che l'ISIN sia presente
     nella pagina Quantalys (altrimenti l'URL e' sbagliato e va rifatto).
  2. Se l'URL manca o e' sbagliato: cerca su DuckDuckGo
     "site:quantalys.it {ISIN}" e verifica il primo risultato.
  3. Se ancora non trovato: cerca per nome-concetto del fondo.

Salva solo URL verificati (ISIN presente nella pagina).
"""

import json
import re
import time
from pathlib import Path

import requests
from ddgs import DDGS

BASE_DIR   = Path(__file__).parent
CACHE_FILE = BASE_DIR / "data" / "quantalys_cache.json"
EXCEL_FILE = BASE_DIR / "data" / "excel_cache.json"

_QT_RE = re.compile(r"quantalys\.it/[Ff]onds(?:/[A-Za-z/]*)?/(\d+)", re.I)

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "it-IT,it;q=0.9",
}

session = requests.Session()
session.headers.update(HEADERS)


# ── helpers ──────────────────────────────────────────────────────────────────

def canonical(numeric_id: str) -> str:
    return f"https://www.quantalys.it/Fonds/{numeric_id}"


def isin_on_page(url: str, isin: str) -> bool:
    """True se la pagina Quantalys contiene l'ISIN (conferma che e' il fondo giusto)."""
    try:
        r = session.get(url, timeout=12)
        return isin.upper() in r.text.upper()
    except Exception:
        return False


def ddg_search(query: str, max_results: int = 5) -> list[str]:
    """Restituisce lista di URL da DuckDuckGo."""
    try:
        with DDGS() as ddgs:
            hits = list(ddgs.text(query, max_results=max_results))
        return [h.get("href", "") for h in hits if h.get("href")]
    except Exception:
        return []


def find_url_for_isin(isin: str, nome: str) -> str:
    """
    Strategia 1: DDG con ISIN quotato.
    Strategia 2: DDG con nome-concetto.
    Per ogni candidato, verifica che l'ISIN sia nella pagina.
    Ritorna URL verificato o '' se non trovato.
    """
    # Strategia 1 — ISIN
    candidates = ddg_search(f'site:quantalys.it "{isin}"')
    for href in candidates:
        m = _QT_RE.search(href)
        if m:
            url = canonical(m.group(1))
            time.sleep(0.8)
            if isin_on_page(url, isin):
                return url

    time.sleep(1.5)

    # Strategia 2 — nome-concetto
    concept = _strip_class(nome)
    if concept:
        candidates2 = ddg_search(f'site:quantalys.it "{concept}"')
        for href in candidates2:
            m = _QT_RE.search(href)
            if m:
                url = canonical(m.group(1))
                time.sleep(0.8)
                if isin_on_page(url, isin):
                    return url

    return ""


def _strip_class(name: str) -> str:
    """Rimuove prefisso AZ e suffisso classe (A Cap EUR ecc.)."""
    n = re.sub(
        r"^AZ\s+(?:F\.\d+\s+\w+[\. ]+|Fund\s+\d+\s*[-]+\s*|\w+\s*[-]+\s*)",
        "", name.strip(), flags=re.I
    ).strip()
    n = re.sub(r"\s+[A-Z](?:-[A-Z0-9]+)?\s+(?:Cap|Dis|Acc|Inc)\b.*", "", n, flags=re.I).strip()
    return n


# ── main ─────────────────────────────────────────────────────────────────────

def load_isins() -> list[tuple[str, str]]:
    with open(EXCEL_FILE, encoding="utf-8") as f:
        data = json.load(f)
    fida = data.get("FIDA", [])
    seen: set[str] = set()
    pairs = []
    for entry in fida:
        isin = str(entry.get("isin") or "").strip()
        nome = str(entry.get("nome") or "").strip()
        if isin and isin not in seen:
            seen.add(isin)
            pairs.append((isin, nome))
    return pairs


def main():
    # Carica cache esistente
    cache: dict[str, str] = {}
    if CACHE_FILE.exists():
        with open(CACHE_FILE, encoding="utf-8") as f:
            cache = json.load(f)
    print(f"Cache attuale: {len(cache)} voci, {sum(1 for v in cache.values() if v)} con URL")

    pairs = load_isins()
    print(f"ISIN totali da processare: {len(pairs)}")

    # Separa in categorie
    to_verify  = [(i, n) for i, n in pairs if cache.get(i)]       # URL presenti -> da verificare
    to_search  = [(i, n) for i, n in pairs if not cache.get(i)]   # assenti o vuoti -> da cercare

    print(f"Da verificare: {len(to_verify)}  |  Da cercare: {len(to_search)}")
    print()

    verified_ok = verified_bad = 0

    # ── FASE 1: verifica URL esistenti ───────────────────────────────────────
    print("=== FASE 1: verifica URL esistenti ===")
    bad_isins: list[tuple[str, str]] = []

    for idx, (isin, nome) in enumerate(to_verify, 1):
        url = cache[isin]
        ok = isin_on_page(url, isin)
        if ok:
            verified_ok += 1
            print(f"[{idx:3d}/{len(to_verify)}] OK  {isin} | {nome[:40]}")
        else:
            verified_bad += 1
            cache[isin] = ""    # azzera — da ricercare di nuovo
            bad_isins.append((isin, nome))
            print(f"[{idx:3d}/{len(to_verify)}] BAD {isin} | {nome[:40]} -- URL sbagliata, da rifare")
        time.sleep(0.5)

        if idx % 30 == 0:
            _save(cache)
            print(f"  >> Salvato (verifica: {idx} processati, {verified_bad} sbagliati)")

    print(f"\nVerifica completata: {verified_ok} OK, {verified_bad} sbagliate")
    _save(cache)

    # ── FASE 2: ricerca nuova per assenti + sbagliate ─────────────────────────
    to_find = to_search + bad_isins
    print(f"\n=== FASE 2: ricerca per {len(to_find)} ISIN ===")

    found = not_found = 0
    for idx, (isin, nome) in enumerate(to_find, 1):
        print(f"[{idx:3d}/{len(to_find)}] {isin} | {nome[:40]:40} ... ", end="", flush=True)
        url = find_url_for_isin(isin, nome)
        if url:
            cache[isin] = url
            found += 1
            print(f"OK {url}")
        else:
            cache[isin] = ""
            not_found += 1
            print("-- non trovato")
        time.sleep(2.5)

        if idx % 20 == 0:
            _save(cache)
            total_ok = sum(1 for v in cache.values() if v)
            print(f"  >> Salvato ({idx} processati, totale URL: {total_ok})")

    _save(cache)
    total_ok = sum(1 for v in cache.values() if v)
    print(f"\nCompletato. URL verificate: {total_ok}/{len(cache)}")
    print(f"Fase 1: {verified_ok} confermate, {verified_bad} corrette")
    print(f"Fase 2: {found} trovate, {not_found} non trovate")


def _save(cache: dict):
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    main()
