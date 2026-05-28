"""
build_morningstar_cache.py
--------------------------
Discovers Morningstar URLs for every fund ISIN found in data/excel_cache.json
by querying DuckDuckGo for 'site:morningstar.it "{ISIN}"'.

Saves results to data/morningstar_cache.json as:
  { "ISIN": "https://www.morningstar.it/it/funds/snapshot/snapshot.aspx?id=XXXXX", ... }

Run once (or periodically to refresh):
  python build_morningstar_cache.py

Uses 2.5-second delay between requests to be polite to DuckDuckGo.
"""

import json
import re
import time
from pathlib import Path

from ddgs import DDGS

# ── Paths ────────────────────────────────────────────────────────────────────
BASE_DIR   = Path(__file__).parent
CACHE_FILE = BASE_DIR / "data" / "morningstar_cache.json"
EXCEL_FILE = BASE_DIR / "data" / "excel_cache.json"

# Pattern per estrarre ID Morningstar dal URL
_MS_RE = re.compile(
    r'morningstar\.it/it/funds/snapshot/snapshot\.aspx\?id=([A-Z0-9]+)',
    re.IGNORECASE,
)

# ── URL corretti manualmente (DuckDuckGo aveva trovato pagine sbagliate) ──────
# Aggiungere qui ogni ISIN per cui la ricerca automatica non funziona o
# restituisce un URL errato. Questi valori hanno sempre la precedenza.
#
#   "ISIN": "https://www.morningstar.it/it/funds/snapshot/snapshot.aspx?id=<ID>"
#
MANUAL_OVERRIDES: dict[str, str] = {
    "LU3081792221": "https://global.morningstar.com/it/investimenti/fondi/0P0001XRB9/quote",  # AZ F.1 All. Escalator 2030 A Cap EUR
    "LU0738951036": "https://global.morningstar.com/it/investimenti/fondi/0P0001JMJQ/quote",  # AZ F.1 Bd Patriot A Cap EUR
    "LU2951609937": "https://global.morningstar.com/it/investimenti/fondi/0P0001US0L/quote",  # AZ F.1 Bd Target 2029 A Cap EUR
    "LU0346933400": "https://global.morningstar.com/it/investimenti/fondi/0P0000J14M/quote",  # AZ F.1 All. Balanced FoF A Cap EUR
    "LU2637786422": "https://global.morningstar.com/it/investimenti/fondi/0P0001TD5F/quote",  # AZ F.1 All. Potential Income Upside 2030 A Cap EUR
    # Esempio:
    # "LU2168564065": "https://www.morningstar.it/it/funds/snapshot/snapshot.aspx?id=XXXXXXXX",
}


def canonical_url(ms_id: str) -> str:
    return f"https://www.morningstar.it/it/funds/snapshot/snapshot.aspx?id={ms_id}"


def search_morningstar(isin: str) -> str:
    """Restituisce l'URL Morningstar per un ISIN via DuckDuckGo, o '' se non trovato."""
    try:
        with DDGS() as ddgs:
            results = list(ddgs.text(f'site:morningstar.it "{isin}"', max_results=5))
        for r in results:
            href = r.get("href", "")
            m = _MS_RE.search(href)
            if m:
                return canonical_url(m.group(1).upper())
        return ""
    except Exception as exc:
        print(f"  [!] Errore per {isin}: {exc}")
        return ""


def load_isins() -> list[tuple[str, str]]:
    """Ritorna lista (nome_fondo, isin) dalla cache excel (sheet FIDA)."""
    with open(EXCEL_FILE, encoding="utf-8") as f:
        data = json.load(f)
    fida = data.get("FIDA", [])
    pairs = []
    seen: set[str] = set()
    for entry in fida:
        isin = (entry.get("isin") or "").strip()
        nome = (entry.get("nome") or "").strip()
        if isin and isin not in seen:
            seen.add(isin)
            pairs.append((nome, isin))
    return pairs


def main():
    # Carica cache esistente (per riprendere da dove si era interrotti)
    existing: dict[str, str] = {}
    if CACHE_FILE.exists():
        with open(CACHE_FILE, encoding="utf-8") as f:
            existing = json.load(f)
        print(f"Cache esistente caricata: {len(existing)} voci")

    # Applica sempre gli override manuali
    overridden = []
    for isin, url in MANUAL_OVERRIDES.items():
        if existing.get(isin) != url:
            existing[isin] = url
            overridden.append(isin)
    if overridden:
        print(f"Override manuali applicati: {overridden}")

    pairs = load_isins()
    print(f"ISIN totali da processare: {len(pairs)}")

    todo = [(n, isin) for n, isin in pairs if isin not in existing]
    print(f"Da cercare: {len(todo)}\n")

    if not todo:
        print("Niente da fare — cache completa.")
        _print_summary(existing, pairs)
        return

    found = 0
    not_found = 0

    try:
        for i, (nome, isin) in enumerate(todo, 1):
            print(f"[{i:3d}/{len(todo)}] {isin}  {nome[:50]}", end=" ... ", flush=True)

            # Override manuale ha precedenza
            if isin in MANUAL_OVERRIDES:
                url = MANUAL_OVERRIDES[isin]
                print(f"OVERRIDE {url}")
                existing[isin] = url
                found += 1
                continue

            url = search_morningstar(isin)
            if url:
                existing[isin] = url
                print(f"OK {url}")
                found += 1
            else:
                existing[isin] = ""   # segnato come cercato-ma-non-trovato
                print("-- non trovato")
                not_found += 1

            # Salva ogni 10 voci
            if i % 10 == 0:
                with open(CACHE_FILE, "w", encoding="utf-8") as f:
                    json.dump(existing, f, ensure_ascii=False, indent=2)
                print(f"  >> Cache salvata ({i} processati, {found} trovati)")

            time.sleep(2.5)

    except KeyboardInterrupt:
        print("\n[!] Interrotto dall'utente")

    # Salvataggio finale
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(existing, f, ensure_ascii=False, indent=2)

    print()
    _print_summary(existing, pairs)


def _print_summary(existing: dict, pairs: list) -> None:
    total_with_url = sum(1 for v in existing.values() if v)
    not_found_list = [(n, isin) for n, isin in pairs if not existing.get(isin)]
    print(f"Cache totale: {len(existing)} voci, {total_with_url} con URL")
    print(f"Salvata in {CACHE_FILE}")
    if not_found_list:
        print(f"\nNon trovati ({len(not_found_list)}) — da aggiungere manualmente a MANUAL_OVERRIDES:")
        for nome, isin in not_found_list:
            print(f"  {isin}  {nome}")


if __name__ == "__main__":
    main()
