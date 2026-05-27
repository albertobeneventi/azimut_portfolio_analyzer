"""
build_quantalys_cache2.py  — Secondo passaggio
-----------------------------------------------
Per i fondi il cui ISIN non ha dato risultati nel primo passaggio,
prova a cercare per nome-concetto del fondo (es. "Balanced FoF").
Aggiorna data/quantalys_cache.json.

Eseguire dopo build_quantalys_cache.py.
"""
import json
import re
import time
from pathlib import Path

from ddgs import DDGS

BASE_DIR   = Path(__file__).parent
CACHE_FILE = BASE_DIR / "data" / "quantalys_cache.json"
EXCEL_FILE = BASE_DIR / "data" / "excel_cache.json"

_QT_RE = re.compile(
    r"quantalys\.it/[Ff]onds(?:/[A-Za-z/]*)?/(\d+)",
    re.IGNORECASE,
)


def canonical_url(numeric_id: str) -> str:
    return f"https://www.quantalys.it/Fonds/{numeric_id}"


def strip_class(name: str) -> str:
    """Rimuove prefisso AZ e suffisso di classe per ottenere il nome-concetto."""
    n = name.strip()
    # Rimuove prefisso: "AZ F.1 All. " / "AZ F.1 Eq. " / "AZ Allocation - " ecc.
    n = re.sub(
        r'^AZ\s+(?:F\.\d+\s+\w+[\. ]+|Fund\s+\d+\s*[-]+\s*|\w+\s*[-]+\s*)',
        '', n, flags=re.I
    ).strip()
    # Rimuove suffisso classe: "A Cap EUR", "B-HU Cap EUR Hdg", ecc.
    n = re.sub(
        r'\s+[A-Z](?:-[A-Z0-9]+)?\s+(?:Cap|Dis|Acc|Inc)\b.*',
        '', n, flags=re.I
    ).strip()
    return n


def search_by_name(concept: str, retries: int = 3) -> str:
    """Cerca su Quantalys per nome concetto. Ritorna URL o ''."""
    query = f'site:quantalys.it "{concept}"'
    for attempt in range(retries):
        try:
            with DDGS() as ddgs:
                hits = list(ddgs.text(query, max_results=5))
            for h in hits:
                m = _QT_RE.search(h.get("href", ""))
                if m:
                    return canonical_url(m.group(1))
            return ""   # trovato risultati ma nessun URL Quantalys
        except Exception as exc:
            print(f"  [!] tentativo {attempt+1}/{retries} per '{concept}': {exc}")
            time.sleep(4.0 * (attempt + 1))
    return ""


def main():
    with open(CACHE_FILE, encoding="utf-8") as f:
        cache: dict[str, str] = json.load(f)

    with open(EXCEL_FILE, encoding="utf-8") as f:
        excel = json.load(f)

    fida = excel.get("FIDA", [])
    isin_to_nome = {
        r["isin"]: r["nome"]
        for r in fida
        if r.get("isin") and r.get("nome")
    }

    # Fondi con ISIN in cache ma URL vuota
    missing = {
        isin: nome
        for isin, nome in isin_to_nome.items()
        if cache.get(isin, "X") == ""   # vuoto = cercato ma non trovato
    }
    print(f"Fondi con URL vuota da ricercare per nome: {len(missing)}")

    # Deduplica: stesso concetto potrebbe comparire piu volte (classi diverse)
    concept_done: dict[str, str] = {}   # concept -> url (o 'tried')
    found = updated = 0

    try:
        items = list(missing.items())
        for i, (isin, nome) in enumerate(items, 1):
            concept = strip_class(nome)
            if not concept:
                print(f"[{i:3d}] {isin} -- nome-concetto vuoto, salto")
                continue

            # Se gia cercato questo concetto, riusa il risultato
            if concept in concept_done:
                url = concept_done[concept]
                if url:
                    cache[isin] = url
                    updated += 1
                    print(f"[{i:3d}] {isin} | {nome[:40]:40} -- riuso {url}")
                else:
                    print(f"[{i:3d}] {isin} | {nome[:40]:40} -- concetto gia fallito")
                continue

            print(f"[{i:3d}/{len(items)}] {isin} | {nome[:40]:40} ... ", end="", flush=True)
            url = search_by_name(concept)
            concept_done[concept] = url

            if url:
                cache[isin] = url
                found += 1
                updated += 1
                print(f"OK {url}")
            else:
                print("-- non trovato")

            # Salva ogni 20 ricerche reali
            if len(concept_done) % 20 == 0:
                with open(CACHE_FILE, "w", encoding="utf-8") as f:
                    json.dump(cache, f, ensure_ascii=False, indent=2)
                print(f"  >> Cache salvata ({len(concept_done)} concetti cercati)")

            time.sleep(2.5)

    except KeyboardInterrupt:
        print("\n[!] Interrotto")

    # Salva finale
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)

    total_with_url = sum(1 for v in cache.values() if v)
    print(f"\nNuovi URL trovati: {found}  (aggiornamenti totali: {updated})")
    print(f"Cache totale: {len(cache)} voci, {total_with_url} con URL")
    print(f"Salvato in {CACHE_FILE}")


if __name__ == "__main__":
    main()
