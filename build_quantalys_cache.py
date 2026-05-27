"""
build_quantalys_cache.py
------------------------
Discovers Quantalys URLs for every fund ISIN found in data/excel_cache.json
by querying DuckDuckGo for 'site:quantalys.it "{ISIN}"'.

Saves results to data/quantalys_cache.json as:
  { "ISIN": "https://www.quantalys.it/Fonds/{id}", ... }

Run once (or periodically to refresh):
  python build_quantalys_cache.py

Uses 2-second delay between requests to be polite to DuckDuckGo.
"""

import json
import re
import time
from pathlib import Path

from ddgs import DDGS

# ── Paths ────────────────────────────────────────────────────────────────────
BASE_DIR   = Path(__file__).parent
CACHE_FILE = BASE_DIR / "data" / "quantalys_cache.json"
EXCEL_FILE = BASE_DIR / "data" / "excel_cache.json"

# Pattern to extract a valid Quantalys fund URL
_QT_RE = re.compile(
    r"quantalys\.it/[Ff]onds(?:/[A-Za-z/]*)?/(\d+)",
    re.IGNORECASE,
)


def canonical_url(numeric_id: str) -> str:
    return f"https://www.quantalys.it/Fonds/{numeric_id}"


def search_quantalys(isin: str) -> str:
    """Return the canonical Quantalys URL for an ISIN, or '' if not found."""
    try:
        with DDGS() as ddgs:
            results = list(ddgs.text(f'site:quantalys.it "{isin}"', max_results=5))
        for r in results:
            href = r.get("href", "")
            m = _QT_RE.search(href)
            if m:
                return canonical_url(m.group(1))
        return ""
    except Exception as exc:
        print(f"  [!] Error for {isin}: {exc}")
        return ""


def load_isins() -> list[tuple[str, str]]:
    """Return list of (fund_name, isin) from excel_cache FIDA sheet."""
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
    # Load existing cache (resume-friendly)
    existing: dict[str, str] = {}
    if CACHE_FILE.exists():
        with open(CACHE_FILE, encoding="utf-8") as f:
            existing = json.load(f)
        print(f"Loaded existing cache: {len(existing)} entries")

    pairs = load_isins()
    print(f"Total ISINs to process: {len(pairs)}")

    todo = [(n, isin) for n, isin in pairs if isin not in existing]
    print(f"Remaining to search: {len(todo)}")

    if not todo:
        print("Nothing to do — cache is complete.")
        return

    found = 0
    not_found = 0

    try:
        for i, (nome, isin) in enumerate(todo, 1):
            print(f"[{i:3d}/{len(todo)}] {isin}  {nome[:50]}", end=" ... ", flush=True)
            url = search_quantalys(isin)
            if url:
                existing[isin] = url
                print(f"OK {url}")
                found += 1
            else:
                existing[isin] = ""   # mark as searched-but-not-found
                print("-- not found")
                not_found += 1

            # Save incrementally every 20 entries
            if i % 20 == 0:
                with open(CACHE_FILE, "w", encoding="utf-8") as f:
                    json.dump(existing, f, ensure_ascii=False, indent=2)
                print(f"  >> Cache saved ({i} processed, {found} found so far)")

            # Polite delay
            time.sleep(2.2)

    except KeyboardInterrupt:
        print("\n[!] Interrupted by user")

    # Final save
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(existing, f, ensure_ascii=False, indent=2)

    total_with_url = sum(1 for v in existing.values() if v)
    print(f"\nDone. Found: {found}  Not found: {not_found}")
    print(f"Cache total: {len(existing)} entries, {total_with_url} with URLs")
    print(f"Saved to {CACHE_FILE}")


if __name__ == "__main__":
    main()
