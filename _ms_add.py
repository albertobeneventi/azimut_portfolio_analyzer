"""
Aggiunge un URL Morningstar verificato manualmente.

Uso:
    python _ms_add.py <ISIN> <URL>

Esempio:
    python _ms_add.py LU3081792221 https://www.morningstar.it/it/funds/snapshot/snapshot.aspx?id=F00000T56F

Effetti:
  1. Aggiunge/aggiorna la voce in MANUAL_OVERRIDES in build_morningstar_cache.py
  2. Aggiorna immediatamente data/morningstar_cache.json
"""
import sys, json, re
from pathlib import Path

if len(sys.argv) != 3:
    print("Uso: python _ms_add.py <ISIN> <URL>")
    sys.exit(1)

isin = sys.argv[1].strip().upper()
url  = sys.argv[2].strip()

# Validazione minima
if not re.match(r'^[A-Z]{2}[A-Z0-9]{10}$', isin):
    print(f"ISIN non valido: {isin}")
    sys.exit(1)
if 'morningstar' not in url.lower() and 'morningstar' not in url.lower():
    print(f"URL non sembra Morningstar: {url}")
    sys.exit(1)
# Accetta sia il vecchio formato (morningstar.it/...snapshot.aspx?id=...)
# sia il nuovo (global.morningstar.com/it/investimenti/fondi/.../quote)

BUILD_SCRIPT = Path('build_morningstar_cache.py')
CACHE_FILE   = Path('data/morningstar_cache.json')

# ── 1. Aggiorna MANUAL_OVERRIDES nel build script ────────────────────────────
text = BUILD_SCRIPT.read_text(encoding='utf-8')

# Cerca il blocco MANUAL_OVERRIDES
marker = 'MANUAL_OVERRIDES: dict[str, str] = {'
if marker not in text:
    print("Impossibile trovare MANUAL_OVERRIDES in build_morningstar_cache.py")
    sys.exit(1)

# Controlla se l'ISIN è già presente
isin_pattern = re.compile(rf'"{isin}"\s*:')
if isin_pattern.search(text):
    # Aggiorna URL esistente
    text = re.sub(
        rf'"{isin}"\s*:.*',
        f'"{isin}": "{url}",',
        text
    )
    action = 'aggiornato'
else:
    # Inserisce prima della riga con il commento finale o della chiusura }
    # Trova la riga del commento "# Esempio:" o la riga "}"  che chiude il dict
    close_re = re.compile(r'^(\s*)(# Esempio:|}\s*$)', re.MULTILINE)
    m = close_re.search(text, text.index(marker))
    if m:
        indent  = '    '
        # Trova il nome del fondo dalla fund_cache se disponibile
        fname = ''
        try:
            fc = json.loads(Path('data/fund_cache.json').read_text(encoding='utf-8'))
            for nome, d in (fc.get('fund_data') or {}).items():
                if isinstance(d, dict) and d.get('isin') == isin:
                    fname = nome
                    break
        except Exception:
            pass
        comment = f'  # {fname}' if fname else ''
        new_line = f'{indent}"{isin}": "{url}",{comment}\n'
        text = text[:m.start()] + new_line + text[m.start():]
        action = 'aggiunto'
    else:
        print("Struttura MANUAL_OVERRIDES non riconosciuta")
        sys.exit(1)

BUILD_SCRIPT.write_text(text, encoding='utf-8')

# ── 2. Aggiorna immediatamente morningstar_cache.json ────────────────────────
cache = {}
if CACHE_FILE.exists():
    cache = json.loads(CACHE_FILE.read_text(encoding='utf-8'))
cache[isin] = url
CACHE_FILE.write_text(json.dumps(cache, ensure_ascii=False, indent=2), encoding='utf-8')

print(f"OK — {action} {isin}")
print(f"  URL: {url}")
print(f"  Cache aggiornata ({len(cache)} voci totali)")
print()
print("Ora esegui:")
print("  git add build_morningstar_cache.py data/morningstar_cache.json && git commit -m \"fix: URL Morningstar manuale per " + isin + "\"")
