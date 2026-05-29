"""
Test: ispeziona la pagina Quantalys Historique con Playwright
e trova i selettori CSS per i grafici SVG.
Salva screenshot completo + screenshot dei 6 riquadri.

Uso:
    pip install playwright
    playwright install chromium
    python _test_qtl_chart.py
"""
import time
from pathlib import Path

try:
    from playwright.sync_api import sync_playwright
except ImportError:
    print("Installa playwright: pip install playwright && playwright install chromium")
    raise

URL = "https://www.quantalys.it/Fonds/Historique/825616"
OUT_FULL   = Path("data/_qtl_full_page.png")
OUT_CHARTS = Path("data/_qtl_charts_section.png")
OUT_MAIN   = Path("data/_qtl_main_chart.png")
OUT_6BOX   = Path("data/_qtl_6boxes.png")

def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(viewport={"width": 1400, "height": 900})

        print(f"Carico: {URL}")
        page.goto(URL, wait_until="networkidle", timeout=30_000)
        print("Pagina caricata — attendo rendering grafici...")
        page.wait_for_timeout(5_000)   # attende JS charts (SVG)

        # ── Screenshot pagina intera ──────────────────────────────────────
        OUT_FULL.parent.mkdir(exist_ok=True)
        page.screenshot(path=str(OUT_FULL), full_page=True)
        print(f"Screenshot pagina intera: {OUT_FULL}")

        # ── Cerca tutti gli SVG ──────────────────────────────────────────
        svgs = page.evaluate("""() => {
            const els = document.querySelectorAll('svg');
            return [...els].map((el, i) => {
                const r = el.getBoundingClientRect();
                const p1 = el.parentElement;
                const p2 = p1 ? p1.parentElement : null;
                const p3 = p2 ? p2.parentElement : null;
                return {
                    i,
                    width:  Math.round(r.width),
                    height: Math.round(r.height),
                    top:    Math.round(r.top + window.scrollY),
                    left:   Math.round(r.left),
                    svgClass: [...el.classList].join(' '),
                    p1_tag:   p1 ? p1.tagName : '',
                    p1_class: p1 ? [...p1.classList].join(' ') : '',
                    p1_id:    p1 ? p1.id : '',
                    p2_tag:   p2 ? p2.tagName : '',
                    p2_class: p2 ? [...p2.classList].join(' ') : '',
                    p2_id:    p2 ? p2.id : '',
                    p3_tag:   p3 ? p3.tagName : '',
                    p3_class: p3 ? [...p3.classList].join(' ') : '',
                    p3_id:    p3 ? p3.id : '',
                };
            });
        }""")

        print(f"\n=== SVG trovati: {len(svgs)} ===")
        for s in svgs:
            if s['width'] > 100:   # ignora SVG icone piccole
                print(f"  [{s['i']}] {s['width']}x{s['height']} top={s['top']} left={s['left']}")
                print(f"       svg.class='{s['svgClass']}'")
                print(f"       p1: <{s['p1_tag']}> id='{s['p1_id']}' class='{s['p1_class']}'")
                print(f"       p2: <{s['p2_tag']}> id='{s['p2_id']}' class='{s['p2_class']}'")
                print(f"       p3: <{s['p3_tag']}> id='{s['p3_id']}' class='{s['p3_class']}'")

        # ── Cerca sezione con i 6 riquadri (titolo "Rendimenti annuali" o simile) ──
        sections = page.evaluate("""() => {
            const results = [];
            const allDivs = document.querySelectorAll('div, section');
            for (const el of allDivs) {
                const svgCount = el.querySelectorAll('svg').length;
                if (svgCount >= 3) {
                    const r = el.getBoundingClientRect();
                    if (r.width > 300) {
                        results.push({
                            tag:      el.tagName,
                            id:       el.id,
                            cls:      [...el.classList].join(' '),
                            svgCount,
                            width:    Math.round(r.width),
                            height:   Math.round(r.height),
                            top:      Math.round(r.top + window.scrollY),
                            textSnip: (el.innerText || '').slice(0, 120).replace(/\\n/g,' ')
                        });
                    }
                    if (results.length >= 8) break;
                }
            }
            return results;
        }""")

        print(f"\n=== Contenitori con >=3 SVG ===")
        for s in sections:
            print(f"  <{s['tag']}> id='{s['id']}' class='{s['cls']}'")
            print(f"    svgCount={s['svgCount']} {s['width']}x{s['height']} top={s['top']}")
            print(f"    testo: {s['textSnip']}")

        # ── Cerca sezione con i 6 riquadri per testo "Rendimento" / "Volatilità" ──
        box6 = page.evaluate("""() => {
            const results = [];
            const allEls = document.querySelectorAll('div, section, article');
            for (const el of allEls) {
                const txt = (el.innerText || '').toLowerCase();
                const svgCount = el.querySelectorAll('svg').length;
                if (svgCount >= 4 && svgCount <= 12
                    && (txt.includes('rendimento') || txt.includes('volatil') || txt.includes('drawdown'))) {
                    const r = el.getBoundingClientRect();
                    if (r.width > 400 && r.height > 100) {
                        results.push({
                            tag:      el.tagName,
                            id:       el.id,
                            cls:      [...el.classList].join(' '),
                            svgCount,
                            width:    Math.round(r.width),
                            height:   Math.round(r.height),
                            top:      Math.round(r.top + window.scrollY),
                            textSnip: (el.innerText || '').slice(0, 200).replace(/\\n/g,' ')
                        });
                        if (results.length >= 6) break;
                    }
                }
            }
            return results;
        }""")

        print(f"\n=== Candidati sezione 6 riquadri ===")
        for s in box6:
            print(f"  <{s['tag']}> id='{s['id']}' class='{s['cls']}'")
            print(f"    svgCount={s['svgCount']} {s['width']}x{s['height']} top={s['top']}")
            print(f"    testo: {s['textSnip'][:120]}")

        # ── Cerca tutti i div con classe che contiene 'chart' o 'graph' ──
        chart_divs = page.evaluate("""() => {
            const results = [];
            const allEls = document.querySelectorAll('[class*="chart"], [class*="Chart"], [class*="graph"], [id*="chart"], [id*="Chart"]');
            for (const el of allEls) {
                const r = el.getBoundingClientRect();
                if (r.width > 200 && r.height > 100) {
                    results.push({
                        tag:   el.tagName,
                        id:    el.id,
                        cls:   [...el.classList].join(' '),
                        width: Math.round(r.width),
                        height:Math.round(r.height),
                        top:   Math.round(r.top + window.scrollY),
                    });
                    if (results.length >= 12) break;
                }
            }
            return results;
        }""")

        print(f"\n=== Div con 'chart'/'graph' nella classe/id ===")
        for c in chart_divs:
            print(f"  <{c['tag']}> id='{c['id']}' class='{c['cls']}' {c['width']}x{c['height']} top={c['top']}")

        # ── Struttura HTML grezza dei contenitori più grandi ─────────────
        structure = page.evaluate("""() => {
            // Cerca il main content area (dove stanno i grafici)
            const mainContent = document.querySelector('main, #main, .main-content, .content, #content');
            if (!mainContent) return "no main found";
            // Prendi la struttura a 2 livelli
            const children = [...mainContent.children].map(c => ({
                tag: c.tagName,
                id: c.id,
                cls: [...c.classList].join(' '),
                svgs: c.querySelectorAll('svg').length,
                text: (c.innerText||'').slice(0,60)
            }));
            return children;
        }""")
        print(f"\n=== Struttura main content ===")
        if isinstance(structure, list):
            for ch in structure:
                print(f"  <{ch['tag']}> id='{ch['id']}' class='{ch['cls']}' svgs={ch['svgs']} | {ch['text'][:60]}")
        else:
            print(f"  {structure}")

        # ── Screenshot region: scroll e cattura area grafici ────────────
        # Prova a screenshottare il primo candidato sezione 6 riquadri
        if box6:
            best = min(box6, key=lambda x: abs(x['svgCount'] - 6))
            sel = None
            if best['id']:
                sel = f"#{best['id']}"
            elif best['cls']:
                first_cls = best['cls'].split()[0]
                sel = f".{first_cls}"
            if sel:
                try:
                    elem = page.query_selector(sel)
                    if elem:
                        elem.screenshot(path=str(OUT_6BOX))
                        print(f"\nScreenshot 6 riquadri: {OUT_6BOX}  (sel='{sel}')")
                except Exception as e:
                    print(f"\nScreenshot 6 riquadri fallito ({sel}): {e}")

        # ── Screenshot sezione SVG principale (primo SVG grande) ─────────
        big_svgs = [s for s in svgs if s['width'] > 400 and s['height'] > 200]
        if big_svgs:
            best_svg = big_svgs[0]
            parent_sel = None
            if best_svg['p1_id']:
                parent_sel = f"#{best_svg['p1_id']}"
            elif best_svg['p1_class']:
                parent_sel = f".{best_svg['p1_class'].split()[0]}"
            if parent_sel:
                try:
                    elem = page.query_selector(parent_sel)
                    if elem:
                        elem.screenshot(path=str(OUT_MAIN))
                        print(f"Screenshot grafico principale: {OUT_MAIN}  (sel='{parent_sel}')")
                except Exception as e:
                    print(f"Screenshot grafico principale fallito: {e}")

        browser.close()
        print("\nFatto.")

if __name__ == "__main__":
    main()
