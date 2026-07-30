#!/usr/bin/env python3
"""Multi-language regression harness for the dashboard.

For each language it: presets the stored preference, loads the workbook through
the real file input, reads back the headline figures, switches language live via
the header picker, and re-reads. Digits must be identical across languages;
words must not be.

Also covers the two generic paths the prompt requires: a workbook with the
Config sheet deleted, and one whose Config names a different currency + locale.

Usage: python3 test_i18n.py <dashboard.html> <workbook.xlsx>
"""
import sys, json, re, pathlib
from playwright.sync_api import sync_playwright

HTML = pathlib.Path(sys.argv[1]).resolve()
WB = pathlib.Path(sys.argv[2]).resolve()
HERE = pathlib.Path(__file__).resolve().parent

src = HTML.read_text()
src = src.replace("https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.0/chart.umd.min.js", "vendor_chart.js")
src = src.replace("https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js", "vendor_xlsx.js")
tmp = HERE / "_test_i18n.html"
tmp.write_text(src)

PROBES = {"verdict": "ovVt", "verdictSub": "ovVs", "safe": "ovSafe", "nw": "ovNW",
          "nl": "ovNL", "sr": "ovSR", "rw": "ovRW", "goalNow": "ovGoalNow",
          "goalTgt": "ovGoalTgt", "safeSub": "ovSafeSub"}
CHROME = {"tileNW": ".heroes .card:nth-child(2) .klabel",
          "secChanged": ".sec-h", "footer": ".appfooter span"}
CANVASES = ["ovFc", "ovTrend", "ovDonut", "cashbal", "savrate", "netsav", "subcat", "liabMix"]
LANGS = ["en", "it", "es", "fr", "de"]


def digits(s):
    """Strip everything but the digits — locale formatting changes separators, not values."""
    return re.sub(r"\D", "", s or "")


def snapshot(pg):
    out = {"scalars": {}, "chrome": {}, "canvases": {}}
    for name, eid in PROBES.items():
        out["scalars"][name] = pg.eval_on_selector(f"#{eid}", "e=>(e.textContent||'').trim()")
    for name, sel in CHROME.items():
        try:
            out["chrome"][name] = pg.eval_on_selector(sel, "e=>(e.textContent||'').trim()")
        except Exception:
            out["chrome"][name] = "<<MISSING>>"
    out["bs"] = pg.eval_on_selector("#bsWrap", "e=>e.children.length")
    out["bills"] = pg.eval_on_selector("#billsDueBody", "e=>e.children.length")
    out["budget"] = pg.eval_on_selector("#ovBudget", "e=>e.children.length")
    out["insights"] = pg.eval_on_selector("#ovInsights", "e=>e.children.length")
    out["helpBadges"] = pg.evaluate("document.querySelectorAll('.help').length")
    out["lang"] = pg.evaluate("document.documentElement.getAttribute('lang')")
    for cid in CANVASES:
        out["canvases"][cid] = pg.evaluate(
            """cid=>{const el=document.getElementById(cid); if(!el)return 'no element';
                 const c=window.Chart&&Chart.getChart(el); if(!c)return 'no chart';
                 return (c.data.datasets||[]).reduce((t,d)=>t+(d.data||[]).filter(v=>v!=null&&v!=='').length,0);}""", cid)
    return out


def run(pg, wb_path, preset_lang=None, switch_to=None, analyst=True):
    logs = []
    pg.on("console", lambda m: logs.append((m.type, m.text)))
    pg.on("pageerror", lambda e: logs.append(("pageerror", str(e))))
    pg.on("dialog", lambda d: (logs.append(("dialog", d.message)), d.dismiss()))
    pg.goto(tmp.as_uri())
    if preset_lang:
        pg.evaluate("l=>localStorage.setItem('fin_lang',l)", preset_lang)
        pg.reload()
    pg.wait_for_timeout(500)
    pg.set_input_files("#wbFile", str(wb_path))
    pg.wait_for_timeout(3200)
    if analyst:
        pg.click("#anBtn"); pg.wait_for_timeout(2200); pg.click("#ovBtn"); pg.wait_for_timeout(800)
    if switch_to:
        pg.select_option("#langSel", switch_to)
        pg.wait_for_timeout(3000)
        if analyst:
            pg.click("#anBtn"); pg.wait_for_timeout(2000); pg.click("#ovBtn"); pg.wait_for_timeout(800)
    snap = snapshot(pg)
    # privacy round-trip in this language
    pg.keyboard.press("h"); pg.wait_for_timeout(600)
    snap["privacyBlurred"] = pg.evaluate("document.querySelectorAll('body.privacy .pv').length")
    # A leak = a money or percentage figure sitting in a bare text node, i.e. one the
    # privacy wrapper never wrapped in a .pv span. Text already inside .pv is the blur
    # itself, not a leak, so it must be excluded or every blurred figure reads as a miss.
    snap["privacyLeak"] = pg.evaluate(
        r"""(()=>{const bad=[];
             const w=document.createTreeWalker(document.body,NodeFilter.SHOW_TEXT);
             const FIG=/[£€$¥₹]\s?\d|\d\s?%/;
             while(w.nextNode()){ const n=w.currentNode,p=n.parentElement; if(!p)continue;
               if(p.closest('.pv'))continue;
               const tag=p.nodeName; if(tag==='SCRIPT'||tag==='STYLE'||tag==='BUTTON'||tag==='OPTION'||tag==='SELECT')continue;
               if(FIG.test(n.nodeValue)) bad.push(n.nodeValue.trim().slice(0,40)); }
             return bad.slice(0,8);})()""")
    pg.keyboard.press("h"); pg.wait_for_timeout(400)
    snap["privacyOff"] = pg.evaluate("!document.body.classList.contains('privacy')")
    snap["errors"] = [t for k, t in logs if k in ("error", "pageerror")]
    snap["warnings"] = [t for k, t in logs if k == "warning"]
    snap["dialogs"] = [t for k, t in logs if k == "dialog"]
    return snap


results, failures = {}, []
with sync_playwright() as p:
    b = p.chromium.launch()

    for lang in LANGS:
        pg = b.new_page(viewport={"width": 1400, "height": 1000})
        results[lang] = run(pg, WB, preset_lang=lang)
        pg.close()

    # live switch en -> it (no stored preference beforehand)
    pg = b.new_page(viewport={"width": 1400, "height": 1000})
    results["switch_en_to_it"] = run(pg, WB, preset_lang="en", switch_to="it")
    pg.close()

    for label, path in (("noConfig", HERE / "wb-noconfig.xlsx"), ("eurIT", HERE / "wb-eur-it.xlsx")):
        if path.exists():
            pg = b.new_page(viewport={"width": 1400, "height": 1000})
            results[label] = run(pg, path, preset_lang="en")
            pg.close()
    b.close()

# ---- assertions ----------------------------------------------------------
base = results["en"]
for lang in LANGS:
    r = results[lang]
    if r["errors"]:
        failures.append(f"{lang}: console errors {r['errors'][:2]}")
    if r["warnings"]:
        failures.append(f"{lang}: console warnings {r['warnings'][:2]}")
    if r["lang"] != lang:
        failures.append(f"{lang}: <html lang> is {r['lang']!r}")
    for k in ("safe", "nw", "nl", "sr", "goalNow"):
        if digits(r["scalars"][k]) != digits(base["scalars"][k]):
            failures.append(f"{lang}: figure {k} changed {base['scalars'][k]!r} -> {r['scalars'][k]!r}")
    for cid in CANVASES:
        if r["canvases"][cid] != base["canvases"][cid]:
            failures.append(f"{lang}: canvas {cid} {base['canvases'][cid]} -> {r['canvases'][cid]}")
    for k in ("bs", "bills", "budget", "insights", "helpBadges"):
        if r[k] != base[k]:
            failures.append(f"{lang}: container {k} {base[k]} -> {r[k]}")
    if r["privacyBlurred"] < 60:
        failures.append(f"{lang}: only {r['privacyBlurred']} blurred figures")
    if r["privacyLeak"]:
        failures.append(f"{lang}: UNBLURRED FIGURES in privacy mode: {r['privacyLeak']}")
    if not r["privacyOff"]:
        failures.append(f"{lang}: privacy mode did not toggle back off")
    if lang != "en" and r["chrome"]["tileNW"] == base["chrome"]["tileNW"]:
        failures.append(f"{lang}: 'Net worth' tile label was not translated")
    if lang != "en" and r["scalars"]["verdict"] == base["scalars"]["verdict"]:
        failures.append(f"{lang}: verdict headline was not translated")

sw = results["switch_en_to_it"]
if sw["lang"] != "it":
    failures.append(f"live switch: <html lang> is {sw['lang']!r}")
if sw["chrome"]["tileNW"] != results["it"]["chrome"]["tileNW"]:
    failures.append("live switch: did not match a fresh Italian load")
for k in ("nw", "nl"):
    if digits(sw["scalars"][k]) != digits(base["scalars"][k]):
        failures.append(f"live switch: figure {k} changed")
if sw["errors"]:
    failures.append(f"live switch: console errors {sw['errors'][:2]}")

for label in ("noConfig", "eurIT"):
    if label in results:
        r = results[label]
        if r["errors"]:
            failures.append(f"{label}: console errors {r['errors'][:2]}")
        if not r["scalars"]["nw"] or r["scalars"]["nw"] == "—":
            failures.append(f"{label}: net worth did not render")
        if r["privacyLeak"]:
            failures.append(f"{label}: UNBLURRED FIGURES in privacy mode: {r['privacyLeak']}")

print(json.dumps({k: {"lang": v["lang"], "scalars": v["scalars"], "chrome": v["chrome"],
                      "blurred": v["privacyBlurred"], "leak": v["privacyLeak"],
                      "help": v["helpBadges"], "errors": v["errors"], "warnings": v["warnings"]}
                  for k, v in results.items()}, indent=1, ensure_ascii=False))
print("\n" + ("=" * 70))
if failures:
    print(f"FAIL — {len(failures)} problem(s):")
    for f in failures:
        print("  ·", f)
    sys.exit(1)
print("PASS — all languages render identical figures with translated chrome.")
