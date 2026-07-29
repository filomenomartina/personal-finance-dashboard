#!/usr/bin/env python3
"""
Headless smoke/regression test for the dashboard.

Loads the HTML in Chromium with the CDN <script> tags rewritten to local vendor
copies (the sandbox has no CDN access), feeds it the starter workbook through
the real file input, and reports what actually rendered.

Usage: python3 test_dashboard.py <dashboard.html> <workbook.xlsx> [label]
"""
import sys, json, re, pathlib, tempfile
from playwright.sync_api import sync_playwright

HTML = pathlib.Path(sys.argv[1]).resolve()
WB = pathlib.Path(sys.argv[2]).resolve()
LABEL = sys.argv[3] if len(sys.argv) > 3 else HTML.stem
HERE = pathlib.Path(__file__).resolve().parent

# Rewrite CDN tags -> local vendor files so the page works offline.
src = HTML.read_text()
src = src.replace(
    "https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.0/chart.umd.min.js",
    "vendor_chart.js")
src = src.replace(
    "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js",
    "vendor_xlsx.js")
tmp = HERE / f"_test_{LABEL}.html"
tmp.write_text(src)

PROBES = {
    "verdict title": "ovVt",
    "safe to spend": "ovSafe",
    "net worth": "ovNW",
    "net liquid": "ovNL",
    "savings rate": "ovSR",
    "cash runway": "ovRW",
    "goal now": "ovGoalNow",
    "goal target": "ovGoalTgt",
    "data as of": "dataAsOf",
}
CONTAINERS = {
    "insights": "ovInsights",
    "budget rows": "ovBudget",
    "balance sheet": "bsWrap",
    "bills due": "billsDueBody",
    "small multiples": "smgrid",
    "wish list": "wlInner",
}
CANVASES = ["ovFc", "ovTrend", "ovDonut", "cashbal", "savrate", "netsav",
            "subcat", "liabMix"]

with sync_playwright() as p:
    b = p.chromium.launch()
    pg = b.new_page(viewport={"width": 1400, "height": 1000})
    logs = []
    pg.on("console", lambda m: logs.append((m.type, m.text)))
    pg.on("pageerror", lambda e: logs.append(("pageerror", str(e))))
    dialogs = []
    pg.on("dialog", lambda d: (dialogs.append(d.message), d.dismiss()))

    pg.goto(tmp.as_uri())
    pg.wait_for_timeout(600)
    pg.set_input_files("#wbFile", str(WB))
    pg.wait_for_timeout(3500)

    out = {"label": LABEL, "scalars": {}, "containers": {}, "canvases": {},
           "dialogs": dialogs}
    for name, eid in PROBES.items():
        try:
            out["scalars"][name] = pg.eval_on_selector(
                f"#{eid}", "e => (e.textContent||'').trim().slice(0,60)")
        except Exception:
            out["scalars"][name] = "<<MISSING ELEMENT>>"
    for name, eid in CONTAINERS.items():
        try:
            out["containers"][name] = pg.eval_on_selector(
                f"#{eid}", "e => e.children.length")
        except Exception:
            out["containers"][name] = "<<MISSING ELEMENT>>"

    # switch to Analyst so its charts actually build
    pg.click("#anBtn")
    pg.wait_for_timeout(2500)
    for name, eid in CONTAINERS.items():
        try:
            n = pg.eval_on_selector(f"#{eid}", "e => e.children.length")
            if isinstance(out["containers"].get(name), int):
                out["containers"][name] = max(out["containers"][name], n)
            else:
                out["containers"][name] = n
        except Exception:
            pass
    for cid in CANVASES:
        out["canvases"][cid] = pg.evaluate(
            """cid => { const el=document.getElementById(cid);
                 if(!el) return 'no element';
                 const c=window.Chart && Chart.getChart(el);
                 if(!c) return 'no chart';
                 const n=(c.data.datasets||[]).reduce((t,d)=>t+(d.data||[]).filter(v=>v!=null&&v!=='').length,0);
                 return n; }""", cid)

    pg.screenshot(path=str(HERE / f"shot_{LABEL}_analyst.png"), full_page=True)
    pg.click("#ovBtn")
    pg.wait_for_timeout(1200)
    pg.screenshot(path=str(HERE / f"shot_{LABEL}_overview.png"), full_page=True)

    # privacy mode round-trip
    pg.keyboard.press("h")
    pg.wait_for_timeout(700)
    out["privacy_on_blurred_spans"] = pg.evaluate(
        "document.querySelectorAll('body.privacy .pv').length")
    pg.keyboard.press("h")
    pg.wait_for_timeout(400)
    out["privacy_off"] = pg.evaluate(
        "!document.body.classList.contains('privacy')")

    errs = [t for k, t in logs if k in ("error", "pageerror")]
    warns = [t for k, t in logs if k == "warning"]
    out["errors"] = errs
    out["warnings"] = warns
    b.close()

print(json.dumps(out, indent=2))
