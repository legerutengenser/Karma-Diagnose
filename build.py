#!/usr/bin/env python3
"""Bygger index.html: limer data/kar-icd10-synonymer.json inn i <script id="karma-data">.

Bruk:  python3 build.py
Kilden til sannhet er JSON-filen. Rediger aldri dataene direkte i index.html.
Scriptet er idempotent, og oppdaterer også CACHE_VERSION i sw.js hvis dataene
eller index.html har endret seg (slik at installerte apper henter ny versjon).
"""
import json, re, sys, hashlib, pathlib

root = pathlib.Path(__file__).parent
src = root / "data" / "kar-icd10-synonymer.json"
html_path = root / "index.html"
sw_path = root / "sw.js"

data = json.loads(src.read_text(encoding="utf8"))  # validerer JSON
payload = json.dumps(data, ensure_ascii=False, indent=2)
if "</script" in payload.lower():
    sys.exit("FEIL: JSON inneholder '</script' som vil ødelegge HTML-en.")

html = html_path.read_text(encoding="utf8")
pat = re.compile(r'(<script id="karma-data" type="application/json">)(.*?)(</script>)', re.S)
if not pat.search(html):
    sys.exit('FEIL: fant ikke <script id="karma-data"> i index.html')
new_html = pat.sub(lambda m: m.group(1) + "\n" + payload + "\n" + m.group(3), html, count=1)

changed = new_html != html
if changed:
    html_path.write_text(new_html, encoding="utf8")

# Bump service worker cache version when index.html changed (data OR layout/code).
# We compare a hash of index.html with the hash stored in sw.js.
sw = sw_path.read_text(encoding="utf8")
cur_hash = hashlib.sha1(html_path.read_bytes()).hexdigest()[:10]
hm = re.search(r"// index-hash: (\w+)", sw)
vm = re.search(r'const CACHE_VERSION = "v(\d+)";', sw)
bumped = None
if vm and (not hm or hm.group(1) != cur_hash):
    bumped = int(vm.group(1)) + 1
    sw = sw.replace(vm.group(0), f'const CACHE_VERSION = "v{bumped}";')
    if hm:
        sw = sw.replace(hm.group(0), f"// index-hash: {cur_hash}")
    else:
        sw = sw.replace("// Karma Diagnose - service worker", f"// Karma Diagnose - service worker\n// index-hash: {cur_hash}", 1)
    sw_path.write_text(sw, encoding="utf8")

n_main, n_adj = len(data["codes"]), len(data["tilstotende_koder"])
print(f"{'Oppdatert' if changed else 'Uendret'}: index.html  ({n_main} hovedkoder + {n_adj} tilstøtende)")
if bumped:
    print(f"sw.js CACHE_VERSION -> v{bumped}")
