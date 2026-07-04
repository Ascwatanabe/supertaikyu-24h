import json, pathlib

p = pathlib.Path("C:/Users/k-wat/src/supertaikyu-24h/data/view_data.json")
if p.exists():
    size = p.stat().st_size / 1024
    d = json.loads(p.read_text("utf-8"))
    print(f"size: {size:.0f} KB")
    print(f"cars: {len(d.get('cars', []))}")
    print(f"lap_history: {len(d.get('lap_history', []))}")
    print(f"status: {d.get('status', '')}")
    print(f"saved_at: {d.get('saved_at', '')}")
    print(f"classes: {d.get('classes', [])}")
    if d.get("cars"):
        print(f"first car: {d['cars'][0]}")
else:
    print("view_data.json NOT FOUND")

# index.html の DATA_URL 確認
ih = pathlib.Path("C:/Users/k-wat/src/supertaikyu-24h/data/index.html")
if ih.exists():
    text = ih.read_text("utf-8", errors="replace")
    import re
    m = re.search(r'const DATA_URL = "([^"]+)"', text)
    print(f"\nDATA_URL in index.html: {m.group(1) if m else 'NOT FOUND'}")
    m2 = re.search(r'const DEFAULT_BASE_URL = "([^"]+)"', text)
    print(f"DEFAULT_BASE_URL in index.html: {m2.group(1) if m2 else 'NOT FOUND'}")
    m3 = re.search(r'const POLL_INTERVAL = ([^;]+);', text)
    print(f"POLL_INTERVAL: {m3.group(1) if m3 else 'NOT FOUND'}")
else:
    print("index.html NOT FOUND")
