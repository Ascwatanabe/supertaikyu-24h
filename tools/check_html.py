"""index.html のJS初期化部分の問題を確認する。"""
import re, sys
import pathlib

sys.stdout.reconfigure(encoding="utf-8")

html = pathlib.Path("C:/Users/k-wat/src/supertaikyu-24h/data/index.html")
if not html.exists():
    print("index.html NOT FOUND")
    exit()

text = html.read_text("utf-8", errors="replace")
print(f"index.html size: {len(text)//1024} KB")

ids_in_html = set(re.findall(r'id="([^"]+)"', text))
ids_needed = [
    "dataUrlInput", "dataUrlApplyBtn", "dataUrlDiscoverBtn",
    "dataUrlResetBtn", "dataUrlStatus", "clearDataBtn",
    "youtubeUrlInput", "youtubeApplyBtn", "youtubeResetBtn",
    "classFilter", "rows", "tableWrap", "youtubeFrame",
    "lapColumn", "lapRows",
]
print("\n=== DOM element check ===")
for id_ in ids_needed:
    status = "OK" if id_ in ids_in_html else "MISSING"
    print(f"  [{status}] #{id_}")

print("\n=== JS variables ===")
for pattern, label in [
    (r'const DATA_URL = "([^"]+)"', "DATA_URL"),
    (r'const DEFAULT_BASE_URL = "([^"]+)"', "DEFAULT_BASE_URL"),
    (r'const DEFAULT_CLASS = "([^"]+)"', "DEFAULT_CLASS"),
    (r'const POLL_INTERVAL = ([^;]+);', "POLL_INTERVAL"),
]:
    m = re.search(pattern, text)
    print(f"  {label}: {m.group(1) if m else 'NOT FOUND'}")

print("\n=== Script brace balance ===")
script_blocks = re.findall(r'<script>(.*?)</script>', text, re.DOTALL)
for i, block in enumerate(script_blocks):
    opens = block.count("{")
    closes = block.count("}")
    diff = opens - closes
    status = "OK" if abs(diff) < 5 else f"MISMATCH diff={diff}"
    print(f"  script#{i+1}: {{ {opens} / }} {closes}  [{status}]")
