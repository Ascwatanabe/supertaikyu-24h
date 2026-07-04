"""supertaikyu.live/timings/ のAPIエンドポイントを調査するスクリプト。"""
import re
import requests

s = requests.Session()
s.headers["User-Agent"] = "Mozilla/5.0"

print("=== /timings/ ページを取得中 ===")
resp = s.get("https://www.supertaikyu.live/timings/", timeout=15)
text = resp.text
print(f"ステータス: {resp.status_code}, サイズ: {len(text)} bytes")

# script src を抽出
scripts = re.findall(r'src=["\']([^"\']+)["\']', text)
print("\n=== ロードされるScript ===")
for sc in scripts[:15]:
    print(" ", sc)

# HTML内のJSON/APIパスを探す
hits = re.findall(r'["\'`](/[^"\'`\s]{3,100})["\'`]', text)
json_hits = [h for h in hits if any(k in h.lower() for k in ("json", "api", "data", "master", "live"))]
print("\n=== HTML内のJSON/APIパス候補 ===")
for h in sorted(set(json_hits))[:30]:
    print(" ", h)

# JSファイルをスキャン
print("\n=== JSファイル内のURLパターン ===")
for src in scripts[:8]:
    if src.startswith("http"):
        js_url = src
    elif src.startswith("/"):
        js_url = "https://www.supertaikyu.live" + src
    else:
        continue
    try:
        js = s.get(js_url, timeout=10).text
        found = re.findall(r'["\'`](/[^"\'`\s]{3,80})["\'`]', js)
        api_found = [f for f in found if any(k in f.lower() for k in ("json", "api", "master", "live", "timing"))]
        if api_found:
            print(f"\n  [{src}]")
            for f in sorted(set(api_found))[:15]:
                print(f"    {f}")
    except Exception as e:
        print(f"  {src}: {e}")

# probeテスト
print("\n=== 候補URLのmaster.json確認 ===")
candidates = [
    "https://www.supertaikyu.live/json",
    "https://www.supertaikyu.live/timings/json",
    "https://www.supertaikyu.live/timings/data",
    "https://www.supertaikyu.live/timings/api",
    "https://www.supertaikyu.live/api",
    "https://www.supertaikyu.live/data",
]
for url in candidates:
    try:
        r = s.get(f"{url}/master.json", timeout=8)
        print(f"  {url}/master.json → {r.status_code}")
    except Exception as e:
        print(f"  {url}/master.json → ERROR: {e}")
