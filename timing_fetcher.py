#!/usr/bin/env python3
"""Super Taikyu live timing fetcher."""

from __future__ import annotations

import argparse
import csv
import html
import json
import sys
import time
import webbrowser
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

import requests

APP_DIR = Path(__file__).resolve().parent
BASE_URL = "https://www.supertaikyu.live/json"
DEFAULT_CLASS = "ST-5F"
DEFAULT_INTERVAL = 5
DEFAULT_OUTPUT_DIR = APP_DIR / "data"
REQUEST_TIMEOUT = 10
MAX_RETRIES = 3

ALL_CLASSES = [
    "ST-X",
    "ST-Z",
    "ST-TCR",
    "ST-USA",
    "ST-Q",
    "ST-1",
    "ST-2",
    "ST-3",
    "ST-4",
    "ST-5F",
    "ST-5R",
]


@dataclass
class TimingRow:
    pos: str
    pic: str
    car_no: str
    car_class: str
    laps: str
    current_driver: str
    all_drivers: str
    team_name: str
    car_name: str
    best_lap: str
    last_lap: str
    best_lap_ms: int
    last_lap_ms: int
    has_live_data: bool


@dataclass
class DriverRow:
    driver_name: str
    driver_slot: str
    car_no: str
    car_class: str
    team_name: str
    car_name: str
    best_lap: str
    last_lap: str
    best_lap_ms: int
    last_lap_ms: int
    is_current: bool
    pos: str
    pic: str
    laps: str


def ms_to_laptime(ms: int | float | str) -> str:
    """Convert milliseconds to the same format used on supertaikyu.live."""
    try:
        time_ms = int(ms)
    except (TypeError, ValueError):
        return "-"

    if time_ms <= 0:
        return "-"

    hh = time_ms // 3_600_000
    mm = (time_ms % 3_600_000) // 60_000
    ss = (time_ms % 60_000) // 1_000
    milli = time_ms % 1_000

    if hh > 0:
        return f"{hh}:{mm:02d}'{ss:02d}.{milli:03d}"
    if mm > 0:
        return f"{mm}'{ss:02d}.{milli:03d}"
    return f"{ss}.{milli:03d}"


def fetch_json(session: requests.Session, name: str) -> dict[str, Any]:
    url = f"{BASE_URL}/{name}"
    last_error: Exception | None = None

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = session.get(url, timeout=REQUEST_TIMEOUT)
            response.raise_for_status()
            return response.json()
        except (requests.RequestException, ValueError) as exc:
            last_error = exc
            if attempt < MAX_RETRIES:
                time.sleep(1)

    raise RuntimeError(f"Failed to fetch {url}: {last_error}") from last_error


def get_driver_name(
    entry: dict[str, Any] | None,
    driver_id: str,
    use_english: bool = False,
) -> str:
    if not entry:
        return ""

    drivers = entry.get("Drivers", {})
    driver = drivers.get(driver_id, {})
    key = "eNameL" if use_english else "jNameL"
    return str(driver.get(key, "")).strip()


def get_all_driver_names(entry: dict[str, Any] | None, use_english: bool = False) -> str:
    if not entry:
        return ""

    drivers = entry.get("Drivers", {})
    key = "eNameL" if use_english else "jNameL"
    names: list[str] = []

    for slot in sorted(drivers.keys()):
        name = str(drivers[slot].get(key, "")).strip()
        if name:
            names.append(name)

    return " / ".join(names)


def has_active_timing(live_data: list[dict[str, Any]]) -> bool:
    for row in live_data:
        if int(row.get("BestLapTime", 0) or 0) > 0:
            return True
        if int(row.get("LapTime", 0) or 0) > 0:
            return True
        if str(row.get("LAPS", "")).strip().isdigit() and int(row.get("LAPS", 0)) > 0:
            return True
    return False


def build_rows(
    master: dict[str, Any],
    live: dict[str, Any],
    class_filter: str | None,
    use_english: bool = False,
) -> list[TimingRow]:
    entry_info: dict[str, Any] = master.get("EntryInfo", {})
    live_data: list[dict[str, Any]] = live.get("LiveData", [])
    live_by_car = {str(row.get("CarNo", "")): row for row in live_data}

    rows: list[TimingRow] = []

    for car_no, entry in entry_info.items():
        car_class = str(entry.get("ClassStr", "")).strip()
        if class_filter and car_class != class_filter:
            continue

        live_row = live_by_car.get(str(car_no), {})
        current_driver_id = str(live_row.get("Driver", "")).strip()
        current_driver = get_driver_name(entry, current_driver_id, use_english)

        best_lap_ms = int(live_row.get("BestLapTime", 0) or 0)
        last_lap_ms = int(live_row.get("LapTime", 0) or 0)
        has_live = bool(
            live_row
            and (
                best_lap_ms > 0
                or last_lap_ms > 0
                or bool(str(live_row.get("LAPS", "")).strip())
            )
        )

        rows.append(
            TimingRow(
                pos=str(live_row.get("POS", "")).strip() or "-",
                pic=str(live_row.get("PIC", "")).strip() or "-",
                car_no=str(car_no),
                car_class=car_class or "-",
                laps=str(live_row.get("LAPS", "")).strip() or "-",
                current_driver=current_driver or "-",
                all_drivers=get_all_driver_names(entry, use_english) or "-",
                team_name=str(entry.get("TeamNameL", "")).strip() or "-",
                car_name=str(entry.get("CarNameL", "")).strip() or "-",
                best_lap=ms_to_laptime(best_lap_ms),
                last_lap=ms_to_laptime(last_lap_ms),
                best_lap_ms=best_lap_ms,
                last_lap_ms=last_lap_ms,
                has_live_data=has_live,
            )
        )

    def sort_key(row: TimingRow) -> tuple:
        pic = int(row.pic) if row.pic.isdigit() else 9999
        pos = int(row.pos) if row.pos.isdigit() else 9999
        return (pic, pos, int(row.car_no) if row.car_no.isdigit() else 9999)

    rows.sort(key=sort_key)
    return rows


def build_driver_rows(
    master: dict[str, Any],
    live: dict[str, Any],
    class_filter: str | None,
    use_english: bool = False,
) -> list[DriverRow]:
    entry_info: dict[str, Any] = master.get("EntryInfo", {})
    live_data: list[dict[str, Any]] = live.get("LiveData", [])
    live_by_car = {str(row.get("CarNo", "")): row for row in live_data}
    driver_rows: list[DriverRow] = []

    for car_no, entry in entry_info.items():
        car_class = str(entry.get("ClassStr", "")).strip()
        if class_filter and car_class != class_filter:
            continue

        live_row = live_by_car.get(str(car_no), {})
        current_driver_id = str(live_row.get("Driver", "")).strip()
        best_lap_ms = int(live_row.get("BestLapTime", 0) or 0)
        last_lap_ms = int(live_row.get("LapTime", 0) or 0)

        for slot in sorted(entry.get("Drivers", {}).keys()):
            driver_name = get_driver_name(entry, slot, use_english)
            if not driver_name:
                continue

            driver_rows.append(
                DriverRow(
                    driver_name=driver_name,
                    driver_slot=slot,
                    car_no=str(car_no),
                    car_class=car_class or "-",
                    team_name=str(entry.get("TeamNameL", "")).strip() or "-",
                    car_name=str(entry.get("CarNameL", "")).strip() or "-",
                    best_lap=ms_to_laptime(best_lap_ms),
                    last_lap=ms_to_laptime(last_lap_ms),
                    best_lap_ms=best_lap_ms,
                    last_lap_ms=last_lap_ms,
                    is_current=slot == current_driver_id and bool(current_driver_id),
                    pos=str(live_row.get("POS", "")).strip() or "-",
                    pic=str(live_row.get("PIC", "")).strip() or "-",
                    laps=str(live_row.get("LAPS", "")).strip() or "-",
                )
            )

    driver_rows.sort(key=lambda row: (row.driver_name, row.car_no))
    return driver_rows


def search_rows(rows: list[TimingRow], query: str) -> list[TimingRow]:
    query = query.strip().casefold()
    if not query:
        return rows

    matched: list[TimingRow] = []
    for row in rows:
        haystack = " ".join(
            [row.car_no, row.current_driver, row.all_drivers, row.team_name, row.car_name]
        ).casefold()
        if query in haystack:
            matched.append(row)
    return matched


def format_update_time(raw: str) -> str:
    raw = str(raw).strip()
    if len(raw) < 15:
        return raw or "-"

    try:
        dt = datetime.strptime(raw[:15], "%Y%m%d_%H%M%S")
        return dt.strftime("%Y-%m-%d %H:%M:%S")
    except ValueError:
        return raw


def race_status_text(live: dict[str, Any]) -> str:
    live_data = live.get("LiveData", [])
    if has_active_timing(live_data):
        return "レースデータ受信中"
    if int(live.get("IsFinal", 0) or 0) == 1:
        return "待機中（エントリー情報のみ / タイム未配信）"
    return "待機中（タイム未配信）"


def clear_screen() -> None:
    if sys.platform == "win32":
        import os

        os.system("cls")
    else:
        print("\033[2J\033[H", end="")


def print_table(rows: list[TimingRow], header_lines: list[str]) -> None:
    for line in header_lines:
        print(line)

    if not rows:
        print("\n該当データがありません。")
        return

    columns = [
        ("POS", 4),
        ("PIC", 4),
        ("No.", 5),
        ("Class", 7),
        ("Driver", 18),
        ("BestLap", 12),
        ("LastLap", 12),
        ("Laps", 5),
        ("Team", 24),
    ]

    header = " ".join(name.ljust(width) for name, width in columns)
    print()
    print(header)
    print("-" * len(header))

    for row in rows:
        values = [
            row.pos,
            row.pic,
            row.car_no,
            row.car_class,
            row.current_driver[:18],
            row.best_lap,
            row.last_lap,
            row.laps,
            row.team_name[:24],
        ]
        print(" ".join(value.ljust(width) for value, (_, width) in zip(values, columns)))

    print()
    print(f"表示件数: {len(rows)}")


def print_search_results(rows: list[TimingRow], query: str) -> None:
    print(f"\n検索: \"{query}\"")
    if not rows:
        print("該当するドライバーが見つかりませんでした。")
        return

    for row in rows:
        print("-" * 60)
        print(f"車番      : {row.car_no}")
        print(f"クラス    : {row.car_class}")
        print(f"現在走行  : {row.current_driver}")
        print(f"全ドライバー: {row.all_drivers}")
        print(f"BestLap   : {row.best_lap}")
        print(f"LastLap   : {row.last_lap}")
        print(f"チーム    : {row.team_name}")
        print(f"マシン    : {row.car_name}")


def load_timing_data(
    session: requests.Session,
    class_filter: str | None,
    use_english: bool,
) -> tuple[list[TimingRow], dict[str, Any], dict[str, Any]]:
    master = fetch_json(session, "master.json")
    live = fetch_json(session, "livemoni.json")
    rows = build_rows(master, live, class_filter, use_english)
    return rows, master, live


def ensure_output_dir(output_dir: Path) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    return output_dir


def build_snapshot(
    rows: list[TimingRow],
    master: dict[str, Any],
    live: dict[str, Any],
    class_filter: str | None,
) -> dict[str, Any]:
    class_label = class_filter or "ALL"
    return {
        "saved_at": datetime.now().isoformat(timespec="seconds"),
        "race_name": master.get("RaceNameL", ""),
        "race_year": master.get("RaceYear", ""),
        "round_no": master.get("RoundNo", ""),
        "update_time": live.get("UpdateTime", ""),
        "update_time_display": format_update_time(live.get("UpdateTime", "")),
        "status": race_status_text(live),
        "class_filter": class_label,
        "row_count": len(rows),
        "rows": [asdict(row) for row in rows],
    }


def save_latest_files(
    output_dir: Path,
    rows: list[TimingRow],
    master: dict[str, Any],
    live: dict[str, Any],
    class_filter: str | None,
) -> dict[str, Path]:
    ensure_output_dir(output_dir)
    class_label = class_filter or "ALL"
    snapshot = build_snapshot(rows, master, live, class_filter)

    json_path = output_dir / f"latest_{class_label}.json"
    csv_path = output_dir / f"latest_{class_label}.csv"

    json_path.write_text(
        json.dumps(snapshot, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    fieldnames = [
        "saved_at",
        "update_time_display",
        "status",
        "pos",
        "pic",
        "car_no",
        "car_class",
        "current_driver",
        "all_drivers",
        "best_lap",
        "last_lap",
        "best_lap_ms",
        "last_lap_ms",
        "laps",
        "team_name",
        "car_name",
        "has_live_data",
    ]

    with csv_path.open("w", encoding="utf-8-sig", newline="") as csv_file:
        writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow(
                {
                    "saved_at": snapshot["saved_at"],
                    "update_time_display": snapshot["update_time_display"],
                    "status": snapshot["status"],
                    **asdict(row),
                }
            )

    return {"json": json_path, "csv": csv_path}


def append_history(
    output_dir: Path,
    rows: list[TimingRow],
    master: dict[str, Any],
    live: dict[str, Any],
    class_filter: str | None,
) -> Path:
    ensure_output_dir(output_dir)
    class_label = class_filter or "ALL"
    history_path = output_dir / f"history_{class_label}.csv"
    snapshot = build_snapshot(rows, master, live, class_filter)

    fieldnames = [
        "saved_at",
        "update_time_display",
        "status",
        "pos",
        "pic",
        "car_no",
        "car_class",
        "current_driver",
        "all_drivers",
        "best_lap",
        "last_lap",
        "best_lap_ms",
        "last_lap_ms",
        "laps",
        "team_name",
        "car_name",
        "has_live_data",
    ]

    write_header = not history_path.exists()
    with history_path.open("a", encoding="utf-8-sig", newline="") as csv_file:
        writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
        if write_header:
            writer.writeheader()
        for row in rows:
            writer.writerow(
                {
                    "saved_at": snapshot["saved_at"],
                    "update_time_display": snapshot["update_time_display"],
                    "status": snapshot["status"],
                    **asdict(row),
                }
            )

    return history_path


def save_html_file(
    output_dir: Path,
    rows: list[TimingRow],
    master: dict[str, Any],
    live: dict[str, Any],
    class_filter: str | None,
    interval: int,
) -> Path:
    ensure_output_dir(output_dir)
    class_label = class_filter or "ALL"
    snapshot = build_snapshot(rows, master, live, class_filter)
    html_path = output_dir / "index.html"

    table_rows: list[str] = []
    for row in rows:
        best_class = "lap-best" if row.best_lap_ms > 0 else ""
        live_class = "row-live" if row.has_live_data else ""
        table_rows.append(
            f"""<tr class="{live_class}" data-search="{html.escape(
                ' '.join(
                    [
                        row.car_no,
                        row.current_driver,
                        row.all_drivers,
                        row.team_name,
                        row.car_name,
                    ]
                ),
                quote=True,
            )}">
  <td class="num">{html.escape(row.pos)}</td>
  <td class="num">{html.escape(row.pic)}</td>
  <td class="num car-no">{html.escape(row.car_no)}</td>
  <td><span class="class-badge">{html.escape(row.car_class)}</span></td>
  <td class="driver">{html.escape(row.current_driver)}</td>
  <td class="drivers">{html.escape(row.all_drivers)}</td>
  <td class="lap {best_class}">{html.escape(row.best_lap)}</td>
  <td class="lap">{html.escape(row.last_lap)}</td>
  <td class="num">{html.escape(row.laps)}</td>
  <td class="team">{html.escape(row.team_name)}</td>
  <td class="car">{html.escape(row.car_name)}</td>
</tr>"""
        )

    status = html.escape(snapshot["status"])
    status_class = "status-live" if "受信中" in snapshot["status"] else "status-wait"

    page = f"""<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta http-equiv="refresh" content="{interval}">
  <title>スーパー耐久 24H タイミング | {html.escape(class_label)}</title>
  <style>
    :root {{
      --bg: #11151c;
      --panel: #1a2030;
      --border: #2d3648;
      --text: #e8edf7;
      --muted: #9aa7bd;
      --accent: #e10600;
      --accent-soft: #ff4d4d;
      --class: #1f8f4e;
      --best: #7dd3fc;
      --live: rgba(31, 143, 78, 0.12);
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: "Segoe UI", "Hiragino Sans", "Yu Gothic UI", sans-serif;
      background: var(--bg);
      color: var(--text);
    }}
    .wrap {{ max-width: 1600px; margin: 0 auto; padding: 20px; }}
    .header {{
      background: linear-gradient(135deg, #1a2030 0%, #10141c 100%);
      border: 1px solid var(--border);
      border-radius: 12px;
      padding: 20px 24px;
      margin-bottom: 16px;
    }}
    h1 {{ margin: 0 0 8px; font-size: 1.5rem; }}
    .meta {{ color: var(--muted); font-size: 0.95rem; line-height: 1.7; }}
    .status {{
      display: inline-block;
      padding: 4px 10px;
      border-radius: 999px;
      font-size: 0.85rem;
      font-weight: 700;
      margin-top: 8px;
    }}
    .status-live {{ background: rgba(31, 143, 78, 0.2); color: #7dffb2; }}
    .status-wait {{ background: rgba(255, 180, 0, 0.15); color: #ffd166; }}
    .toolbar {{
      display: flex;
      gap: 12px;
      align-items: center;
      margin-bottom: 16px;
      flex-wrap: wrap;
    }}
    .toolbar input {{
      flex: 1;
      min-width: 240px;
      padding: 10px 14px;
      border-radius: 8px;
      border: 1px solid var(--border);
      background: var(--panel);
      color: var(--text);
      font-size: 1rem;
    }}
    .toolbar .info {{ color: var(--muted); font-size: 0.9rem; }}
    .nav-link {{
      color: #9fd0ff;
      text-decoration: none;
      padding: 10px 14px;
      border: 1px solid var(--border);
      border-radius: 8px;
      white-space: nowrap;
    }}
    .nav-link:hover {{ background: rgba(255, 255, 255, 0.05); }}
    .table-wrap {{
      overflow: auto;
      border: 1px solid var(--border);
      border-radius: 12px;
      background: var(--panel);
    }}
    table {{
      width: 100%;
      border-collapse: collapse;
      min-width: 1100px;
      font-size: 0.92rem;
    }}
    th, td {{
      padding: 10px 12px;
      border-bottom: 1px solid var(--border);
      text-align: left;
      vertical-align: top;
    }}
    tbody td {{
      font-size: calc(1em - 1pt);
    }}
    th {{
      position: sticky;
      top: 0;
      background: #232b3d;
      color: #d7e0f1;
      font-size: 0.8rem;
      letter-spacing: 0.04em;
      text-transform: uppercase;
      z-index: 1;
    }}
    tr:hover {{ background: rgba(255, 255, 255, 0.03); }}
    tr.row-live {{ background: var(--live); }}
    .num {{ text-align: center; white-space: nowrap; }}
    .car-no {{ font-weight: 700; color: #fff; }}
    .class-badge {{
      display: inline-block;
      background: rgba(31, 143, 78, 0.2);
      color: #8dffb8;
      padding: 2px 8px;
      border-radius: 6px;
      font-weight: 700;
      font-size: 0.82rem;
    }}
    .driver {{ font-weight: 700; min-width: 120px; }}
    .drivers, .team, .car {{ color: var(--muted); min-width: 160px; }}
    .lap {{
      font-family: Consolas, "Courier New", monospace;
      white-space: nowrap;
      text-align: right;
      min-width: 90px;
    }}
    .lap-best {{ color: var(--best); font-weight: 700; }}
    .hidden {{ display: none; }}
    .footer {{
      margin-top: 14px;
      color: var(--muted);
      font-size: 0.85rem;
    }}
  </style>
</head>
<body>
  <div class="wrap">
    <div class="header">
      <h1>スーパー耐久 24時間 ライブタイミング</h1>
      <div class="meta">
        レース: {html.escape(str(snapshot["race_name"]))} ({html.escape(str(snapshot["race_year"]))})<br>
        クラス: {html.escape(class_label)}<br>
        更新時刻: {html.escape(str(snapshot["update_time_display"]))}<br>
        保存時刻: {html.escape(str(snapshot["saved_at"]))}
      </div>
      <span class="status {status_class}">{status}</span>
    </div>

    <div class="toolbar">
      <input id="search" type="search" placeholder="車番・ドライバー・チーム名で検索...">
      <a class="nav-link" href="drivers.html">ドライバー一覧</a>
      <div class="info">表示: <span id="count">{len(rows)}</span> 件 / {interval}秒ごとに自動更新</div>
    </div>

    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>POS</th>
            <th>PIC</th>
            <th>No.</th>
            <th>Class</th>
            <th>Driver</th>
            <th>全ドライバー</th>
            <th>BestLap</th>
            <th>LastLap</th>
            <th>Laps</th>
            <th>Team</th>
            <th>Car</th>
          </tr>
        </thead>
        <tbody id="rows">
          {"".join(table_rows)}
        </tbody>
      </table>
    </div>

    <div class="footer">
      データ取得元: supertaikyu.live / このページはローカル保存された HTML です
    </div>
  </div>

  <script>
    const search = document.getElementById("search");
    const rows = Array.from(document.querySelectorAll("#rows tr"));
    const count = document.getElementById("count");

    function filterRows() {{
      const query = search.value.trim().toLowerCase();
      let visible = 0;
      for (const row of rows) {{
        const haystack = (row.dataset.search || "").toLowerCase();
        const show = !query || haystack.includes(query);
        row.classList.toggle("hidden", !show);
        if (show) visible++;
      }}
      count.textContent = String(visible);
    }}

    search.addEventListener("input", filterRows);
  </script>
</body>
</html>
"""

    html_path.write_text(page, encoding="utf-8")
    return html_path


def save_drivers_html(
    output_dir: Path,
    driver_rows: list[DriverRow],
    master: dict[str, Any],
    live: dict[str, Any],
    class_filter: str | None,
    interval: int,
) -> Path:
    ensure_output_dir(output_dir)
    class_label = class_filter or "ALL"
    status = race_status_text(live)
    status_class = "status-live" if "受信中" in status else "status-wait"
    html_path = output_dir / "drivers.html"

    driver_names = sorted({row.driver_name for row in driver_rows})
    option_rows = ['<option value="">すべてのドライバー</option>']
    for name in driver_names:
        option_rows.append(
            f'<option value="{html.escape(name, quote=True)}">{html.escape(name)}</option>'
        )

    table_rows: list[str] = []
    for row in driver_rows:
        best_class = "lap-best" if row.best_lap_ms > 0 else ""
        current_badge = '<span class="current-badge">走行中</span>' if row.is_current else "-"
        row_class = "row-live" if row.is_current else ""
        table_rows.append(
            f"""<tr class="{row_class}" data-driver="{html.escape(row.driver_name, quote=True)}" data-search="{html.escape(
                ' '.join([row.driver_name, row.car_no, row.team_name, row.car_name]),
                quote=True,
            )}">
  <td class="driver">{html.escape(row.driver_name)}</td>
  <td class="num">{html.escape(row.driver_slot)}</td>
  <td class="current">{current_badge}</td>
  <td class="num car-no">{html.escape(row.car_no)}</td>
  <td><span class="class-badge">{html.escape(row.car_class)}</span></td>
  <td class="lap {best_class}">{html.escape(row.best_lap)}</td>
  <td class="lap">{html.escape(row.last_lap)}</td>
  <td class="num">{html.escape(row.laps)}</td>
  <td class="num">{html.escape(row.pos)}</td>
  <td class="num">{html.escape(row.pic)}</td>
  <td class="team">{html.escape(row.team_name)}</td>
  <td class="car">{html.escape(row.car_name)}</td>
</tr>"""
        )

    page = f"""<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta http-equiv="refresh" content="{interval}">
  <title>ドライバー一覧 | スーパー耐久 24H | {html.escape(class_label)}</title>
  <style>
    :root {{
      --bg: #11151c;
      --panel: #1a2030;
      --border: #2d3648;
      --text: #e8edf7;
      --muted: #9aa7bd;
      --best: #7dd3fc;
      --live: rgba(31, 143, 78, 0.12);
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: "Segoe UI", "Hiragino Sans", "Yu Gothic UI", sans-serif;
      background: var(--bg);
      color: var(--text);
    }}
    .wrap {{ max-width: 1600px; margin: 0 auto; padding: 20px; }}
    .header {{
      background: linear-gradient(135deg, #1a2030 0%, #10141c 100%);
      border: 1px solid var(--border);
      border-radius: 12px;
      padding: 20px 24px;
      margin-bottom: 16px;
    }}
    h1 {{ margin: 0 0 8px; font-size: 1.5rem; }}
    .meta {{ color: var(--muted); font-size: 0.95rem; line-height: 1.7; }}
    .status {{
      display: inline-block;
      padding: 4px 10px;
      border-radius: 999px;
      font-size: 0.85rem;
      font-weight: 700;
      margin-top: 8px;
    }}
    .status-live {{ background: rgba(31, 143, 78, 0.2); color: #7dffb2; }}
    .status-wait {{ background: rgba(255, 180, 0, 0.15); color: #ffd166; }}
    .toolbar {{
      display: flex;
      gap: 12px;
      align-items: center;
      margin-bottom: 16px;
      flex-wrap: wrap;
    }}
    .toolbar select, .toolbar input {{
      padding: 10px 14px;
      border-radius: 8px;
      border: 1px solid var(--border);
      background: var(--panel);
      color: var(--text);
      font-size: 1rem;
    }}
    .toolbar select {{ min-width: 220px; }}
    .toolbar input {{ flex: 1; min-width: 220px; }}
    .toolbar .info {{ color: var(--muted); font-size: 0.9rem; }}
    .nav-link {{
      color: #9fd0ff;
      text-decoration: none;
      padding: 10px 14px;
      border: 1px solid var(--border);
      border-radius: 8px;
      white-space: nowrap;
    }}
    .nav-link:hover {{ background: rgba(255, 255, 255, 0.05); }}
    .table-wrap {{
      overflow: auto;
      border: 1px solid var(--border);
      border-radius: 12px;
      background: var(--panel);
    }}
    table {{
      width: 100%;
      border-collapse: collapse;
      min-width: 1000px;
      font-size: 0.92rem;
    }}
    th, td {{
      padding: 10px 12px;
      border-bottom: 1px solid var(--border);
      text-align: left;
      vertical-align: top;
    }}
    tbody td {{ font-size: calc(1em - 1pt); }}
    th {{
      position: sticky;
      top: 0;
      background: #232b3d;
      color: #d7e0f1;
      font-size: 0.8rem;
      letter-spacing: 0.04em;
      text-transform: uppercase;
      z-index: 1;
    }}
    tr:hover {{ background: rgba(255, 255, 255, 0.03); }}
    tr.row-live {{ background: var(--live); }}
    .num {{ text-align: center; white-space: nowrap; }}
    .car-no {{ font-weight: 700; color: #fff; }}
    .class-badge {{
      display: inline-block;
      background: rgba(31, 143, 78, 0.2);
      color: #8dffb8;
      padding: 2px 8px;
      border-radius: 6px;
      font-weight: 700;
      font-size: 0.82rem;
    }}
    .driver {{ font-weight: 700; min-width: 140px; }}
    .team, .car {{ color: var(--muted); min-width: 160px; }}
    .lap {{
      font-family: Consolas, "Courier New", monospace;
      white-space: nowrap;
      text-align: right;
      min-width: 90px;
    }}
    .lap-best {{ color: var(--best); font-weight: 700; }}
    .current-badge {{
      display: inline-block;
      background: rgba(225, 6, 0, 0.2);
      color: #ff8f8f;
      padding: 2px 8px;
      border-radius: 6px;
      font-size: 0.82rem;
      font-weight: 700;
    }}
    .hidden {{ display: none; }}
    .footer {{
      margin-top: 14px;
      color: var(--muted);
      font-size: 0.85rem;
      line-height: 1.6;
    }}
  </style>
</head>
<body>
  <div class="wrap">
    <div class="header">
      <h1>ドライバー結果一覧</h1>
      <div class="meta">
        レース: {html.escape(str(master.get("RaceNameL", "")))} ({html.escape(str(master.get("RaceYear", "")))})<br>
        クラス: {html.escape(class_label)}<br>
        更新時刻: {html.escape(format_update_time(live.get("UpdateTime", "")))}<br>
        保存時刻: {html.escape(datetime.now().isoformat(timespec="seconds"))}
      </div>
      <span class="status {status_class}">{html.escape(status)}</span>
    </div>

    <div class="toolbar">
      <select id="driverFilter">
        {"".join(option_rows)}
      </select>
      <input id="search" type="search" placeholder="ドライバー・車番・チーム名で検索...">
      <a class="nav-link" href="index.html">車両一覧</a>
      <div class="info">表示: <span id="count">{len(driver_rows)}</span> 件 / {interval}秒ごとに自動更新</div>
    </div>

    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Driver</th>
            <th>Slot</th>
            <th>状態</th>
            <th>No.</th>
            <th>Class</th>
            <th>BestLap</th>
            <th>LastLap</th>
            <th>Laps</th>
            <th>POS</th>
            <th>PIC</th>
            <th>Team</th>
            <th>Car</th>
          </tr>
        </thead>
        <tbody id="rows">
          {"".join(table_rows)}
        </tbody>
      </table>
    </div>

    <div class="footer">
      BestLap / LastLap は車両単位のタイムです（同じ車のドライバーは同じ値が表示されます）<br>
      データ取得元: supertaikyu.live / このページはローカル保存された HTML です
    </div>
  </div>

  <script>
    const driverFilter = document.getElementById("driverFilter");
    const search = document.getElementById("search");
    const rows = Array.from(document.querySelectorAll("#rows tr"));
    const count = document.getElementById("count");

    function filterRows() {{
      const selectedDriver = driverFilter.value;
      const query = search.value.trim().toLowerCase();
      let visible = 0;

      for (const row of rows) {{
        const driverName = row.dataset.driver || "";
        const haystack = (row.dataset.search || "").toLowerCase();
        const matchDriver = !selectedDriver || driverName === selectedDriver;
        const matchSearch = !query || haystack.includes(query);
        const show = matchDriver && matchSearch;
        row.classList.toggle("hidden", !show);
        if (show) visible++;
      }}

      count.textContent = String(visible);
    }}

    driverFilter.addEventListener("change", filterRows);
    search.addEventListener("input", filterRows);
  </script>
</body>
</html>
"""

    html_path.write_text(page, encoding="utf-8")
    return html_path


def persist_outputs(
    args: argparse.Namespace,
    rows: list[TimingRow],
    master: dict[str, Any],
    live: dict[str, Any],
) -> dict[str, Path]:
    saved_paths = save_latest_files(args.output_dir, rows, master, live, args.class_filter)
    if args.history:
        saved_paths["history"] = append_history(
            args.output_dir, rows, master, live, args.class_filter
        )
    if args.html:
        driver_rows = build_driver_rows(master, live, args.class_filter, args.english)
        saved_paths["html"] = save_html_file(
            args.output_dir,
            rows,
            master,
            live,
            args.class_filter,
            args.interval,
        )
        saved_paths["drivers_html"] = save_drivers_html(
            args.output_dir,
            driver_rows,
            master,
            live,
            args.class_filter,
            args.interval,
        )
    return saved_paths


def run_once(args: argparse.Namespace) -> int:
    session = requests.Session()
    session.headers.update({"User-Agent": "supertaikyu-timing-fetcher/1.0"})

    try:
        rows, master, live = load_timing_data(session, args.class_filter, args.english)
    except RuntimeError as exc:
        print(f"エラー: {exc}", file=sys.stderr)
        return 1

    if args.search:
        print_search_results(search_rows(rows, args.search), args.search)
        return 0

    saved_paths = persist_outputs(args, rows, master, live)

    if args.html:
        print(f"HTML: {saved_paths['html']}")
    if not args.quiet:
        header = [
            "スーパー耐久 ライブタイミング取得ツール",
            f"状態: {race_status_text(live)}",
            f"更新時刻: {format_update_time(live.get('UpdateTime', ''))}",
            f"レース: {master.get('RaceNameL', '-')} ({master.get('RaceYear', '-')})",
            f"クラス: {args.class_filter or '全クラス'}",
            f"保存先: {args.output_dir}",
            f"JSON: {saved_paths['json']}",
            f"CSV : {saved_paths['csv']}",
        ]
        if "html" in saved_paths:
            header.append(f"HTML: {saved_paths['html']}")
        if "drivers_html" in saved_paths:
            header.append(f"Drivers HTML: {saved_paths['drivers_html']}")
        print_table(rows, header)
    return 0


def run_watch(args: argparse.Namespace) -> int:
    session = requests.Session()
    session.headers.update({"User-Agent": "supertaikyu-timing-fetcher/1.0"})

    html_path = args.output_dir / "index.html"
    browser_opened = False

    if args.html and not args.quiet:
        print("HTML 表示モードを開始します。ブラウザを見ながらバックグラウンド更新します。")
        print("終了は Ctrl+C です。")
        time.sleep(1)
    elif not args.quiet:
        print("自動更新モードを開始します。終了は Ctrl+C です。")
        time.sleep(1)

    while True:
        try:
            rows, master, live = load_timing_data(session, args.class_filter, args.english)
            if args.search:
                rows = search_rows(rows, args.search)

            saved_paths = persist_outputs(args, rows, master, live)

            if args.open_browser and args.html and not browser_opened and "html" in saved_paths:
                webbrowser.open(saved_paths["html"].resolve().as_uri())
                browser_opened = True
                if not args.quiet:
                    print(f"ブラウザで HTML を開きました: {saved_paths['html']}")

            if args.quiet or args.html:
                status = race_status_text(live)
                updated = format_update_time(live.get("UpdateTime", ""))
                html_info = ""
                if "html" in saved_paths:
                    html_info = f" | HTML: {saved_paths['html']}"
                if "drivers_html" in saved_paths:
                    html_info += f" | Drivers: {saved_paths['drivers_html']}"
                print(
                    f"[{datetime.now().strftime('%H:%M:%S')}] {status} | 更新: {updated} | {len(rows)}件{html_info}"
                )
            else:
                clear_screen()
                header = [
                    "スーパー耐久 ライブタイミング取得ツール",
                    f"状態: {race_status_text(live)}",
                    f"更新時刻: {format_update_time(live.get('UpdateTime', ''))}",
                    f"レース: {master.get('RaceNameL', '-')} ({master.get('RaceYear', '-')})",
                    f"クラス: {args.class_filter or '全クラス'}",
                    f"自動更新: {args.interval}秒",
                ]
                if args.search:
                    header.append(f"検索: {args.search}")
                if "history" in saved_paths:
                    header.append(f"履歴: {saved_paths['history']}")
                header.append(f"保存先: {args.output_dir}")
                header.append(f"JSON: {saved_paths['json']}")
                header.append(f"CSV : {saved_paths['csv']}")
                if "html" in saved_paths:
                    header.append(f"HTML: {saved_paths['html']}")
                print_table(rows, header)
                print("Ctrl+C で終了")

            time.sleep(args.interval)
        except KeyboardInterrupt:
            print("\n終了します。")
            return 0
        except RuntimeError as exc:
            if not args.quiet and not args.html:
                clear_screen()
            print(f"取得エラー: {exc}")
            print(f"{args.interval}秒後に再試行します...")
            time.sleep(args.interval)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Fetch Super Taikyu live timing data from supertaikyu.live JSON.",
    )
    parser.add_argument(
        "--class",
        dest="class_filter",
        default=DEFAULT_CLASS,
        help=f"Filter by class (default: {DEFAULT_CLASS}). Use ALL for every class.",
    )
    parser.add_argument(
        "--search",
        help="Search by driver, team, or car name (partial match).",
    )
    parser.add_argument(
        "--english",
        action="store_true",
        help="Use English driver names.",
    )
    parser.add_argument(
        "--interval",
        type=int,
        default=DEFAULT_INTERVAL,
        help=f"Auto-refresh interval in seconds (default: {DEFAULT_INTERVAL}).",
    )
    parser.add_argument(
        "--once",
        action="store_true",
        help="Fetch once and exit.",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=DEFAULT_OUTPUT_DIR,
        help=f"Directory to save JSON/CSV data (default: {DEFAULT_OUTPUT_DIR}).",
    )
    parser.add_argument(
        "--history",
        action="store_true",
        help="Append each fetch to history CSV.",
    )
    parser.add_argument(
        "--html",
        action="store_true",
        help="Generate data/index.html for browser viewing.",
    )
    parser.add_argument(
        "--quiet",
        action="store_true",
        help="Minimize terminal output.",
    )
    parser.add_argument(
        "--open-browser",
        action="store_true",
        help="Open index.html in the default browser on start.",
    )
    return parser


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8")

    parser = build_parser()
    args = parser.parse_args()

    if str(args.class_filter).upper() == "ALL":
        args.class_filter = None
    elif args.class_filter not in ALL_CLASSES:
        print(
            f"警告: 未知のクラス '{args.class_filter}' です。データがない場合があります。",
            file=sys.stderr,
        )

    if args.once:
        return run_once(args)
    return run_watch(args)


if __name__ == "__main__":
    raise SystemExit(main())
