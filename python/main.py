import argparse
import csv
import json
import time
import threading
import hashlib
from http.server import BaseHTTPRequestHandler, HTTPServer
import os
import shutil
import sys
from datetime import datetime
from typing import Iterable, List, Dict, Set


def ensure_directory(path: str) -> None:
    if not os.path.isdir(path):
        os.makedirs(path, exist_ok=True)


def list_csv_files(directory: str) -> List[str]:
    if not os.path.isdir(directory):
        return []
    return [
        os.path.join(directory, name)
        for name in os.listdir(directory)
        if name.lower().endswith(".csv") and os.path.isfile(os.path.join(directory, name))
    ]


def safe_copy(src: str, dest_dir: str) -> str:
    base = os.path.basename(src)
    name, ext = os.path.splitext(base)
    candidate = os.path.join(dest_dir, base)
    counter = 1
    while os.path.exists(candidate):
        candidate = os.path.join(dest_dir, f"{name}_{counter}{ext}")
        counter += 1
    shutil.copy2(src, candidate)
    return candidate


def collect_from_source(source_dir: str, cache_dir: str) -> List[str]:
    collected: List[str] = []
    if not source_dir or not os.path.isdir(source_dir):
        return collected
    ensure_directory(cache_dir)
    for entry in os.listdir(source_dir):
        src_path = os.path.join(source_dir, entry)
        if os.path.isfile(src_path) and entry.lower().endswith(".csv"):
            copied = safe_copy(src_path, cache_dir)
            collected.append(copied)
    return collected


def read_all_rows(csv_paths: Iterable[str]) -> (List[str], List[Dict[str, str]]):
    columns: List[str] = []
    seen: Set[str] = set()
    rows: List[Dict[str, str]] = []

    for path in csv_paths:
        try:
            with open(path, "r", encoding="utf-8-sig", newline="") as f:
                reader = csv.DictReader(f)
                # Union columns preserving order of first appearance
                for col in reader.fieldnames or []:
                    if col not in seen:
                        seen.add(col)
                        columns.append(col)
                for row in reader:
                    rows.append(dict(row))
        except Exception as ex:
            print(f"WARN: Failed to read {path}: {ex}")
    return columns, rows


def write_merged_csv(columns: List[str], rows: List[Dict[str, str]], output_path: str) -> None:
    ensure_directory(os.path.dirname(output_path))
    with open(output_path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=columns, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            complete = {col: row.get(col, "") for col in columns}
            writer.writerow(complete)


STATUS_CANDIDATE_NAMES: Set[str] = {"status", "state", "done", "fixed", "completed", "対応状況", "進捗", "完了"}


def normalize_status_kind(value: str) -> str:
    v = (value or "").strip().lower()
    if v in {"done", "fixed", "completed", "complete", "closed", "resolved", "ok", "true", "1", "完了", "対応済", "解決"}:
        return "done"
    if v in {"in progress", "progress", "working", "wip", "対応中", "進行中"}:
        return "doing"
    return "todo"


def status_label_from_kind(kind: str) -> str:
    if kind == "done":
        return "完了"
    if kind == "doing":
        return "対応中"
    return "未対応"


def detect_status_column(columns: List[str]) -> int:
    if not columns:
        return -1
    for idx, col in enumerate(columns):
        key = (col or "").strip().lower()
        if key in STATUS_CANDIDATE_NAMES:
            return idx
    return -1


def compute_row_key(columns: List[str], row: Dict[str, str], status_index: int) -> str:
    # Prefer id-like columns
    for cand in ("id", "Id", "ID", "issue_id", "ticket", "番号"):
        if cand in row and str(row.get(cand, "")).strip() != "":
            return f"id:{str(row.get(cand))}"
    # Build a fingerprint from all non-status columns
    pieces: List[str] = []
    for i, col in enumerate(columns):
        if i == status_index:
            continue
        val = row.get(col, "")
        pieces.append(f"{col}={val}")
    raw = "|".join(pieces)
    h = hashlib.sha1(raw.encode("utf-8", errors="ignore")).hexdigest()
    return f"sha1:{h}"


def load_overrides(overrides_path: str) -> Dict[str, str]:
    try:
        if os.path.isfile(overrides_path):
            with open(overrides_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict):
                    return {str(k): str(v) for k, v in data.items()}
    except Exception as ex:
        print(f"WARN: Failed to load overrides: {ex}")
    return {}


def save_overrides(overrides_path: str, data: Dict[str, str]) -> None:
    ensure_directory(os.path.dirname(overrides_path))
    tmp = overrides_path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, overrides_path)


def apply_overrides_to_rows(columns: List[str], rows: List[Dict[str, str]], overrides: Dict[str, str]) -> List[Dict[str, str]]:
    status_index = detect_status_column(columns)
    if status_index < 0:
        columns.append("status")
        status_index = len(columns) - 1
    status_col_name = columns[status_index]

    for row in rows:
        key = compute_row_key(columns, row, status_index)
        if key in overrides:
            kind = normalize_status_kind(overrides[key])
            row[status_col_name] = status_label_from_kind(kind)
        else:
            row.setdefault(status_col_name, row.get(status_col_name, ""))
    return rows


def generate_feedback_html(merged_csv_path: str, overrides_path: str, html_output_path: str, title: str = "Feedback CSV 一覧") -> None:
    ensure_directory(os.path.dirname(html_output_path))
    columns: List[str] = []
    rows_dict: List[Dict[str, str]] = []

    if os.path.isfile(merged_csv_path):
        with open(merged_csv_path, "r", encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            columns = list(reader.fieldnames or [])
            for row in reader:
                rows_dict.append(dict(row))
    # Apply overrides
    overrides = load_overrides(overrides_path)
    rows_dict = apply_overrides_to_rows(columns, rows_dict, overrides)

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    # Detect status column and build badges
    status_index = detect_status_column(columns)


    def badge_html(kind: str, text: str) -> str:
        if kind == "done":
            cls = "badge-done"
        elif kind == "doing":
            cls = "badge-doing"
        else:
            cls = "badge-todo"
        safe = (text or "").replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        return f"<span class=\"badge {cls}\">{safe or ('完了' if kind=='done' else ('対応中' if kind=='doing' else '未対応'))}</span>"

    total = len(rows_dict)
    cnt_done = 0
    cnt_doing = 0
    cnt_todo = 0

    # Build head
    table_head = "".join([f"<th>{c}</th>" for c in columns]) if columns else "<th>データなし</th>"
    action_head = "<th>操作</th>"

    # Build body with per-row status
    body_parts: List[str] = []
    if rows_dict:
        for row in rows_dict:
            status_kind = ""
            status_text = ""
            if status_index >= 0:
                status_col_name = columns[status_index]
                status_text = row.get(status_col_name, "")
                status_kind = normalize_status_kind(status_text)
            else:
                status_kind = ""

            if status_kind == "done":
                cnt_done += 1
            elif status_kind == "doing":
                cnt_doing += 1
            else:
                cnt_todo += 1

            tds: List[str] = []
            for i, col in enumerate(columns):
                cell = row.get(col, "")
                safe_cell = (cell or '').replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                if i == status_index:
                    tds.append(f"<td>{badge_html(status_kind or 'todo', status_text)}</td>")
                else:
                    # Preserve line breaks for multiline feedback field
                    if col == "修正または改善してほしい内容":
                        safe_cell = safe_cell.replace("\r\n", "<br>").replace("\n", "<br>")
                    tds.append(f"<td>{safe_cell}</td>")
            key = compute_row_key(columns, row, status_index)
            action_cell = (
                f"<td>"
                f"<button class=\"set-st btn\" data-key=\"{key}\" data-status=\"done\">完了</button> "
                f"<button class=\"set-st btn\" data-key=\"{key}\" data-status=\"doing\">対応中</button> "
                f"<button class=\"set-st btn\" data-key=\"{key}\" data-status=\"todo\">未対応</button>"
                f"</td>"
            )
            tr_class = f" class=\"st-{status_kind}\"" if status_kind else ""
            data_attr = f" data-status=\"{status_kind or 'todo'}\" data-key=\"{key}\""
            body_parts.append("<tr" + tr_class + data_attr + ">" + "".join(tds) + action_cell + "</tr>")
        table_body = "\n".join(body_parts)
    else:
        table_body = "<tr><td>データがありません</td></tr>"

    content = f"""<!DOCTYPE html>
<html lang=\"ja\">
<head>
    <meta charset=\"UTF-8\">
    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">
    <title>{title}</title>
    <link rel=\"stylesheet\" href=\"../css/styles.css\">
    <style>
        table {{ border-collapse: collapse; width: 100%; }}
        th, td {{ border: 1px solid #ddd; padding: 8px; }}
        th {{ background: #f2f2f2; text-align: left; }}
        caption {{ text-align: left; margin: 8px 0; font-weight: bold; }}
        .badges {{ display: flex; gap: 12px; align-items: center; margin: 8px 0 16px; }}
        .badge {{ display: inline-block; padding: 2px 8px; border-radius: 12px; font-size: 12px; color: #fff; }}
        .badge-done {{ background: #28a745; }}
        .badge-doing {{ background: #17a2b8; }}
        .badge-todo {{ background: #6c757d; }}
        .filter-btn {{ margin-left: 8px; padding: 4px 10px; }}
        .muted {{ color: #6c757d; }}
        .btn {{ padding: 2px 8px; margin-right: 6px; }}
    </style>
    <script defer src=\"../js/script.js\"></script>
</head>
<body>
<main class=\"container\">
    <h1>{title}</h1>
    <p>更新日時: {now}</p>
    <section class=\"badges\">
        <span>合計: <strong>{total}</strong></span>
        <span class=\"badge badge-done\">完了: {cnt_done}</span>
        <span class=\"badge badge-doing\">対応中: {cnt_doing}</span>
        <span class=\"badge badge-todo\">未対応: {cnt_todo}</span>
        <button id=\"togglePending\" class=\"filter-btn\" aria-pressed=\"false\">未完了のみ表示</button>
        <span class=\"muted\">ステータス列が無い場合は全て未対応として表示</span>
    </section>
    <section class=\"section\">
        <table aria-label=\"統合CSV\">
            <caption>統合ファイル: {os.path.basename(merged_csv_path) if os.path.basename(merged_csv_path) else ''}</caption>
            <thead>
                <tr>{table_head}{action_head}</tr>
            </thead>
            <tbody>
{table_body}
            </tbody>
        </table>
    </section>
    <p><a href=\"../index.html\">← 一覧へ戻る</a></p>
</main>
<script>
// シンプルな未完了フィルタ
(function() {{
  var btn = document.getElementById('togglePending');
  if (!btn) return;
  var pendingOnly = false;
  btn.addEventListener('click', function(){{
    pendingOnly = !pendingOnly;
    btn.setAttribute('aria-pressed', String(pendingOnly));
    btn.textContent = pendingOnly ? '全件表示' : '未完了のみ表示';
    var rows = document.querySelectorAll('tbody tr');
    rows.forEach(function(tr){{
      var st = tr.getAttribute('data-status') || 'todo';
      if (pendingOnly) {{
        tr.style.display = (st === 'todo' || st === 'doing') ? '' : 'none';
      }} else {{
        tr.style.display = '';
      }}
    }});
  }});
}})();

// ステータス更新（ローカルAPIにPOST）
(function() {{
  function updateBadge(tr, status){{
    if (!tr) return;
    tr.setAttribute('data-status', status);
    var badge = tr.querySelector('.badge');
    if (!badge) return;
    if (status === 'done') {{ badge.textContent = '完了'; badge.className = 'badge badge-done'; }}
    else if (status === 'doing') {{ badge.textContent = '対応中'; badge.className = 'badge badge-doing'; }}
    else {{ badge.textContent = '未対応'; badge.className = 'badge badge-todo'; }}
  }}
  document.addEventListener('click', function(e){{
    var t = e.target;
    if (!t.classList.contains('set-st')) return;
    var key = t.getAttribute('data-key');
    var status = t.getAttribute('data-status');
    if (!key || !status) return;
    fetch('http://127.0.0.1:8765/update-status', {{
      method: 'POST',
      headers: {{ 'Content-Type': 'application/json' }},
      body: JSON.stringify({{ key: key, status: status }})
    }}).then(function(r){{ return r.json(); }}).then(function(json){{
      if (json && json.ok) {{
        var tr = t.closest('tr');
        updateBadge(tr, status);
      }} else {{
        alert('更新に失敗しました');
      }}
    }}).catch(function(){{
      alert('ローカルサーバーに接続できません。run.cmd で起動中か確認してください。');
    }});
  }});
}})();
</script>
</body>
</html>
"""
    with open(html_output_path, "w", encoding="utf-8", newline="") as f:
        f.write(content)


def ensure_index_html(index_path: str) -> None:
    # Create minimal index with a link to feedback page if file is empty or missing
    if os.path.isfile(index_path):
        with open(index_path, "r", encoding="utf-8", errors="ignore") as f:
            content = f.read()
        if content.strip():
            return
    ensure_directory(os.path.dirname(index_path))
    html = """<!DOCTYPE html>
<html lang=\"ja\">
<head>
    <meta charset=\"UTF-8\">
    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">
    <title>Python Docs</title>
    <link rel=\"stylesheet\" href=\"css/styles.css\">
    <script defer src=\"js/script.js\"></script>
    <style>
        ul { list-style: none; padding-left: 0; }
        li { margin: 8px 0; }
    </style>
    </head>
<body>
<main class=\"container\">
    <h1>Python ドキュメント</h1>
    <ul>
        <li><a href=\"pages/feedback.html\">Feedback CSV 一覧</a></li>
    </ul>
</main>
</body>
</html>
"""
    with open(index_path, "w", encoding="utf-8", newline="") as f:
        f.write(html)


def load_config(base_dir: str) -> Dict[str, str]:
    cfg_path = os.path.join(base_dir, "_json", "config.json")
    try:
        if os.path.isfile(cfg_path):
            with open(cfg_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict):
                    return {str(k): str(v) for k, v in data.items() if isinstance(k, str) and isinstance(v, (str, int, float))}
    except Exception as ex:
        print(f"WARN: Failed to load config.json: {ex}")
    return {}


def run_merge_and_render(csv_dir: str, cache_dir: str, pages_dir: str) -> None:
    # Merge from cache_dir and csv_dir (excluding the merged output)
    candidates = list_csv_files(cache_dir) + [
        p for p in list_csv_files(csv_dir) if os.path.basename(p) != "feedback_merged.csv"
    ]
    candidates = sorted(set(candidates))

    print(f"Merging {len(candidates)} CSV(s)")
    columns, rows = read_all_rows(candidates)
    merged_path = os.path.join(csv_dir, "feedback_merged.csv")
    write_merged_csv(columns, rows, merged_path)
    print(f"Wrote merged CSV: {merged_path}")

    # Generate HTML
    feedback_html = os.path.join(pages_dir, "feedback.html")
    overrides_path = os.path.join(csv_dir, "status_overrides.json")
    generate_feedback_html(merged_path, overrides_path, feedback_html)
    print(f"Generated HTML: {feedback_html}")


def make_status_handler(overrides_path: str, csv_dir: str, cache_dir: str, pages_dir: str):
    class _Handler(BaseHTTPRequestHandler):
        def _set_headers(self, code=200):
            self.send_response(code)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.send_header('Access-Control-Allow-Methods', 'POST, OPTIONS')
            self.send_header('Access-Control-Allow-Headers', 'Content-Type')
            self.end_headers()

        def do_OPTIONS(self):
            self._set_headers(204)

        def do_POST(self):
            if self.path != '/update-status':
                self._set_headers(404)
                self.wfile.write(b'{"ok":false,"error":"not_found"}')
                return
            try:
                length = int(self.headers.get('Content-Length', '0'))
                body = self.rfile.read(length).decode('utf-8') if length > 0 else '{}'
                payload = json.loads(body or '{}')
                key = str(payload.get('key', ''))
                status = str(payload.get('status', ''))
                if not key or not status:
                    self._set_headers(400)
                    self.wfile.write(b'{"ok":false,"error":"bad_request"}')
                    return
                overrides = load_overrides(overrides_path)
                overrides[key] = status
                save_overrides(overrides_path, overrides)
                # Re-render HTML after update
                run_merge_and_render(csv_dir, cache_dir, pages_dir)
                self._set_headers(200)
                self.wfile.write(b'{"ok":true}')
            except Exception as ex:
                self._set_headers(500)
                msg = json.dumps({"ok": False, "error": str(ex)}).encode('utf-8')
                self.wfile.write(msg)
    return _Handler


def start_status_server(overrides_path: str, csv_dir: str, cache_dir: str, pages_dir: str, port: int = 8765) -> None:
    def _serve():
        try:
            server = HTTPServer(('127.0.0.1', port), make_status_handler(overrides_path, csv_dir, cache_dir, pages_dir))
            print(f"Status server listening on http://127.0.0.1:{port}")
            server.serve_forever()
        except OSError as ex:
            print(f"WARN: Status server not started: {ex}")
    t = threading.Thread(target=_serve, daemon=True)
    t.start()


def watch_and_collect(source_dir: str, cache_dir: str, csv_dir: str, pages_dir: str, interval_sec: float = 1.0) -> None:
    print(f"Watching: {source_dir} (interval={interval_sec}s). Press Ctrl+C to stop.")
    last_seen: Dict[str, tuple] = {}
    ensure_directory(cache_dir)
    ensure_directory(csv_dir)
    ensure_directory(pages_dir)

    # Initial merge/render so pages exist even before first drop
    run_merge_and_render(csv_dir, cache_dir, pages_dir)

    try:
        while True:
            copied_any = False
            if os.path.isdir(source_dir):
                for name in os.listdir(source_dir):
                    if not name.lower().endswith(".csv"):
                        continue
                    src_path = os.path.join(source_dir, name)
                    if not os.path.isfile(src_path):
                        continue
                    try:
                        stat = os.stat(src_path)
                        key = os.path.abspath(src_path)
                        size_mtime = (stat.st_size, int(stat.st_mtime))
                        prev = last_seen.get(key)
                        if prev == size_mtime:
                            # Stable for two checks → copy
                            dest_path = safe_copy(src_path, cache_dir)
                            try:
                                if os.path.isfile(dest_path) and os.path.getsize(dest_path) == stat.st_size:
                                    os.remove(src_path)
                                    copied_any = True
                                    print(f"Copied and removed source: {src_path} -> {dest_path}")
                                    last_seen.pop(key, None)
                                else:
                                    print(f"WARN: Size mismatch after copy, will retry: {src_path}")
                            except Exception as ex:
                                print(f"WARN: Post-copy validation/removal failed for {src_path}: {ex}")
                        else:
                            last_seen[key] = size_mtime
                    except FileNotFoundError:
                        # Source removed between list and stat
                        last_seen.pop(os.path.abspath(src_path), None)
                    except Exception as ex:
                        print(f"WARN: Failed to inspect {src_path}: {ex}")
            else:
                print(f"Waiting for directory: {source_dir}")

            if copied_any:
                run_merge_and_render(csv_dir, cache_dir, pages_dir)

            time.sleep(interval_sec)
    except KeyboardInterrupt:
        print("Stopping watch.")


def main(argv: List[str]) -> int:
    base_dir = os.path.dirname(os.path.abspath(__file__))
    docs_dir = os.path.join(base_dir, "_docs")
    csv_dir = os.path.join(docs_dir, "csv")
    cache_dir = os.path.join(csv_dir, "sources")
    pages_dir = os.path.join(docs_dir, "pages")

    ensure_directory(csv_dir)
    ensure_directory(cache_dir)
    ensure_directory(pages_dir)

    parser = argparse.ArgumentParser(description="Collect, merge, and render feedback CSVs")
    parser.add_argument("--source", dest="source", default=None, help="CSV格納のファイルサーバーディレクトリ パス (例: \\\\server\\share\\path)")
    parser.add_argument("--clear-cache", dest="clear_cache", action="store_true", help="sources キャッシュを削除してから収集")
    parser.add_argument("--watch", dest="watch", action="store_true", help="監視モードで起動 (1秒間隔で監視)")
    parser.add_argument("--interval", dest="interval", type=float, default=1.0, help="監視間隔(秒)")
    args = parser.parse_args(argv)

    if args.clear_cache and os.path.isdir(cache_dir):
        for name in os.listdir(cache_dir):
            path = os.path.join(cache_dir, name)
            if os.path.isfile(path):
                try:
                    os.remove(path)
                except Exception as ex:
                    print(f"WARN: Failed to remove {path}: {ex}")

    # Resolve source location priority: CLI > ENV > config file
    cfg = load_config(base_dir)
    cli_source = args.source
    env_source = os.environ.get("FEEDBACK_CSV_SOURCE")
    cfg_source = cfg.get("feedback_csv_source") or cfg.get("FEEDBACK_CSV_SOURCE") or cfg.get("source")
    resolved_source = cli_source or env_source or cfg_source or ""

    ensure_index_html(os.path.join(docs_dir, "index.html"))

    if args.watch:
        if not resolved_source:
            print("ERROR: 監視先が未設定です。--source か 環境変数 FEEDBACK_CSV_SOURCE か config.json を設定してください。")
            return 1
        overrides_path = os.path.join(csv_dir, "status_overrides.json")
        start_status_server(overrides_path, csv_dir, cache_dir, pages_dir)
        watch_and_collect(resolved_source, cache_dir, csv_dir, pages_dir, args.interval)
        return 0
    else:
        collected = []
        if resolved_source:
            print(f"Collecting CSVs from: {resolved_source}")
            collected = collect_from_source(resolved_source, cache_dir)
            print(f"Collected {len(collected)} file(s)")
        else:
            print("No --source specified. Skipping collection.")

        overrides_path = os.path.join(csv_dir, "status_overrides.json")
        start_status_server(overrides_path, csv_dir, cache_dir, pages_dir)
        run_merge_and_render(csv_dir, cache_dir, pages_dir)

    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))


