#!/usr/bin/env python3
"""Synthetic load test for the Package Log tab.

Development tool only: it is not part of the remote-update payload.

It imports Log_checker in-process with an isolated data directory, starts only
the Package Log emitter and SQLite writer (no adb, no device threads, no
app-open audit), and feeds synthetic rows through _append_package_log_row at a
fixed rate.  It reports what reached the UI channel and what reached SQLite.

    python3 tools/package_log_load_test.py --rate 2000 --seconds 20
    python3 tools/package_log_load_test.py --rate 2000 --seconds 600 --serve 5014

With --serve the real UI is served on 127.0.0.1:<port> so the browser side can
be measured while the synthetic stream runs.
"""

from __future__ import annotations

import argparse
import json
import os
import random
import sqlite3
import sys
import tempfile
import threading
import time
import types
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PACKAGE_ID = "com.eventinspector.loadtest"
DEVICE_ID = "LOADTEST-DEVICE"


def _prepare_environment(data_dir: str) -> None:
    os.environ["EVENTINSPECTOR_DATA_DIR"] = data_dir
    if str(ROOT) not in sys.path:
        sys.path.insert(0, str(ROOT))
    sys.modules.setdefault("webview", types.ModuleType("webview"))


def _synthetic_row(seq: int, rng: random.Random):
    now = time.time()
    stamp = time.strftime("%m-%d %H:%M:%S", time.localtime(now)) + f".{int(now * 1000) % 1000:03d}"
    level = rng.choices("VDIWE", weights=(5, 40, 40, 10, 5))[0]
    tag = rng.choice(("Unity", "AppLovinSdk", "CloudX", "Firebase", "ActivityManager", "chromium"))
    if rng.random() < 0.02:
        payload = json.dumps({"seq": seq, "items": ["x" * 40] * 90})
        message = f"[LoadTest] large payload {payload}"
    else:
        message = f"[LoadTest] seq={seq} " + "y" * rng.randint(40, 160)
    raw = f"{stamp}  4321  4400 {level} {tag}: {message}"
    time_display = stamp.split(" ", 1)[1]
    return raw, stamp, time_display, level, tag, message, level in {"E", "F"}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--rate", type=int, default=2000, help="rows per second")
    parser.add_argument("--seconds", type=float, default=20.0, help="generation duration")
    parser.add_argument("--serve", type=int, default=0, help="also serve the UI on this port")
    parser.add_argument("--start-delay", type=float, default=8.0,
                        help="with --serve: seconds to wait before streaming, so the platform can be chosen first "
                             "(choosing a platform resets the Package Log target)")
    parser.add_argument("--data-dir", default="", help="isolated data dir (default: new temp dir)")
    parser.add_argument("--json", action="store_true", help="print only the JSON summary")
    args = parser.parse_args()

    data_dir = args.data_dir or tempfile.mkdtemp(prefix="eventinspector_loadtest_")
    _prepare_environment(data_dir)
    import Log_checker as lc  # noqa: E402  (must follow the environment setup)

    if not str(lc.PACKAGE_LOG_DB_PATH).startswith(os.path.abspath(data_dir)):
        print(f"Refusing to run: Package Log DB is outside the isolated dir: {lc.PACKAGE_LOG_DB_PATH}", file=sys.stderr)
        return 2

    emits = []
    seen_ui_keys = set()
    original_emit = lc.socketio.emit

    def capturing_emit(event, payload=None, *a, **kw):
        if event in ("package_log_cache", "package_log_delta"):
            started = time.perf_counter()
            encoded = json.dumps(payload, ensure_ascii=False)
            encode_ms = (time.perf_counter() - started) * 1000
            rows = payload if isinstance(payload, list) else (payload or {}).get("rows", [])
            for row in rows:
                seen_ui_keys.add(row.get("seq") or row.get("log") or row.get("message"))
            emits.append({"t": time.time(), "rows": len(rows), "bytes": len(encoded.encode("utf-8")), "encode_ms": encode_ms})
        if args.serve:
            return original_emit(event, payload, *a, **kw)
        return None

    lc.socketio.emit = capturing_emit
    lc.active_platform = "android"
    lc._init_package_log_db()
    threading.Thread(target=lc._package_log_db_writer, daemon=True).start()
    threading.Thread(target=lc.package_log_emitter, daemon=True).start()
    if args.serve:
        threading.Thread(
            target=lambda: lc.socketio.run(lc.app, host="127.0.0.1", port=args.serve, use_reloader=False, allow_unsafe_werkzeug=True),
            daemon=True,
        ).start()
        print(f"Serving UI on http://127.0.0.1:{args.serve}: choose the platform and open the Package Log tab; "
              f"stream starts in {args.start_delay:g}s", flush=True)
        time.sleep(args.start_delay)

    lc.spl({"package_id": PACKAGE_ID})
    session_id = lc.active_package_log_session_id
    rng = random.Random(7)
    produced = 0
    produced_keys = []
    start = time.perf_counter()
    end = start + args.seconds
    while True:
        now = time.perf_counter()
        if now >= end:
            break
        due = int((now - start) * args.rate)
        while produced < due:
            produced += 1
            row = _synthetic_row(produced, rng)
            produced_keys.append(row[0])
            lc._append_package_log_row(DEVICE_ID, *row)
        time.sleep(0.005)
    gen_seconds = time.perf_counter() - start

    time.sleep(2.5)  # let the emitter and writer catch up
    cache_rows = len(lc.package_log_cache)
    dropped_db_queue = lc.package_log_db_dropped_rows
    lc.spl({"package_id": ""})
    time.sleep(1.5)

    conn = sqlite3.connect(lc.PACKAGE_LOG_DB_PATH)
    db_rows = conn.execute("SELECT COUNT(*) FROM package_log_entries WHERE session_id = ?", (session_id,)).fetchone()[0]
    conn.close()

    ui_keys = {k for k in seen_ui_keys if k}
    reached_ui = sum(1 for k in produced_keys if k in ui_keys) if not any(isinstance(k, int) for k in ui_keys) else len(ui_keys)
    stream_emits = [e for e in emits if e["rows"]]
    total_bytes = sum(e["bytes"] for e in stream_emits)
    summary = {
        "rate_rows_per_s": args.rate,
        "seconds": round(gen_seconds, 1),
        "produced_rows": produced,
        "reached_ui_rows": reached_ui,
        "never_reached_ui": produced - reached_ui,
        "db_rows": db_rows,
        "db_missing_rows": produced - db_rows,
        "db_queue_dropped": dropped_db_queue,
        "server_cache_rows_at_end": cache_rows,
        "ui_emits": len(stream_emits),
        "ui_rows_per_emit_avg": round(sum(e["rows"] for e in stream_emits) / max(len(stream_emits), 1)),
        "ui_bytes_per_emit_avg_kb": round(total_bytes / max(len(stream_emits), 1) / 1024, 1),
        "ui_bytes_per_second_mb": round(total_bytes / max(gen_seconds, 0.001) / 1024 / 1024, 2),
        "json_encode_ms_avg": round(sum(e["encode_ms"] for e in stream_emits) / max(len(stream_emits), 1), 1),
        "data_dir": data_dir,
    }
    print(json.dumps(summary, indent=None if args.json else 2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
