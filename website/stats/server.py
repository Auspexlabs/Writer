#!/usr/bin/env python3
"""Website download clicks only; no visitor identifiers or request logs."""
import argparse
import json
import os
import sqlite3
import tempfile
import threading
from contextlib import closing
from datetime import datetime, timedelta, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

SHANGHAI = timezone(timedelta(hours=8))
ORIGINS = {"https://thewriter.cn", "https://www.thewriter.cn"}


def now():
    return datetime.now(SHANGHAI)


class Store:
    def __init__(self, path):
        self.path = str(path)
        Path(path).parent.mkdir(parents=True, exist_ok=True)
        with closing(self.connect()) as db, db:
            db.execute("PRAGMA journal_mode=WAL")
            db.execute("CREATE TABLE IF NOT EXISTS daily (day TEXT NOT NULL, platform TEXT NOT NULL CHECK(platform IN ('mac','windows')), clicks INTEGER NOT NULL, PRIMARY KEY(day, platform))")
            db.execute("CREATE TABLE IF NOT EXISTS metadata (key TEXT PRIMARY KEY, value TEXT NOT NULL)")
            db.execute("INSERT OR IGNORE INTO metadata VALUES ('started_at', ?)", (now().isoformat(),))

    def connect(self):
        return sqlite3.connect(self.path, timeout=10)

    def record(self, platform, at=None):
        if platform not in ("mac", "windows"):
            raise ValueError("unknown platform")
        day = (at or now()).astimezone(SHANGHAI).date().isoformat()
        with closing(self.connect()) as db, db:
            db.execute("INSERT INTO daily VALUES (?, ?, 1) ON CONFLICT(day, platform) DO UPDATE SET clicks = clicks + 1", (day, platform))

    def summary(self, at=None):
        current = (at or now()).astimezone(SHANGHAI)
        today = current.date()
        first = (today - timedelta(days=29)).isoformat()
        with closing(self.connect()) as db, db:
            # One read transaction keeps the totals and daily rows consistent during clicks.
            db.execute("BEGIN")
            totals = dict(db.execute("SELECT platform, SUM(clicks) FROM daily GROUP BY platform"))
            rows = list(db.execute("SELECT day, platform, clicks FROM daily WHERE day >= ? AND day <= ?", (first, today.isoformat())))
            started = db.execute("SELECT value FROM metadata WHERE key = 'started_at'").fetchone()[0]
        daily = {}
        for offset in range(30):
            day = (today - timedelta(days=offset)).isoformat()
            daily[day] = {"day": day, "mac": 0, "windows": 0, "total": 0}
        for day, platform, count in rows:
            daily[day][platform] = count
            daily[day]["total"] += count
        days = list(daily.values())
        return {"started_at": started, "updated_at": current.isoformat(), "timezone": "Asia/Shanghai",
                "total": sum(totals.values()), "mac": totals.get("mac", 0), "windows": totals.get("windows", 0),
                "today": days[0]["total"], "last_7_days": sum(d["total"] for d in days[:7]),
                "last_30_days": sum(d["total"] for d in days), "daily": days}

    def publish(self, path, at=None):
        """Atomically replace the public aggregate snapshot; never publish the database."""
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        body = json.dumps(self.summary(at), ensure_ascii=False)
        temporary = None
        try:
            with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", dir=path.parent, delete=False) as output:
                temporary = Path(output.name)
                output.write(body)
                output.flush()
                os.fsync(output.fileno())
                os.fchmod(output.fileno(), 0o644)
            os.replace(temporary, path)
        finally:
            if temporary is not None:
                temporary.unlink(missing_ok=True)


class Handler(BaseHTTPRequestHandler):
    server_version = "WriterStats"
    sys_version = ""

    def setup(self):
        super().setup()
        self.connection.settimeout(5)

    def log_message(self, *_args):
        pass  # Do not store IP addresses, user agents, referrers or request bodies.

    def reply(self, status, value=None):
        body = b"" if value is None else json.dumps(value, ensure_ascii=False).encode()
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if self.path == "/healthz":
            self.reply(200, {"ok": True})
        else:
            self.reply(404)

    def do_POST(self):
        if self.path != "/api/download-clicks":
            return self.reply(404)
        if self.headers.get("Origin") not in ORIGINS:
            return self.reply(403)
        if self.headers.get("Transfer-Encoding"):
            return self.reply(400)
        try:
            length = int(self.headers.get("Content-Length", "0"))
        except ValueError:
            return self.reply(400)
        if not 0 < length <= 128:
            return self.reply(413)
        try:
            payload = json.loads(self.rfile.read(length))
            if not isinstance(payload, dict) or set(payload) != {"platform"}:
                return self.reply(400)
            self.server.store.record(payload["platform"])
        except (ValueError, UnicodeError):
            return self.reply(400)
        except sqlite3.Error:
            return self.reply(503)
        self.reply(204)


def make_server(path, port=8787):
    server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    server.store = Store(path)
    return server


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--db", required=True)
    parser.add_argument("--port", type=int, default=8787)
    parser.add_argument("--snapshot", help="Public JSON file, refreshed every 30 seconds")
    args = parser.parse_args()
    with make_server(args.db, args.port) as server:
        stop = threading.Event()
        if args.snapshot:
            server.store.publish(args.snapshot)

            def publish_periodically():
                while not stop.wait(30):
                    try:
                        server.store.publish(args.snapshot)
                    except (OSError, sqlite3.Error):
                        print("Could not refresh download snapshot; keeping previous file.", flush=True)

            threading.Thread(target=publish_periodically, daemon=True).start()
        try:
            server.serve_forever()
        finally:
            stop.set()
