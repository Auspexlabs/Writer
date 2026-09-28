#!/usr/bin/env bash
# Used by deploy.sh; installs the counter without touching its persistent database.
set -euo pipefail
cd "$(dirname "$0")"
[ ! -f .deploy.local ] || . ./.deploy.local
: "${WRITER_HOST:?set WRITER_HOST}" "${WRITER_KEY:?set WRITER_KEY}"
ssh_args=(-i "$WRITER_KEY" -o BatchMode=yes -o StrictHostKeyChecking=accept-new)

ssh "${ssh_args[@]}" "$WRITER_HOST" 'mkdir -p /srv/writer/download-stats'

scp -q "${ssh_args[@]}" stats/server.py stats/writer-download-stats.service "$WRITER_HOST:/srv/writer/download-stats/"
ssh "${ssh_args[@]}" "$WRITER_HOST" 'set -e
  command -v python3 >/dev/null || { apt-get update -qq && DEBIAN_FRONTEND=noninteractive apt-get install -yqq python3 >/dev/null; }
  getent group writer-stats-public >/dev/null || groupadd --system writer-stats-public
  install -d -m 2775 -o root -g writer-stats-public /srv/writer/public-stats
  chmod 755 /srv/writer/download-stats
  chmod 644 /srv/writer/download-stats/*
  install -m 644 /srv/writer/download-stats/writer-download-stats.service /etc/systemd/system/writer-download-stats.service
  systemctl daemon-reload
  systemctl enable writer-download-stats.service >/dev/null
  systemctl restart writer-download-stats.service
  python3 - <<'"'"'PY'"'"'
import json, time, urllib.request
for attempt in range(20):
    try:
        with urllib.request.urlopen("http://127.0.0.1:8787/healthz", timeout=2) as response:
            assert response.status == 200
        with open("/srv/writer/public-stats/data.json") as snapshot:
            assert "total" in json.load(snapshot)
        break
    except Exception:
        if attempt == 19:
            raise
        time.sleep(0.25)
PY
  runuser -u caddy -- test -r /srv/writer/public-stats/data.json'
echo 'Download counter ready. Public dashboard: https://thewriter.cn/stats/'
