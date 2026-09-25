#!/usr/bin/env bash
# Publish the website, and optionally a release, to the website server.
#
#   website/deploy.sh                                               the site only
#   website/deploy.sh [--dmg desktop/dist/Writer-1.0.0-mac.dmg] [--exe desktop/dist/Writer-1.0.0-windows-x64-setup.exe]
#                     [--updates desktop/dist/updates]
#
# --dmg and --exe put an installer in /download/ behind /download/mac or /download/windows; --updates mirrors the
# in-app updater's mac/ and windows/ folders and the interface-only ui/ (desktop/scripts/release-*.sh) to /updates/.
#
# The first run installs Caddy. The server and its SSH key come from website/.deploy.local
# (not committed): WRITER_HOST=user@host and WRITER_KEY=/path/to/private-key.
set -euo pipefail

local_conf="$(cd "$(dirname "$0")" && pwd)/.deploy.local"
[ -f "$local_conf" ] && . "$local_conf"
host=${WRITER_HOST:?set WRITER_HOST in website/.deploy.local}
ssh_="ssh -i ${WRITER_KEY:?set WRITER_KEY in website/.deploy.local} -o StrictHostKeyChecking=accept-new"
rsync_() { rsync -az --no-owner --no-group -e "$ssh_" "$@"; }
dmg= exe= updates=
while [ $# -gt 0 ]; do
  case $1 in
    --dmg) dmg=$(realpath "$2"); shift 2 ;;        # absolute before the cd below
    --exe) exe=$(realpath "$2"); shift 2 ;;
    --updates) updates=$(realpath "$2"); shift 2 ;;
    *) echo "usage: $0 [--dmg Writer-<version>-mac.dmg] [--exe Writer-<version>-windows-x64-setup.exe] [--updates dir]" >&2; exit 2 ;;
  esac
done
[ -z "$updates" ] || [ -d "$updates/mac" ] || [ -d "$updates/windows" ] || [ -d "$updates/ui" ] || { echo "expected mac/, windows/ or ui/ in $updates" >&2; exit 2; }
cd "$(dirname "$0")"

# /download/<platform> → <file>: that platform's line in redirects.caddy, the other platform's stays
redirect() { echo "f=/srv/writer/redirects.caddy; { grep -v ' /download/$1 ' \$f; echo 'redir /download/$1 /download/$2 302'; } > \$f.new && mv \$f.new \$f"; }

# Caddy and the folders; does nothing after the first run.
$ssh_ "$host" 'set -e
  command -v caddy >/dev/null || { apt-get update -qq && DEBIAN_FRONTEND=noninteractive apt-get install -yqq caddy rsync >/dev/null; }
  mkdir -p /srv/writer/site/download /srv/writer/site/updates/mac /srv/writer/site/updates/windows /srv/writer/site/updates/ui
  [ -f /srv/writer/redirects.caddy ] || echo "# written by website/deploy.sh --dmg and --exe" > /srv/writer/redirects.caddy'

# download/ and updates/ exist only on the server, so --delete leaves them alone.
rsync_ --delete --exclude /download --exclude /updates dist/ "$host:/srv/writer/site/"
rsync_ Caddyfile "$host:/srv/writer/Caddyfile.new"

if [ -n "$dmg" ]; then
  name=$(basename "$dmg")                      # Writer-1.0.0-mac.dmg
  version=${name#Writer-}; version=${version%-mac.dmg}
  [ "$version" != "$name" ] || { echo "expected Writer-<version>-mac.dmg, got $name" >&2; exit 2; }
  # The file first, then the redirect and latest.json that point to it.
  # ponytail: every release stays on the server; prune old dmgs when the 40 GB disk matters.
  rsync_ "$dmg" $([ -f "$dmg.sha256" ] && echo "$dmg.sha256") "$host:/srv/writer/site/download/"
  $ssh_ "$host" "$(redirect mac "$name")
    printf '{\"version\":\"%s\",\"size\":%s,\"date\":\"%s\"}\n' '$version' $(stat -f%z "$dmg") $(date +%F) > /srv/writer/site/download/latest.json"
fi

if [ -n "$exe" ]; then
  name=$(basename "$exe")                      # Writer-1.0.0-windows-x64-setup.exe
  [[ $name == Writer-*-windows-x64-setup.exe ]] || { echo "expected Writer-<version>-windows-x64-setup.exe, got $name" >&2; exit 2; }
  # The file first, then the redirect that points to it.
  rsync_ "$exe" "$host:/srv/writer/site/download/"
  $ssh_ "$host" "$(redirect windows "$name")"
fi

if [ -n "$updates" ]; then
  # Archive or installer and signature first, manifest last: the app never sees a manifest without its files.
  for platform in mac windows ui; do
    [ -d "$updates/$platform" ] || continue
    rsync_ --exclude '*.json' "$updates/$platform"/ "$host:/srv/writer/site/updates/$platform/"
    rsync_ "$updates/$platform"/*.json "$host:/srv/writer/site/updates/$platform/" # latest.json, or ui.json in ui/
  done
fi

# A broken Caddyfile is rejected here (with Caddy's error) and the running one stays.
$ssh_ "$host" 'out=$(caddy validate --adapter caddyfile --config /srv/writer/Caddyfile.new 2>&1) || { echo "$out" >&2; exit 1; }
  mv /srv/writer/Caddyfile.new /etc/caddy/Caddyfile && systemctl reload caddy'
echo "published to $host"
