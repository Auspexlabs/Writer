#!/usr/bin/env bash
# Publish the website, and optionally a Mac release, to the website server.
#
#   website/deploy.sh                                               the site only
#   website/deploy.sh --dmg desktop/dist/Writer-1.0.0-mac.dmg [--updates desktop/dist/updates]
#
# The first run installs Caddy. The server and its SSH key come from website/.deploy.local
# (not committed): WRITER_HOST=user@host and WRITER_KEY=/path/to/private-key.
set -euo pipefail

local_conf="$(cd "$(dirname "$0")" && pwd)/.deploy.local"
[ -f "$local_conf" ] && . "$local_conf"
host=${WRITER_HOST:?set WRITER_HOST in website/.deploy.local}
ssh_="ssh -i ${WRITER_KEY:?set WRITER_KEY in website/.deploy.local} -o StrictHostKeyChecking=accept-new"
rsync_() { rsync -az --no-owner --no-group -e "$ssh_" "$@"; }
dmg= updates=
while [ $# -gt 0 ]; do
  case $1 in
    --dmg) dmg=$(realpath "$2"); shift 2 ;;        # absolute before the cd below
    --updates) updates=$(realpath "$2"); shift 2 ;;
    *) echo "usage: $0 [--dmg Writer-<version>-mac.dmg] [--updates dir]" >&2; exit 2 ;;
  esac
done
cd "$(dirname "$0")"

# Caddy and the folders; does nothing after the first run.
$ssh_ "$host" 'set -e
  command -v caddy >/dev/null || { apt-get update -qq && DEBIAN_FRONTEND=noninteractive apt-get install -yqq caddy rsync >/dev/null; }
  mkdir -p /srv/writer/site/download /srv/writer/site/updates/mac
  [ -f /srv/writer/redirects.caddy ] || echo "# written by website/deploy.sh --dmg" > /srv/writer/redirects.caddy'

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
  $ssh_ "$host" "echo 'redir /download/mac /download/$name 302' > /srv/writer/redirects.caddy
    printf '{\"version\":\"%s\",\"size\":%s,\"date\":\"%s\"}\n' '$version' $(stat -f%z "$dmg") $(date +%F) > /srv/writer/site/download/latest.json"
fi

if [ -n "$updates" ]; then
  # Archive and signature first, manifest last: the app never sees a manifest without its files.
  rsync_ --exclude latest.json "$updates"/ "$host:/srv/writer/site/updates/mac/"
  rsync_ "$updates/latest.json" "$host:/srv/writer/site/updates/mac/"
fi

# A broken Caddyfile is rejected here (with Caddy's error) and the running one stays.
$ssh_ "$host" 'out=$(caddy validate --adapter caddyfile --config /srv/writer/Caddyfile.new 2>&1) || { echo "$out" >&2; exit 1; }
  mv /srv/writer/Caddyfile.new /etc/caddy/Caddyfile && systemctl reload caddy'
echo "published to $host"
