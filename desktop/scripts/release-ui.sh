#!/bin/bash
# Interface-only release: when only ui/ changed since the last full release (website/, docs and tests do not count),
# copies of that version download the interface (about 1 MB) instead of a new app. Run from desktop/ on the commit to
# ship:
#
#   NOTES_FILE=notes.txt scripts/release-ui.sh
#
# base = version in src-tauri/tauri.conf.json, the last full release, whose commit carries the local tag v<base>
# (README.md). Refuses when anything else changed since that tag: that needs a full release. build = 1 + the highest
# local tag v<base>-ui.<n>, and the run tags HEAD v<base>-ui.<build>. Tags are shared by every worktree, survive a
# cleared dist/ and show which commit each build came from; the repo has no remote, so they stay on this Mac.
# Writes, from HEAD's ui/ without tests/ and tools/: dist/Writer-ui-<base>-<build>.tar.gz (+ .sig, signed with the
# updater key for the version <base>-ui.<build>), dist/ui.json for the GitHub release v<base>, and dist/updates/ui/ for
# the website (website/deploy.sh --updates dist/updates). Uploads nothing.
# Overrides: UPDATER_KEY, UPDATER_PW, NOTES or NOTES_FILE (the notes the app shows).
set -euo pipefail
cd "$(dirname "$0")/.."
OUT=dist
BASE=$(node -p "require('./src-tauri/tauri.conf.json').version")
git -C .. rev-parse -q --verify "refs/tags/v$BASE" >/dev/null ||
  { echo "没有本地标签 v$BASE：请在构建 $BASE 的那个提交上打上它（git tag v$BASE <提交>）。" >&2; exit 1; }
dirty=$(git -C .. status --porcelain -- ui)
[ -z "$dirty" ] || { printf 'ui/ 有未提交的改动（界面包取自 HEAD），请先提交：\n%s\n' "$dirty" >&2; exit 1; }
# --no-renames: a file moved from src/ to docs/ still shows up as removed from src/
app=$(git -C .. diff --no-renames --name-only "v$BASE" HEAD -- . ':!ui' ':!website' ':!docs' ':!tests' ':!*.md')
[ -z "$app" ] || { printf '自 v%s 以来不只是界面有改动，需要完整发布（提高版本号，release-mac.sh web + release-win.sh，见 README.md）：\n%s\n' "$BASE" "$app" >&2; exit 1; }

UPDATER_KEY=${UPDATER_KEY:-../.signing/writer-updater.key}
[ -f "$UPDATER_KEY" ] || { echo "缺少更新签名私钥 $UPDATER_KEY（见 SIGNING.local.md）。" >&2; exit 1; }
[ -n "${UPDATER_PW:-}" ] || UPDATER_PW=$(security find-generic-password -s writer-updater -w) || { echo "钥匙串里缺少 writer-updater（更新签名私钥的密码）。" >&2; exit 1; }
NOTES=${NOTES:-$(cat "${NOTES_FILE:-/dev/null}")}
BUILD=$(git -C .. tag -l "v$BASE-ui.*" | awk -F. '$NF + 0 > max { max = $NF + 0 } END { print max + 1 }')
TGZ="Writer-ui-$BASE-$BUILD.tar.gz"
mkdir -p "$OUT/updates/ui"
git -C .. archive --format=tar.gz --prefix=ui/ HEAD:ui -- ':!tests' ':!tools' > "$OUT/$TGZ"
TAURI_SIGNING_PRIVATE_KEY_PASSWORD="$UPDATER_PW" npx tauri signer sign -f "$UPDATER_KEY" --app-version "$BASE-ui.$BUILD" "$OUT/$TGZ" >/dev/null
cp "$OUT/$TGZ" "$OUT/$TGZ.sig" "$OUT/updates/ui/"
DATE=$(date -u +%Y-%m-%dT%H:%M:%SZ)
ui_json() { # ui_json <file> <archive url>: the manifest the app reads
  node -e 'const fs = require("fs"), [file, base, build, pub_date, url, sig, notes] = process.argv.slice(1);
    const signature = fs.readFileSync(sig, "utf8").trim();
    fs.writeFileSync(file, JSON.stringify({ base, build: +build, notes, pub_date, url, signature }, null, 2) + "\n")' \
    "$1" "$BASE" "$BUILD" "$DATE" "$2" "$OUT/$TGZ.sig" "$NOTES"
}
ui_json "$OUT/ui.json" "https://github.com/Auspexlabs/writer/releases/download/v$BASE/$TGZ"
ui_json "$OUT/updates/ui/ui.json" "https://thewriter.cn/updates/ui/$TGZ"
git -C .. tag "v$BASE-ui.$BUILD"
echo "$OUT/$TGZ (+ .sig), $OUT/ui.json, $OUT/updates/ui/ (build $BUILD of $BASE, tag v$BASE-ui.$BUILD)"
echo "发布：gh release upload v$BASE -R Auspexlabs/writer $OUT/ui.json $OUT/$TGZ --clobber，再 ../website/deploy.sh --updates $OUT/updates"
