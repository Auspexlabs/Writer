#!/bin/bash
# Windows release files for the GitHub release and the website, from the NSIS installer Writer_<version>_x64-setup.exe
# that .github/workflows/windows.yml builds on Windows (gh run download, see README.md). Run from desktop/ on the
# release Mac:
#
#   scripts/release-win.sh <Writer_<version>_x64-setup.exe>
#
# Copies it to dist/Writer-<version>-windows-x64-setup.exe, signs it with the in-app updater's key (the one
# release-mac.sh web uses: ../.signing/writer-updater.key, its password in the keychain item writer-updater), adds
# windows-x86_64 to dist/latest.json, before or after release-mac.sh web adds darwin-aarch64, and writes the website's
# dist/updates/windows/: the installer, its signature and a latest.json with website URLs.
# Overrides: UPDATER_KEY, UPDATER_PW, NOTES or NOTES_FILE (the release notes the update dialog shows).
set -euo pipefail
SRC=$(cd "$(dirname "${1:?用法：scripts/release-win.sh <Writer_<版本>_x64-setup.exe>}")" && pwd)/$(basename "$1")
cd "$(dirname "$0")/.."
OUT=dist
VERSION=$(node -p "require('./src-tauri/tauri.conf.json').version")
EXE="Writer-$VERSION-windows-x64-setup.exe"
# the signature is bound to this version (requireSignedVersion): an installer of another version would be offered forever
case "$(basename "$SRC")" in
  *_"$VERSION"_*|"$EXE") ;;
  *) echo "$(basename "$SRC") 不是 $VERSION 版的安装包（版本见 src-tauri/tauri.conf.json）。" >&2; exit 1 ;;
esac
UPDATER_KEY=${UPDATER_KEY:-../.signing/writer-updater.key}
[ -f "$UPDATER_KEY" ] || { echo "缺少更新签名私钥 $UPDATER_KEY（见 SIGNING.local.md）。" >&2; exit 1; }
[ -n "${UPDATER_PW:-}" ] || UPDATER_PW=$(security find-generic-password -s writer-updater -w) || { echo "钥匙串里缺少 writer-updater（更新签名私钥的密码）。" >&2; exit 1; }
NOTES=${NOTES:-$(cat "${NOTES_FILE:-/dev/null}")}
mkdir -p "$OUT"
[ "$SRC" -ef "$OUT/$EXE" ] || cp "$SRC" "$OUT/$EXE"
TAURI_SIGNING_PRIVATE_KEY_PASSWORD="$UPDATER_PW" npx tauri signer sign -f "$UPDATER_KEY" --app-version "$VERSION" "$OUT/$EXE" >/dev/null
node scripts/latest-json.mjs "$OUT/latest.json" windows-x86_64 "$VERSION" "$OUT/$EXE.sig" \
  "https://github.com/Auspexlabs/writer/releases/download/v$VERSION/$EXE" "$NOTES"
# the website's feed, which installed copies read first (website/deploy.sh --updates dist/updates)
mkdir -p "$OUT/updates/windows" && cp "$OUT/$EXE" "$OUT/$EXE.sig" "$OUT/updates/windows/"
node scripts/latest-json.mjs "$OUT/updates/windows/latest.json" windows-x86_64 "$VERSION" "$OUT/$EXE.sig" \
  "https://thewriter.cn/updates/windows/$EXE" "$NOTES"
echo "$OUT/$EXE (+ .sig), $OUT/latest.json, $OUT/updates/windows/"
