#!/bin/bash
# Mac release builds. Run from desktop/:
#
#   scripts/release-mac.sh web            Developer ID: signed, notarized and stapled .app and .dmg, and the in-app
#                                         updater's .app.tar.gz (+ .sig) with latest.json for the GitHub release and,
#                                         in updates/mac/, for the website
#   scripts/release-mac.sh appstore       Mac App Store: sandboxed .app signed for the store, packaged as a signed .pkg
#   scripts/release-mac.sh sandbox-test   the store's sandbox signed with this Mac's Apple Development certificate, for
#                                         trying the sandboxed app locally; its own bundle id (cn.thewriter.sandboxtest),
#                                         so it never touches the real app's container
#
# Needs (all in the login keychain, created once by the account holder in Xcode › Settings › Accounts › Manage Certificates):
#   web       "Developer ID Application: …" certificate, and a notarytool profile made once with
#             xcrun notarytool store-credentials writer-notary --key AuthKey_XXXX.p8 --key-id XXXX --issuer <issuer id>;
#             the updater's signing key ../.signing/writer-updater.key, its password in the keychain item writer-updater
#   appstore  "Apple Distribution: …" and "3rd Party Mac Developer Installer: …" certificates, and the App Store
#             provisioning profile for cn.thewriter.app saved as ../.signing/Writer_Mac_App_Store.provisionprofile
# Overrides: SIGN_ID, INSTALLER_ID, TEAM_ID, NOTARY_PROFILE (default writer-notary), EXTRA_FEATURES (e.g. test-hooks),
# BUILD_NUMBER (the store's CFBundleVersion, which must grow with every upload; default: the commit count), UPDATER_KEY,
# NOTES or NOTES_FILE (the release notes the update dialog shows).
set -euo pipefail
mode=${1:-}
cd "$(dirname "$0")/.."
TAURI=src-tauri
OUT=dist
TARGET=${TARGET:-universal-apple-darwin} # TARGET=aarch64-apple-darwin for an Apple-silicon-only build
APP="$TAURI/target/$TARGET/release/bundle/macos/Writer.app"
VERSION=$(node -p "require('./$TAURI/tauri.conf.json').version")
mkdir -p "$OUT"

# Identities by SHA-1 (two certificates may share a name, e.g. after a renewal, and codesign refuses an ambiguous name),
# choosing the newest one Apple has not revoked: an app signed with a revoked certificate is killed on launch and
# macOS removes it.
pick() { # pick <name prefix> <find-identity policy args…>
  python3 - "$@" <<'PY'
import re, subprocess, sys, tempfile
name, policy = sys.argv[1], sys.argv[2:]
ids = subprocess.run(['security', 'find-identity', '-v', *policy], capture_output=True, text=True).stdout
hashes = {h for h, n in re.findall(r'\) ([0-9A-F]{40}) "([^"]*)"', ids) if n.startswith(name)}
pem = subprocess.run(['security', 'find-certificate', '-a', '-Z', '-p', '-c', name], capture_output=True, text=True).stdout
best = None
for h, cert in re.findall(r'SHA-1 hash: ([0-9A-F]+)\n(-----BEGIN CERTIFICATE-----.*?-----END CERTIFICATE-----)', pem, re.S):
    if h not in hashes: continue
    with tempfile.NamedTemporaryFile('w', suffix='.pem') as f:
        f.write(cert + '\n'); f.flush()
        if 'REVOKED' in subprocess.run(['security', 'verify-cert', '-c', f.name, '-R', 'ocsp'], capture_output=True, text=True).stdout: continue
        start = subprocess.run(['openssl', 'x509', '-noout', '-startdate', '-dateopt', 'iso_8601', '-in', f.name], capture_output=True, text=True).stdout
    if best is None or start > best[0]: best = (start, h)
print(best[1] if best else '')
PY
}
identity() { pick "$1" -p codesigning; }
installer_identity() { pick "3rd Party Mac Developer Installer"; }
team_of() { security find-identity -v | awk -v h="$1" '$2 == h' | sed -n 's/.*(\([A-Z0-9]\{10\}\))".*/\1/p' | head -1; }
need() { [ -n "$2" ] || { echo "缺少 $1。$3" >&2; exit 1; }; }

updater_files() { # the in-app updater's files, from the stapled app: a signed archive, and its entry in latest.json
  local tgz="Writer-$VERSION-mac.app.tar.gz"
  COPYFILE_DISABLE=1 tar -czf "$OUT/$tgz" -C "$(dirname "$APP")" "$(basename "$APP")" # no ._ files in the archive
  TAURI_SIGNING_PRIVATE_KEY_PASSWORD="$UPDATER_PW" npx tauri signer sign -f "$UPDATER_KEY" --app-version "$VERSION" "$OUT/$tgz" >/dev/null
  # keeps a windows-x86_64 entry of this version (scripts/release-win.sh)
  node scripts/latest-json.mjs "$OUT/latest.json" darwin-aarch64 "$VERSION" "$OUT/$tgz.sig" \
    "https://github.com/Auspexlabs/writer/releases/download/v$VERSION/$tgz" "$NOTES"
  # the website's feed, which installed copies read first (website/deploy.sh --updates dist/updates)
  mkdir -p "$OUT/updates/mac" && cp "$OUT/$tgz" "$OUT/$tgz.sig" "$OUT/updates/mac/"
  node scripts/latest-json.mjs "$OUT/updates/mac/latest.json" darwin-aarch64 "$VERSION" "$OUT/$tgz.sig" \
    "https://thewriter.cn/updates/mac/$tgz" "$NOTES"
}

notarize() { # notarize a zip/dmg/pkg and staple the ticket to what was submitted (or to the app for a zip)
  local file=$1 staple=${2:-$1}
  xcrun notarytool submit "$file" --keychain-profile "${NOTARY_PROFILE:-writer-notary}" --wait
  xcrun stapler staple "$staple"
}

sign_inner() { # the engine and helpers run in the app's sandbox (com.apple.security.inherit), then the app itself
  local id=$1 entitlements=$2
  for bin in "$APP"/Contents/MacOS/*; do
    [ "$(basename "$bin")" = writer-desktop ] && continue
    codesign --force --timestamp --sign "$id" --entitlements "$TAURI/entitlements/Sidecar.appstore.entitlements" "$bin"
  done
  codesign --force --timestamp --sign "$id" --entitlements "$entitlements" "$APP"
  codesign --verify --strict --verbose=2 "$APP"
}

case "$mode" in
web)
  SIGN_ID=${SIGN_ID:-$(identity "Developer ID Application")}
  need "Developer ID Application 证书" "$SIGN_ID" "请账号持有人在开发者网站 Certificates 里创建（node scripts/asc.mjs fetch DEVELOPER_ID_APPLICATION_G2 导入）。"
  UPDATER_KEY=${UPDATER_KEY:-../.signing/writer-updater.key}
  [ -f "$UPDATER_KEY" ] || { echo "缺少更新签名私钥 $UPDATER_KEY（见 SIGNING.local.md）。" >&2; exit 1; }
  UPDATER_PW=$(security find-generic-password -s writer-updater -w) || { echo "钥匙串里缺少 writer-updater（更新签名私钥的密码）。" >&2; exit 1; }
  NOTES=${NOTES:-$(cat "${NOTES_FILE:-/dev/null}")}
  node scripts/prepare-engine.mjs --release
  APPLE_SIGNING_IDENTITY="$SIGN_ID" npx tauri build --bundles app --target "$TARGET"
  codesign --verify --strict --deep --verbose=1 "$APP"
  # notarize and staple the app first, so it also opens offline once copied out of the disk image
  ditto -c -k --keepParent "$APP" "$OUT/Writer-notarize.zip"
  notarize "$OUT/Writer-notarize.zip" "$APP"
  rm -f "$OUT/Writer-notarize.zip"
  # the disk image is made from the stapled app as it is, so nothing re-signs it after notarization
  STAGE=$(mktemp -d); cp -R "$APP" "$STAGE/"; ln -s /Applications "$STAGE/Applications"
  DMG="$OUT/Writer-$VERSION-mac.dmg"; rm -f "$DMG"
  hdiutil create -volname Writer -srcfolder "$STAGE" -ov -format UDZO "$DMG" >/dev/null; rm -rf "$STAGE"
  codesign --force --timestamp --sign "$SIGN_ID" "$DMG"
  notarize "$DMG"
  (cd "$OUT" && shasum -a 256 "Writer-$VERSION-mac.dmg" | tee "Writer-$VERSION-mac.dmg.sha256")
  spctl --assess --type execute --verbose "$APP"
  spctl --assess --type open --context context:primary-signature --verbose "$DMG"
  updater_files
  ;;
appstore|sandbox-test)
  if [ "$mode" = appstore ]; then
    SIGN_ID=${SIGN_ID:-$(identity "Apple Distribution")}
    SIGN_ID=${SIGN_ID:-$(identity "3rd Party Mac Developer Application")}
    INSTALLER_ID=${INSTALLER_ID:-$(installer_identity)}
    need "Apple Distribution 证书" "$SIGN_ID" "请在 Xcode › 设置 › Accounts › Manage Certificates 里添加。"
    need "Mac Installer Distribution 证书" "$INSTALLER_ID" "请在 Xcode › 设置 › Accounts › Manage Certificates 里添加。"
    [ -f ../.signing/Writer_Mac_App_Store.provisionprofile ] || { echo "缺少描述文件 ../.signing/Writer_Mac_App_Store.provisionprofile（开发者网站 › Profiles › Mac App Store）" >&2; exit 1; }
    TEAM_ID=${TEAM_ID:-$(team_of "$SIGN_ID")}
    need "Team ID" "$TEAM_ID" "请设置 TEAM_ID。"
    sed "s/TEAM_ID/$TEAM_ID/g" "$TAURI/entitlements/Writer.appstore.entitlements" > "$TAURI/entitlements/.generated.entitlements"
    node scripts/prepare-engine.mjs --release
  else
    SIGN_ID=${SIGN_ID:-$(identity "Apple Development")}
    need "Apple Development 证书" "$SIGN_ID" "在 Xcode 里登录开发者账号即可。"
    # restricted entitlements (application / team identifier) need the store's profile; the sandbox itself does not
    awk '/application-identifier|team-identifier/ { skip = 2 } skip-- > 0 { next } { print }' \
      "$TAURI/entitlements/Writer.appstore.entitlements" > "$TAURI/entitlements/.generated.entitlements"
    TARGET=$(rustc -vV | sed -n 's/host: //p')
    APP="$TAURI/target/$TARGET/release/bundle/macos/Writer.app"
    node scripts/prepare-engine.mjs --release # the same NativeAOT engine the store build ships
  fi
  # the store build has no private API; its entitlements file is the generated one above
  CONFIG=$TAURI/tauri.appstore.conf.json
  [ "$mode" = appstore ] || CONFIG='{"identifier":"cn.thewriter.sandboxtest","app":{"macOSPrivateApi":false},"bundle":{"macOS":{"entitlements":"./entitlements/.generated.entitlements"}}}'
  cp "$TAURI/Cargo.toml" "$OUT/.Cargo.toml.keep" # the CLI may rewrite tauri's features to match macOSPrivateApi
  BUILD='{"bundle":{"macOS":{"bundleVersion":"'"${BUILD_NUMBER:-$(git rev-list --count HEAD)}"'"}}}'
  status=0
  APPLE_SIGNING_IDENTITY="$SIGN_ID" npx tauri build --bundles app --target "$TARGET" --features "appstore${EXTRA_FEATURES:+,$EXTRA_FEATURES}" --config "$CONFIG" --config "$BUILD" || status=$?
  cp "$OUT/.Cargo.toml.keep" "$TAURI/Cargo.toml"; rm -f "$OUT/.Cargo.toml.keep"
  [ $status -eq 0 ] || exit $status
  sign_inner "$SIGN_ID" "$TAURI/entitlements/.generated.entitlements"
  codesign -d --entitlements - "$APP" 2>/dev/null | grep -q app-sandbox && echo "sandboxed: $APP"
  if [ "$mode" = appstore ]; then
    xcrun productbuild --sign "$INSTALLER_ID" --component "$APP" /Applications "$OUT/Writer-$VERSION.pkg"
    echo "上传（API 密钥见 README）：xcrun altool --upload-package $OUT/Writer-$VERSION.pkg -t macos --api-key <key id> --api-issuer <issuer id> --p8-file-path <AuthKey.p8>"
  fi
  ;;
*)
  sed -n '2,17p' "$0"; exit 1 ;;
esac
