#!/usr/bin/env bash
# Publishes self-contained single-file binaries to dist/<rid>/writer.
# Usage: ./build.sh            (all platforms)
#        ./build.sh osx-arm64  (one platform)
set -euo pipefail
cd "$(dirname "$0")"

rids=("$@")
if [ ${#rids[@]} -eq 0 ]; then
  rids=(osx-arm64 osx-x64 linux-x64 linux-arm64 win-x64)
fi

for rid in "${rids[@]}"; do
  echo "== $rid"
  dotnet publish src/Writer.Cli/Writer.Cli.csproj -c Release -r "$rid" --self-contained \
    -p:PublishSingleFile=true -p:PublishTrimmed=true -p:EnableCompressionInSingleFile=true \
    -p:IncludeNativeLibrariesForSelfExtract=true -p:DebugType=none \
    -o "dist/$rid" --nologo -v quiet
  rm -rf "dist/$rid/ui" && cp -R ui "dist/$rid/ui"   # writer app serves the UI from next to the binary
  ls -la "dist/$rid"
done
