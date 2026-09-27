#!/usr/bin/env bash
# Publishes self-contained single-file binaries to dist/<rid>/writer.
# Usage: ./build.sh            (all platforms)
#        ./build.sh osx-arm64  (one platform)
#        ./build.sh embed      (Writer for other sites' pages: dist/embed/writer, static files, see docs/embed.md)
set -euo pipefail
cd "$(dirname "$0")"

if [ "${1:-}" = "embed" ]; then
  # the editor UI and the engine compiled to WebAssembly: a folder any static host serves (.wasm as application/wasm; the
  # .br / .gz beside each file are precompressed copies for a host that serves them)
  out=dist/embed/writer
  rm -rf dist/embed && mkdir -p "$out"
  dotnet publish src/Writer.Browser/Writer.Browser.csproj -c Release -o dist/embed/.engine --nologo -v quiet
  cp -R ui/. "$out/" && rm -rf "$out/tests" "$out/tools" "$out/_framework"
  cp -R dist/embed/.engine/wwwroot/_framework "$out/_framework"
  rm -rf dist/embed/.engine
  echo "$out: pages load $out/embed/writer-embed.js; the example is embed/example.html"
  exit 0
fi

rids=("$@")
if [ ${#rids[@]} -eq 0 ]; then
  rids=(osx-arm64 osx-x64 linux-x64 linux-arm64 win-x64)
fi

for rid in "${rids[@]}"; do
  echo "== $rid"
  # ReadyToRun: compiled ahead, so a command does not JIT the engine first (a view 0.9 s → 0.4 s; twice the size)
  dotnet publish src/Writer.Cli/Writer.Cli.csproj -c Release -r "$rid" --self-contained \
    -p:PublishSingleFile=true -p:PublishTrimmed=true -p:EnableCompressionInSingleFile=true -p:PublishReadyToRun=true \
    -p:IncludeNativeLibrariesForSelfExtract=true -p:DebugType=none \
    -o "dist/$rid" --nologo -v quiet
  rm -rf "dist/$rid/ui" && cp -R ui "dist/$rid/ui"   # writer app serves the UI from next to the binary
  ls -la "dist/$rid"
done
