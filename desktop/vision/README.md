# writer-vision

The engine's picture helper: a small native program next to the `writer` binary that does what .NET cannot do alone —
cut the subject out of a photo, and decode, scale and re-encode pictures. `main.swift` is the macOS build: Apple Vision
(`VNGenerateForegroundInstanceMaskRequest`, the framework behind Finder's 移除背景, macOS 14 or later) and ImageIO.
A Windows build (`writer-vision.exe`) keeps the same command line.

## Where the engine finds it

1. next to its own executable: `writer-vision`, or `writer-vision.exe` on Windows;
2. the path in the `WRITER_VISION` environment variable;
3. on `PATH`.

Without it, `background=remove` and `compress` fail with a clear error (抠图需要 macOS 14 以上的 Apple Vision …) and the
picture stays as it was. The desktop app ships it as a second sidecar (`src-tauri/tauri.macos.conf.json`).

## Command line

```
writer-vision cutout <in> <out.png> [--at x,y]
writer-vision compress <in> <out.png|out.jpg> --max WxH [--crop l,t,r,b] [--quality q]
```

`<in>` is any picture the platform can read (the engine passes PNG, JPEG, GIF or BMP). Both commands first apply the
picture's EXIF orientation, so the output is upright and carries no orientation tag.

**cutout** writes an RGBA PNG of the same pixel size in which everything but the subject is fully transparent.
- By default every foreground instance is kept (a person and the dog next to them).
- `--at x,y` keeps only the instance under that point; x and y run from 0 to 1, from the top-left corner.

**compress** re-encodes for a smaller file:
- `--crop l,t,r,b` first cuts those fractions (0 to 1) off the left, top, right and bottom edges;
- then the picture is scaled down, never up, to fit within `--max` W×H pixels, keeping its aspect ratio;
- the output format follows the extension of `<out>`: `.jpg`/`.jpeg` is JPEG at `--quality` (0 to 1, default 0.85),
  anything else PNG with its alpha.

## Exit codes

| code | meaning |
|---|---|
| 0 | done, `<out>` written |
| 1 | the picture could not be read, processed or written |
| 2 | cutout found no subject (anywhere, or at `--at`) |
| 64 | usage error |

On failure a one-line message goes to stderr and `<out>` is not written; the engine shows the message and leaves the
picture untouched.

## Build

`node desktop/scripts/build-vision.mjs [--universal]` compiles it with `xcrun swiftc` for macOS 14 into
`desktop/src-tauri/binaries/writer-vision-<target triple>` (both architectures joined by `lipo` with `--universal`).
On other systems the script does nothing. The engine's tests find it there and run the Vision tests; without it they
are skipped.
