# Writer desktop

`desktop/` is Writer for Mac: Tauri 2 windows around the editor in `ui/`, each served by its own bundled, self-contained .NET `writer` engine (a sidecar). Documents are edited in place, wherever they are.

## How it works

1. **One window per folder.** A plain launch opens the drafts folder (`<app data>/Drafts`, created when missing): new documents live there as drafts until their first 存储, which opens the system Save dialog and grants the chosen file to that window's engine over `--grants-stdin` (另存为… asks again, leaving the original as it was). Opening a file — Finder double-click, a drop on the Dock icon, `open -a Writer file`, 文件 › 打开… (⌘O), 打开最近使用, the start page's own recent files, or a drop on a window — focuses the window for that file's folder, or opens one, and opens the file's tab there. Every window runs `writer app --no-browser --list-depth <n> --dir <folder> --ui <resources>/ui --drafts <app data>/Drafts --grants-stdin` (depth 3 for the drafts folder, 0 for a file's folder; `src-tauri/src/main.rs`). Closing a window stops its engine; closing the last one leaves Writer in the Dock, ⌘Q quits.
2. `writer app --no-browser` prints `{"url": "http://127.0.0.1:<port>/app/index.dc.html?auth=<one-time code>"}` on stderr; the window goes from `splash/index.html` to `/app/mac.dc.html?native=1&auth=…` on that port. The first visit swaps the code for an HttpOnly cookie named after the port. The drafts folder's engine keeps its port between launches (`engine.json` in the app-data folder).
3. `ui/mac.dc.html` is the design's Mac window: the 38px title strip (drawn traffic lights, sidebar, undo/redo, title or tabs, toolbar collapse, ✦ AI, ▶) over the shell. The native traffic lights are hidden on every window (the design's are smaller and closer together); the page's lights call the window plugin.
4. **Menu bar** (design: `WriterMac` `menuDefs`): predefined items for 撤销/重做/剪切/拷贝/粘贴/全选, 服务, 隐藏, 退出, 最小化, 缩放, 进入全屏幕; 窗口 and 帮助 are NSApp's window and help menus (macOS adds the window list and the search field). Other items call `window.__writerMenu(id)` in the front window; the page reports its state (`shell_state`) for the window title and the check marks and titles (显示边栏, 收起/展开工具栏, 深色/浅色模式).
5. **设置 / 关于 / 触控板手势 / 键盘快捷键** open `MacSettings` / `MacAbout` / `MacGestures` from the front window's engine as separate transparent windows with macOS vibrancy; `src-tauri/src/native.js` makes the page backgrounds let the glass through.
6. **Shared settings.** Each engine is its own origin, so localStorage is per window and per launch. `native.js`, an initialization script in every window, seeds `writer-settings` and `writer-mac` from `web-storage.json` in the app-data folder and sends every change to Rust (`settings_set`), which saves it, sets the app's light/dark appearance and hands it to the other windows.
7. `src-tauri/capabilities/writer-mac-shell.json` lets the pages from `http://127.0.0.1:*` (Tauri matches the IPC caller's `Origin`, which has no path) call Writer's own commands and the window-plugin calls behind the drawn traffic lights, nothing else.
8. Document types (`src-tauri/Info.plist`): .docx .xlsx .pptx .md/.markdown as Editor, .mm as Editor by extension with rank Alternate (it is also Objective-C++ source), .pdf as Viewer (Alternate), each with its B-set icon.
9. **Updates** (`src-tauri/src/update.rs`, tauri-plugin-updater), without asking: 10 s after launch and then once a day, unless 设置 › 通用 › 自动更新 is off, Writer reads `https://thewriter.cn/updates/mac/latest.json`, and `latest.json` from the newest GitHub release when that fails (GitHub downloads are slow in mainland China; `plugins.updater.endpoints` in `tauri.macos.conf.json` / `tauri.windows.conf.json`). A connection or TLS error or a non-2xx status moves on to GitHub; a 2xx reply that is not JSON ends the check with an error, and when the website serves a manifest, GitHub is not asked. A newer version downloads in the background and its signature is checked against `plugins.updater.pubkey`; nothing is shown, and the app bundle is replaced only when Writer quits, after the engines have stopped (a bundle replaced mid-run would pair this shell with a newer engine and UI), so the next launch runs it. This silent path runs only when the user can write the bundle and its folder (no admin prompt); otherwise only Writer › 检查更新… installs. 检查更新… ends in one dialog: 已是最新版本, or 新版本已下载 with 立即重启 (installs, then restarts) / 好. A failed manual check says so; an automatic one never shows anything.
   **Interface updates**: when there is no new version, Writer reads `ui.json` (`plugins.updater.uiEndpoints`: the website's `updates/ui/ui.json`, then GitHub's): a signed `ui/` for this exact version (`base`) with a higher `build`, made by `scripts/release-ui.sh`. It is verified (the signature's version must be `<base>-ui.<build>`), unpacked to `<app data>/ui/<base>-<build>/` and used from the next launch for every window and panel; About then shows the build, e.g. 0.1.2 (3). A package whose window does not report in within 20 s is marked bad and Writer restarts on the bundled `ui/`; packages for another version are deleted.

## Windows

`ui/win.dc.html` is the design's Windows window (WriterWin) on the same shell hooks as the Mac page; `page_url` picks it on Windows.

1. **Frameless window** (`src-tauri/tauri.windows.conf.json`, merged over `tauri.conf.json` on Windows): `decorations: false` (Windows 11 keeps the shadow and rounded corners), minimum 760 × 480, the `writer-win-shell` capability, an NSIS per-user installer and `icons/windows/icon.ico` (the dark logo tile, the design's Windows icon).
2. **The 40px title bar**: the W app menu (F1; 新建 · 打开 · 存储 · 另存为… · 打印 · 深色模式 · 收起工具栏 · 设置 · 键盘快捷键 · 帮助 · 关于 · 检查更新… · 退出, with a search box), sidebar, undo/redo, the title or tabs, the toolbar collapse, ✦ AI, ▶, and the caption buttons. The bar is a `data-tauri-drag-region="deep"`: it drags the window, a double-click maximises, Aero Snap works; buttons and tabs (`role="tab"`) stay clickable.
3. **Snap layouts**: hovering maximise shows the design's four layouts; a zone calls `snap_window` (`src-tauri/src/win.rs`), which fills that part of the monitor's work area. Layouts whose zones are smaller than the minimum window size are left out, as Windows does.
4. **No menu bar** (`launch` builds the native menu on macOS only). The shell already handles Ctrl+N/T/S/P/J/./,/Shift+L; the page adds Ctrl+O (the system open dialog through `open_files`, files open in place), Ctrl+W, Ctrl+Tab, Ctrl+\\, Ctrl+Alt+T, F1 and Alt+F4, and swallows F5 / Ctrl+R so a stray reload does not reset the window.
5. 设置 and 关于 are in-page panels in the design's Windows style (the `win` prop of `MacSettings` / `MacAbout`). Closing the last window quits Writer. Explorer hands files over as command-line arguments; `canonicalize()`'s `\\?\` prefix is stripped (`win::plain_path`) before paths reach the engine.
6. **Updates**: as on the Mac (item 9 above, with the website's `updates/windows/latest.json`; W › 检查更新… calls `check_update`), from the `windows-x86_64` entry of `latest.json`. The installer cannot replace a running Writer, so a new version is only downloaded and verified in the background; when Writer quits (`RunEvent::Exit`, after the engines stop, together with those of windows closed just before) the NSIS installer runs quietly (`plugins.updater.windows.installMode`) and does not reopen Writer. 立即重启 after a manual check installs it and opens Writer again. Interface packages work as on the Mac.

## Build

Prerequisites: .NET 10 SDK, a Rust toolchain, Node 20 and the platform build tools (Xcode command line tools on macOS, the Visual Studio C++ build tools and WebView2 on Windows). Build on the target OS; on Windows `npm run build` writes the NSIS installer to `src-tauri/target/release/bundle/nsis/`.

```bash
cd desktop
npm ci
npm run build        # = npm run prepare:engine && tauri build
```

`npm run prepare:engine` (`scripts/prepare-engine.mjs`) publishes `src/Writer.Cli` as a single self-contained binary for the current platform into `.engine/<rid>/` and copies it to `src-tauri/binaries/writer-<target triple>` (Tauri's sidecar naming). `tauri build` then bundles that binary and `../../ui/` as resources.

On macOS the app carries a second sidecar, `writer-vision` (`vision/`, see its README): Apple Vision background removal (图片 › 抠图) and ImageIO re-encoding (压缩图片) for the engine. `scripts/build-vision.mjs` compiles it with `xcrun swiftc` into `src-tauri/binaries/writer-vision-<target triple>` and does nothing on other systems; `tauri.macos.conf.json` lists it next to the engine, so only Mac builds need it.

Only the macOS `.app` (no dmg):

```bash
npm run prepare:engine && npx tauri build --bundles app
open src-tauri/target/release/bundle/macos/Writer.app
```

Development window (same engine preparation, then `tauri dev`; in debug builds the shell falls back to the repository's `ui/` folder when no bundled UI exists):

```bash
npm run dev
```

## Signing and release

Everything needed to sign is already on this Mac; nothing secret lives in the repo.

| What | Where | Used for |
|---|---|---|
| `Developer ID Application: <name> (<team>)` | login keychain (certificate + private key) | website build: app and dmg |
| `Apple Distribution: <name> (<team>)` | login keychain | Mac App Store build: the app |
| `3rd Party Mac Developer Installer: <name> (<team>)` | login keychain | Mac App Store build: the .pkg |
| `Apple Development: <name> (<id>)` | login keychain | local sandbox tests. A revoked certificate can keep the same name; an app signed with it is killed at launch and removed by macOS, so pick identities by SHA-1 (`release-mac.sh` does) |
| notarization credentials | keychain profile `writer-notary` | `xcrun notarytool submit … --keychain-profile writer-notary` |
| Mac App Store provisioning profile | `../.signing/Writer_Mac_App_Store.provisionprofile` (gitignored) | embedded in the store build |
| updater signing key (`tauri signer`) | `../.signing/writer-updater.key` (gitignored), its password in the keychain item `writer-updater` | signing the updater archive; its public key is `plugins.updater.pubkey` in `tauri.conf.json`. Losing it means installed copies can never update themselves again |
| App Store Connect API key (team, admin) | outside the repo; export `ASC_KEY_ID`, `ASC_ISSUER_ID`, `ASC_KEY_FILE` | `scripts/asc.mjs`, `xcrun altool` |

The names, team, App Store Connect app and key location of this release Mac are in `desktop/SIGNING.local.md` (gitignored). Bundle id `cn.thewriter.app`. Releases reach the website through `website/deploy.sh` (steps below; the website session owns `website/`). Release engines must come from Microsoft's .NET SDK (`~/.dotnet`): Homebrew's .NET links `/opt/homebrew` brotli, so its engine does not start on other Macs or in the sandbox; `prepare-engine.mjs --release` refuses it.

```bash
cd desktop
NOTES_FILE=notes.txt scripts/release-mac.sh web   # Developer ID: signed, notarized, stapled app → dist/Writer-<version>-mac.dmg (+ .sha256),
                                                  # Writer-<version>-mac.app.tar.gz (+ .sig) and latest.json for the updater,
                                                  # and the website's copy of the updater files in dist/updates/mac/
scripts/release-mac.sh appstore       # sandboxed store build → dist/Writer-<version>.pkg (build number = commit count)
scripts/release-mac.sh sandbox-test   # the store sandbox signed with Apple Development, as cn.thewriter.sandboxtest
xcrun altool --upload-package dist/Writer-<version>.pkg -t macos --api-key <id> --api-issuer <issuer> --p8-file-path <AuthKey.p8>
TARGET=aarch64-apple-darwin scripts/release-mac.sh web   # Apple silicon only (universal needs `rustup target add x86_64-apple-darwin`)
NOTES_FILE=notes.txt scripts/release-win.sh dist/win/Writer_<version>_x64-setup.exe   # the installer from the Windows workflow
                                                  # → dist/Writer-<version>-windows-x64-setup.exe (+ .sig), signed with the updater
                                                  # key; adds windows-x86_64 to latest.json and writes dist/updates/windows/
NOTES_FILE=notes.txt scripts/release-ui.sh        # interface only → dist/Writer-ui-<version>-<build>.tar.gz (+ .sig), dist/ui.json
                                                  # for GitHub and dist/updates/ui/ for the website
```

**Two kinds of release.** A **full release** ships a new version of the app. It is needed when anything but `ui/`, `website/`, `docs/`, `tests/` and Markdown files changed since the last full release: the engine (`src/`), `desktop/`, build files. When only `ui/` changed, an **interface-only release** is enough: copies of the last version download the interface alone (about 1 MB) instead of a new app. `scripts/release-ui.sh` checks which one applies and refuses when a full release is needed.

**Releasing a version** (the full release; Mac and Windows from the same commit; the version is `version` in `src-tauri/tauri.conf.json`). The installers go to GitHub Releases and to the website, and installed copies look for updates on the website first (item 9 above). From `desktop/`:

1. **Windows installer**: pushing the tag `v<version>`, or `gh workflow run windows.yml -R Auspexlabs/writer`, runs `.github/workflows/windows.yml` on a GitHub Windows runner: `npm ci`, `prepare-engine.mjs --release`, `tauri build`. It holds no secrets and signs nothing; the NSIS installer is the run's artifact `writer-windows-x64-setup`.
2. **Mac**: `NOTES_FILE=notes.txt scripts/release-mac.sh web`.
3. **Windows files**: `gh run download <run id> -R Auspexlabs/writer -n writer-windows-x64-setup -D dist/win` (run ids: `gh run list -R Auspexlabs/writer -w windows.yml`), then `NOTES_FILE=notes.txt scripts/release-win.sh dist/win/Writer_<version>_x64-setup.exe`. The updater signature is made here: the key never leaves this Mac.
4. **Website**: `../website/deploy.sh --dmg dist/Writer-<version>-mac.dmg --exe dist/Writer-<version>-windows-x64-setup.exe --updates dist/updates` — the installers behind `/download/mac` and `/download/windows`, the updater files in `/updates/mac/` and `/updates/windows/`.
5. **GitHub**: `cp dist/Writer-<version>-mac.dmg dist/Writer-mac.dmg`, then `gh release create v<version> -R Auspexlabs/writer --notes-file notes.txt dist/Writer-<version>-mac.dmg dist/Writer-mac.dmg dist/Writer-<version>-mac.app.tar.gz dist/Writer-<version>-windows-x64-setup.exe dist/latest.json` (the versioned dmg serves the Homebrew cask, the unversioned one `releases/latest/download/Writer-mac.dmg` in SKILL.md).
6. **Tag** the commit the release was built from: `git tag v<version>` (name the commit if HEAD has moved since). The tag stays local (the repo has no remote); `release-ui.sh` compares against it.

**Releasing the interface only** (on the commit to ship, `ui/` committed; the version stays). From `desktop/`:

1. `NOTES_FILE=notes.txt scripts/release-ui.sh`. The base is `version` in `src-tauri/tauri.conf.json` (the last full release); the build is the next number for that base, counted by the local tags `v<base>-ui.<build>` that the script adds. It refuses when anything but `ui/`, `website/`, docs and tests changed since the tag `v<base>`: make a full release instead.
2. **GitHub**: `gh release upload v<base> -R Auspexlabs/writer dist/ui.json dist/Writer-ui-<base>-<build>.tar.gz --clobber`, onto the full release `v<base>` itself, never a new release: `releases/latest` must keep `latest.json`, `Writer-mac.dmg` and `ui.json` together.
3. **Website**: `../website/deploy.sh --updates dist/updates` (`dist/updates/ui/` goes to `/updates/ui/`, `ui.json` last).

Installed copies read `https://thewriter.cn/updates/ui/ui.json`, then `releases/latest/download/ui.json` on GitHub, so this also goes to both places. `ui.json` holds `base`, `build`, `notes`, `pub_date`, `url` (that host's copy of the archive) and `signature`. The archive is HEAD's `ui/` without `tests/` and `tools/`, under one top-level `ui/` folder (the app bundles the whole folder; those two are not used at runtime), signed with the updater key for the version `<base>-ui.<build>`. Only copies whose version is exactly `base` take it. A full release starts a new base: its copies ignore the `ui.json` of an older base, which stays on the website until the next interface-only release, and the new GitHub release has none.

Every release goes to both places: a copy that reaches the website takes the website's `latest.json` as the answer, so a website left one version behind hides the newer GitHub release. `dist/latest.json` is the GitHub feed, read as `releases/latest/download/latest.json`, so the newest release must carry it; it points to the archive and the installer under `releases/download/v<version>/`. `dist/updates/mac/` and `dist/updates/windows/` are the website's feeds: the same signed files, each with a `latest.json` that points to `https://thewriter.cn/updates/<mac|windows>/`. Their `notes` come from NOTES or NOTES_FILE; the app shows no release notes, since updates install without asking. `release-mac.sh web` and `release-win.sh` each add their platform (`scripts/latest-json.mjs`), in either order; a `latest.json` of another version starts over. `release-win.sh` refuses an installer whose file name has another version. The archive is made from the stapled app, and every signature is bound to the version (`requireSignedVersion`), so a `latest.json` cannot pair a version number with another release's files.

**In the store sandbox** the app may touch only what the user gave it, and the engines share that access (they inherit the app's sandbox, and access the app gains later reaches an engine already running). A file opened from Finder, 打开… or 最近使用 is sent to its window's engine as `list <path>` on the grants channel, so the listing shows it although the folder cannot be read. A Save dialog target is sent as `allow <path>`. Recent files keep security-scoped bookmarks (`<app data>/bookmarks.json`), resolved at launch. `EXTRA_FEATURES=test-hooks scripts/release-mac.sh sandbox-test` builds the sandboxed app with test hooks, and a plan passed inline (`open --env 'WRITER_TEST_PLAN=[…]'`; a sandboxed app cannot read a plan file) drives it.

Local test builds from the main checkout overwrite `src-tauri/binaries/` and `src-tauri/target/`: build and run your own copies from a worktree instead of launching those files while a release is building. A sandboxed window that stays white is usually just occluded (WebKit pauses hidden pages): bring it to the front before judging it.

## Icons

`src-tauri/icons/` is generated from the brand mark:

```bash
npx tauri icon ../brand/writer-mark-1024.png -o src-tauri/icons
rm -rf src-tauri/icons/android src-tauri/icons/ios src-tauri/icons/Square*.png src-tauri/icons/StoreLogo.png
```

`src-tauri/icons/doc/*.icns` (Finder document icons) come from the B set, `ui/assets/icons/b/*.svg`: render each with `qlmanage -t -s 1024 -o <dir> <name>.svg`, scale `<name>.svg.png` with `sips -z` into a `<name>.iconset` (16…512 @1x/@2x) and run `iconutil -c icns <name>.iconset`.

## Diagnostics

`ui/perf.js` is a frame-time HUD (fps, worst frame, dropped frames, long tasks; per-second samples in `window.__perfLog`). In a browser open the editor with `?perf=1` or set `localStorage.writerPerf = 1`. In the release app pass `WRITER_PERF`:

```bash
# HUD only
open -n src-tauri/target/release/bundle/macos/Writer.app --env WRITER_PERF=1
# scripted interactions (panels, ribbon bubble, typing at 100%/200%, scrolling, slide drag, assistant input);
# results land in <workspace>/perf/*.json. A throwaway HOME keeps the real documents out of it.
open -n src-tauri/target/release/bundle/macos/Writer.app --env WRITER_PERF=auto --env HOME=/tmp/writer-perf-home
```

`WRITER_PERF=auto` opens `long.docx`, `Mars-Settlement-Guide.pptx` and `big.xlsx` from the workspace when they exist, otherwise the first document of each type, and edits them (typing, a slide drag that it undoes), so point it at copies.

UI checks without Accessibility access: build with `npx tauri build --bundles app --features test-hooks --config '{"identifier":"dev.writer.desktop.test"}'` (its own app-data folder) and start it with `open -g --env WRITER_TEST_PLAN=plan.json`; the plan is a list of timed steps — `{"at": 9, "menu": "settings"}`, `{"at": 10, "eval": ["main-1", "js"]}`, `{"at": 11, "menus": true}` (prints the menu bar), `{"at": 12, "top": true}` (floats the windows for `screencapture -l`), `{"at": 20, "exit": true}`. Update tests: `--env WRITER_TEST_UPDATE=1` answers the updater's dialogs with their first button (and prints them), and `--config '{"plugins":{"updater":{"endpoints":["http://127.0.0.1:<port>/latest.json"],"dangerousInsecureTransportProtocol":true}}}'` points the build at a local `latest.json`.

## Not done yet

- Mac App Store build: the Save dialog path in the sandbox is only checked by hand (the panel cannot be scripted); Finder-open, writes, the refusal of other files and 最近使用 after a relaunch are checked with the sandboxed test build.
- App builds do not ship the third-party license notices (.NET runtime, NuGet packages, Rust crates) yet; see THIRD_PARTY_NOTICES.md.
- A Windows build has not been run on Windows.
- Windows: opening a file while Writer runs starts a second Writer process instead of handing the file to the first (needs a single-instance plugin).
- A file that is not in its window's listing (older than the folder's 500 newest documents, or created after the window opened) cannot be opened in an existing window yet: the shell has no "open by path".
