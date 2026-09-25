# Third-party notices

This file lists the third-party components that this repository contains, that Writer builds include, or that
Writer loads at run time, with their licenses. Writer's own license is in `LICENSE`.

Versions are the ones pinned in this repository (`*.csproj`, `desktop/package-lock.json`,
`desktop/src-tauri/Cargo.lock`, the URLs in `ui/`). The texts of the licenses are at the end of this file or linked
there.

## 1. Writer engine (`src/`, bundled as the `writer` sidecar)

| Component | Version | License | Copyright | Source |
|---|---|---|---|---|
| DocumentFormat.OpenXml, DocumentFormat.OpenXml.Framework | 3.5.1 | MIT | © Microsoft Corporation | https://github.com/dotnet/Open-XML-SDK |
| Markdig | 1.4.0 | BSD-2-Clause | Copyright (c) Alexandre Mutel | https://github.com/xoofx/markdig |
| PdfPig (UglyToad.PdfPig.*) | 0.1.16 | Apache-2.0 | PdfPig contributors (package author: UglyToad) | https://github.com/UglyToad/PdfPig |
| ModelContextProtocol.Core | 2.2.0 | Apache-2.0 | © Model Context Protocol a Series of LF Projects, LLC. | https://github.com/modelcontextprotocol/csharp-sdk |
| Microsoft.Extensions.AI.Abstractions | 10.8.3 | MIT | © Microsoft Corporation | https://github.com/dotnet/extensions |
| Microsoft.Extensions.DependencyInjection.Abstractions | 10.0.10 | MIT | © Microsoft Corporation | https://github.com/dotnet/runtime |
| Microsoft.Extensions.Logging.Abstractions | 10.0.10 | MIT | © Microsoft Corporation | https://github.com/dotnet/runtime |
| System.IO.Packaging | 10.0.2 | MIT | © Microsoft Corporation | https://github.com/dotnet/runtime |
| .NET runtime (self-contained / NativeAOT publish) | 10.0 | MIT | Copyright (c) .NET Foundation and Contributors | https://github.com/dotnet/runtime |

PdfPig embeds Adobe data files:

- Adobe CMap resources — Copyright 1990-2009 Adobe Systems Incorporated. All rights reserved. BSD-3-Clause.
- Adobe Glyph List 2.0 and ITC Zapf Dingbats Glyph List — Copyright 1997, 1998, 2002, 2007, 2010 Adobe Systems
  Incorporated. BSD-3-Clause.
- Core 14 AFM font metrics — Copyright (c) 1985, 1987, 1989, 1990, 1991, 1992, 1993, 1997 Adobe Systems Incorporated.
  All Rights Reserved. Under Adobe's AFM notice (quoted in "License texts"). Helvetica and Times are trademarks of
  Linotype-Hell AG and/or its subsidiaries.

The .NET runtime carries its own third-party notices (zlib-ng, Brotli, Unicode data, LLVM, mimalloc and others):
`THIRD-PARTY-NOTICES.TXT` in the runtime package, also at https://github.com/dotnet/runtime/blob/main/THIRD-PARTY-NOTICES.TXT.

Used only to build and run the tests (not part of any Writer build):

| Component | Version | License | Copyright |
|---|---|---|---|
| xunit, xunit.core, xunit.assert, xunit.extensibility.core, xunit.extensibility.execution | 2.9.3 | Apache-2.0 | Copyright (C) .NET Foundation |
| xunit.abstractions | 2.0.3 | Apache-2.0 | Copyright (C) .NET Foundation |
| xunit.analyzers | 1.18.0 | Apache-2.0 | Copyright (C) .NET Foundation |
| xunit.runner.visualstudio | 3.1.4 | Apache-2.0 | Copyright (C) .NET Foundation |
| Microsoft.NET.Test.Sdk, Microsoft.TestPlatform.ObjectModel, Microsoft.TestPlatform.TestHost, Microsoft.CodeCoverage | 17.14.1 | MIT | © Microsoft Corporation |
| Newtonsoft.Json | 13.0.3 | MIT | Copyright © James Newton-King 2008 |

## 2. Editor UI (`ui/`)

Bundled:

| Component | Version | License | Copyright | Where |
|---|---|---|---|---|
| React (`react.production.min.js`) | 18.3.1 | MIT | Copyright (c) Facebook, Inc. and its affiliates. | `ui/vendor/`, unmodified (matches the SRI hashes in `ui/support.js`) |
| ReactDOM (`react-dom.production.min.js`, includes a Modernizr 3.0.0pre custom build, MIT) | 18.3.1 | MIT | Copyright (c) Facebook, Inc. and its affiliates. | `ui/vendor/`, unmodified |
| KaTeX (`katex.min.js`, `katex.min.css`, `fonts/*.woff2`) | 0.18.9 | MIT | Copyright (c) 2013-2020 Khan Academy and other contributors | `ui/vendor/katex/`, unmodified — renders `$…$` / `$$…$$` math in `md.js` and the Markdown editor, no network request |
| highlight.js (`highlight.min.js`, the ES build of the "common" bundle; `github-dark.min.css`) | 11.11.1 | BSD-3-Clause | Copyright (c) 2006, Ivan Sagalaev | `ui/vendor/highlight/`, unmodified, license in `ui/vendor/highlight/LICENSE` — colours the Markdown editor's code blocks, no network request |
| IBM Plex Sans 400/500/600, IBM Plex Mono 400 (Latin subsets, via @fontsource 5.3.0) | — | OFL-1.1 | Copyright 2017, 2019 IBM Corp. | `ui/assets/fonts/`, license in `ui/assets/fonts/OFL.txt` |
| dc-runtime (`ui/support.js`) | — | Not stated in the file | — | The runtime of the design tool that exported the `ui/*.dc.html` pages. Generated (per its first line) from `dc-runtime/src/*.ts`; changed in this repository |

Loaded from public CDNs at run time (not stored in this repository):

| Component | Version | License | Copyright | Loaded by |
|---|---|---|---|---|
| PDF.js (`pdfjs-dist`), with its CMaps (Adobe, BSD-3-Clause) and standard fonts (Foxit fonts under Foxit's BSD-style license; Liberation Sans under OFL-1.1) | 4.4.168 | Apache-2.0 | Copyright Mozilla Foundation | `ui/pdf-kit.js` (cdn.jsdelivr.net) |
| pdf-lib, with @pdf-lib/standard-fonts (MIT), @pdf-lib/upng (MIT), pako (MIT AND Zlib), tslib (0BSD) | 1.17.1 | MIT | Copyright (c) 2019 Andrew Dillon | `ui/pdf-kit.js` (esm.sh) |
| JSZip, with pako (MIT AND Zlib), lie (MIT), setimmediate (MIT), readable-stream (MIT) | 3.10.1 | MIT OR GPL-3.0-or-later; used under MIT | Copyright (c) 2009-2016 Stuart Knightley, David Duponchel, Franz Buchinger, António Afonso | `ui/office-io.js` (esm.sh) |
| @babel/standalone (only when a page imports a JSX module) | 7.29.0 | MIT | Copyright (c) 2014-present Sebastian McKenzie and other contributors | `ui/support.js` (unpkg.com) |
| Mermaid (only when a Markdown document has a ```mermaid diagram) | 11.12.0 | MIT | Copyright (c) 2014 - 2022 Knut Sveidqvist | `ui/md.js` (cdn.jsdelivr.net) |
| React, ReactDOM (only when `ui/vendor/offline.js` is not loaded) | 18.3.1 | MIT | Copyright (c) Facebook, Inc. and its affiliates. | `ui/support.js` (unpkg.com) |

## 3. Desktop app (`desktop/`)

Rust crates compiled into the app (`desktop/src-tauri/Cargo.toml`, versions from `Cargo.lock`):

| Crate | Version | License | Copyright |
|---|---|---|---|
| tauri | 2.11.6 | Apache-2.0 OR MIT | Copyright (c) 2017 - Present Tauri Apps Contributors |
| tauri-build (build time) | 2.6.3 | Apache-2.0 OR MIT | Copyright (c) 2017 - Present Tauri Apps Contributors |
| tauri-plugin-dialog | 2.7.3 | Apache-2.0 OR MIT | Copyright (c) 2017 - Present Tauri Apps Contributors |
| tauri-plugin-shell | 2.3.6 | Apache-2.0 OR MIT | Copyright (c) 2017 - Present Tauri Apps Contributors |
| serde_json | 1.0.151 | MIT OR Apache-2.0 | Erick Tryzelaar, David Tolnay (authors) |
| sys-locale | 0.3.2 | MIT OR Apache-2.0 | 1Password (authors) |
| objc2-app-kit (macOS) | 0.3.2 | Zlib OR Apache-2.0 OR MIT | objc2 project, https://github.com/madsmtm/objc2 |
| objc2-foundation (macOS) | 0.3.2 | MIT | objc2 project, https://github.com/madsmtm/objc2 |

The macOS build compiles 269 crates in total (build scripts and macros included). Their licenses: MIT and/or
Apache-2.0 (most), Unicode-3.0 (the ICU4X crates and `unicode-ident`), MPL-2.0 (`cssparser`, `cssparser-macros` and
`selectors` from the Servo project, `dtoa-short`, `option-ext`; unmodified, sources on crates.io), Zlib,
BSD-3-Clause, Unlicense OR MIT, CC0-1.0 OR MIT-0 OR Apache-2.0, 0BSD OR MIT OR Apache-2.0. The other 191 crates in
`Cargo.lock` are for Windows, Linux, Android or wasm targets; all of them offer a permissive license (`r-efi` is
MIT OR Apache-2.0 OR LGPL-2.1-or-later and is used under MIT or Apache-2.0). The full list is in the table at the
end of this file. The copyright notices of each crate are in its source package on https://crates.io.

Build tool only (not part of the app): `@tauri-apps/cli` 2.11.5 and its platform packages, Apache-2.0 OR MIT,
Copyright (c) 2017 - Present Tauri Apps Contributors.

## 4. Website (`website/`)

| Component | License | Copyright | Where |
|---|---|---|---|
| Phosphor Icons (regular weight): apple-logo, arrow-right, caret-left, caret-right, cursor-click, folder-open, key, plugs-connected, shield-check, sparkle, tabs, terminal-window, windows-logo | MIT | Copyright (c) Phosphor Icons | SVG sprite in `website/dist/index.html` |

The Apple and Windows logos depict trademarks of Apple Inc. and Microsoft Corporation.
`website/lab/icon-lab.html` (an internal page, not published) loads IBM Plex Sans, Noto Sans SC and Noto Serif SC
from Google Fonts (OFL-1.1).

## 5. Test fixtures (`tests/Writer.Tests/Fixtures/`)

| Component | License | Copyright |
|---|---|---|
| OfficeCLI example documents: `docx/` (14), `pptx/` (12 of 13), `xlsx/` (10), `md/` (87) | Apache-2.0 | Copyright 2026 OfficeCLI (https://OfficeCLI.AI), created and maintained by goworm |

Copied from https://github.com/iOfficeAI/OfficeCLI `examples/` (commit `ffa8a0afbe2e9686abd636368e3da38c50f22131`)
byte for byte, with OfficeCLI's `LICENSE` and `NOTICE` as `LICENSE-OfficeCLI` and `NOTICE-OfficeCLI`.
`pptx/decor.pptx` is a modified copy of OfficeCLI's `Mars-Settlement-Guide.pptx`: the Writer project added a
decorated master and layouts, notes, a hidden slide and transitions to it (2026). The fixtures are used only by the
tests and are not part of any Writer build.

## License texts

- **Apache-2.0**: https://www.apache.org/licenses/LICENSE-2.0 (full text also in
  `tests/Writer.Tests/Fixtures/LICENSE-OfficeCLI`)
- **OFL-1.1**: `ui/assets/fonts/OFL.txt`, https://openfontlicense.org
- **MPL-2.0**: https://www.mozilla.org/MPL/2.0/
- **Unicode-3.0**: https://www.unicode.org/license.txt
- **Zlib**: https://spdx.org/licenses/Zlib.html
- **ISC**: https://spdx.org/licenses/ISC.html
- **0BSD**: https://spdx.org/licenses/0BSD.html
- **Unlicense**: https://spdx.org/licenses/Unlicense.html
- **CC0-1.0**: https://spdx.org/licenses/CC0-1.0.html
- **MIT-0**: https://spdx.org/licenses/MIT-0.html
- **Apache-2.0 WITH LLVM-exception**: https://spdx.org/licenses/LLVM-exception.html
- **LGPL-2.1-or-later** (offered by `r-efi`, not the license Writer uses): https://spdx.org/licenses/LGPL-2.1-or-later.html
- **GPL-3.0-or-later** (offered by JSZip, not the license Writer uses): https://spdx.org/licenses/GPL-3.0-or-later.html
- **Foxit fonts in PDF.js**: `standard_fonts/LICENSE_FOXIT` in the `pdfjs-dist` package

### MIT

```
Copyright (c) <year> <copyright holders>

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated
documentation files (the "Software"), to deal in the Software without restriction, including without limitation the
rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the
Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
```

### BSD-2-Clause (Markdig)

```
Copyright (c) Alexandre Mutel
All rights reserved.

Redistribution and use in source and binary forms, with or without modification, are permitted provided that the
following conditions are met:

1. Redistributions of source code must retain the above copyright notice, this list of conditions and the following
   disclaimer.
2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the
   following disclaimer in the documentation and/or other materials provided with the distribution.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES,
INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY,
WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
```

### BSD-3-Clause (Adobe CMap resources and glyph lists in PdfPig and PDF.js)

```
Copyright 1990-2009 Adobe Systems Incorporated.
Copyright 1997, 1998, 2002, 2007, 2010 Adobe Systems Incorporated.
All rights reserved.

Redistribution and use in source and binary forms, with or without modification, are permitted provided that the
following conditions are met:

Redistributions of source code must retain the above copyright notice, this list of conditions and the following
disclaimer.

Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following
disclaimer in the documentation and/or other materials provided with the distribution.

Neither the name of Adobe Systems Incorporated nor the names of its contributors may be used to endorse or promote
products derived from this software without specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES,
INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY,
WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
```

### Adobe Core 14 AFM files (in PdfPig)

```
This file and the 14 PostScript(R) AFM files it accompanies may be used, copied, and distributed for any purpose and
without charge, with or without modification, provided that all copyright notices are retained; that the AFM files
are not distributed without this file; that all modifications to this file or any of the AFM files are prominently
noted in the modified file(s); and that this paragraph is not modified. Adobe Systems has no responsibility or
obligation to support the use of the AFM files.
```

## Rust crates in `desktop/src-tauri/Cargo.lock`

"macOS build" marks the crates compiled when the macOS app is built (build scripts and macros included); "other
targets" are used only when building for Windows, Linux, Android or wasm.

<details>
<summary>459 crates</summary>

| Crate | Version | License | Used in |
|---|---|---|---|
| adler2 | 2.0.1 | 0BSD OR MIT OR Apache-2.0 | macOS build |
| aho-corasick | 1.1.5 | Unlicense OR MIT | macOS build |
| alloc-no-stdlib | 2.0.4 | BSD-3-Clause | macOS build |
| alloc-stdlib | 0.2.4 | BSD-3-Clause | macOS build |
| android_system_properties | 0.1.6 | MIT OR Apache-2.0 | other targets |
| anyhow | 1.0.104 | MIT OR Apache-2.0 | macOS build |
| atk | 0.18.2 | MIT | other targets |
| atk-sys | 0.18.2 | MIT | other targets |
| atomic-waker | 1.1.2 | Apache-2.0 OR MIT | other targets |
| autocfg | 1.5.1 | Apache-2.0 OR MIT | macOS build |
| base64 | 0.21.7 | MIT OR Apache-2.0 | macOS build |
| base64 | 0.22.1 | MIT OR Apache-2.0 | macOS build |
| base64 | 0.23.1 | MIT OR Apache-2.0 | macOS build |
| bit-set | 0.8.0 | Apache-2.0 OR MIT | macOS build |
| bit-vec | 0.8.0 | Apache-2.0 OR MIT | macOS build |
| bitflags | 1.3.2 | MIT/Apache-2.0 | macOS build |
| bitflags | 2.13.2 | MIT OR Apache-2.0 | macOS build |
| block-buffer | 0.10.4 | MIT OR Apache-2.0 | macOS build |
| block2 | 0.6.2 | MIT | macOS build |
| brotli | 8.0.4 | BSD-3-Clause AND MIT | macOS build |
| brotli-decompressor | 5.0.3 | BSD-3-Clause/MIT | macOS build |
| bs58 | 0.5.1 | MIT/Apache-2.0 | macOS build |
| bumpalo | 3.20.3 | MIT OR Apache-2.0 | other targets |
| bytemuck | 1.25.2 | Zlib OR Apache-2.0 OR MIT | other targets |
| byteorder | 1.5.0 | Unlicense OR MIT | macOS build |
| bytes | 1.12.1 | MIT | macOS build |
| cairo-rs | 0.18.5 | MIT | other targets |
| cairo-sys-rs | 0.18.2 | MIT | other targets |
| camino | 1.2.6 | MIT OR Apache-2.0 | macOS build |
| cargo-platform | 0.1.9 | MIT OR Apache-2.0 | macOS build |
| cargo_metadata | 0.19.2 | MIT | macOS build |
| cargo_toml | 0.22.3 | Apache-2.0 OR MIT | macOS build |
| cc | 1.4.7 | MIT OR Apache-2.0 | macOS build |
| cesu8 | 1.1.0 | Apache-2.0/MIT | other targets |
| cfb | 0.7.3 | MIT | macOS build |
| cfg-expr | 0.15.8 | MIT OR Apache-2.0 | other targets |
| cfg-if | 1.0.5 | MIT OR Apache-2.0 | macOS build |
| chrono | 0.4.45 | MIT OR Apache-2.0 | macOS build |
| combine | 4.6.8 | MIT | other targets |
| cookie | 0.18.2 | MIT OR Apache-2.0 | macOS build |
| core-foundation | 0.10.1 | MIT OR Apache-2.0 | macOS build |
| core-foundation-sys | 0.8.7 | MIT OR Apache-2.0 | macOS build |
| core-graphics | 0.25.0 | MIT OR Apache-2.0 | macOS build |
| core-graphics-types | 0.2.0 | MIT OR Apache-2.0 | macOS build |
| core_detect | 1.0.0 | MIT/Apache-2.0 | other targets |
| cpufeatures | 0.2.17 | MIT OR Apache-2.0 | macOS build |
| crc32fast | 1.5.2 | MIT OR Apache-2.0 | macOS build |
| crossbeam-channel | 0.5.17 | MIT OR Apache-2.0 | macOS build |
| crossbeam-utils | 0.8.23 | MIT OR Apache-2.0 | macOS build |
| crypto-common | 0.1.7 | MIT OR Apache-2.0 | macOS build |
| cssparser | 0.36.0 | MPL-2.0 | macOS build |
| cssparser-macros | 0.6.1 | MPL-2.0 | macOS build |
| ctor | 0.8.0 | Apache-2.0 OR MIT | macOS build |
| ctor-proc-macro | 0.0.7 | Apache-2.0 OR MIT | macOS build |
| darling | 0.24.1 | MIT | macOS build |
| darling_core | 0.24.1 | MIT | macOS build |
| darling_macro | 0.24.1 | MIT | macOS build |
| dbus | 0.9.12 | Apache-2.0/MIT | other targets |
| defmt | 1.1.1 | MIT OR Apache-2.0 | macOS build |
| defmt-macros | 1.1.1 | MIT OR Apache-2.0 | macOS build |
| defmt-parser | 1.0.0 | MIT OR Apache-2.0 | macOS build |
| deranged | 0.5.8 | MIT OR Apache-2.0 | macOS build |
| derive_more | 2.1.1 | MIT | macOS build |
| derive_more-impl | 2.1.1 | MIT | macOS build |
| digest | 0.10.7 | MIT OR Apache-2.0 | macOS build |
| dirs | 6.0.0 | MIT OR Apache-2.0 | macOS build |
| dirs-sys | 0.5.0 | MIT OR Apache-2.0 | macOS build |
| dispatch2 | 0.3.1 | Zlib OR Apache-2.0 OR MIT | macOS build |
| displaydoc | 0.2.7 | MIT OR Apache-2.0 | macOS build |
| dlopen2 | 0.8.2 | MIT | other targets |
| dlopen2_derive | 0.4.3 | MIT | other targets |
| dom_query | 0.27.0 | MIT | macOS build |
| dpi | 0.1.2 | Apache-2.0 AND MIT | macOS build |
| dtoa | 1.0.11 | MIT OR Apache-2.0 | macOS build |
| dtoa-short | 0.3.5 | MPL-2.0 | macOS build |
| dtor | 0.3.0 | Apache-2.0 OR MIT | macOS build |
| dtor-proc-macro | 0.0.6 | Apache-2.0 OR MIT | macOS build |
| dunce | 1.0.5 | CC0-1.0 OR MIT-0 OR Apache-2.0 | macOS build |
| dyn-clone | 1.0.20 | MIT OR Apache-2.0 | macOS build |
| embed-resource | 3.0.11 | MIT | macOS build |
| embed_plist | 1.2.2 | MIT OR Apache-2.0 | macOS build |
| encoding_rs | 0.8.41 | (Apache-2.0 OR MIT) AND BSD-3-Clause | macOS build |
| equivalent | 1.0.2 | Apache-2.0 OR MIT | macOS build |
| erased-serde | 0.4.10 | MIT OR Apache-2.0 | macOS build |
| errno | 0.3.14 | MIT OR Apache-2.0 | macOS build |
| fastrand | 2.5.0 | Apache-2.0 OR MIT | macOS build |
| fdeflate | 0.3.7 | MIT OR Apache-2.0 | macOS build |
| field-offset | 0.3.6 | MIT OR Apache-2.0 | other targets |
| find-msvc-tools | 0.1.13 | MIT OR Apache-2.0 | macOS build |
| flate2 | 1.1.10 | MIT OR Apache-2.0 | macOS build |
| fnv | 1.0.7 | Apache-2.0 / MIT | macOS build |
| foldhash | 0.2.0 | Zlib | macOS build |
| foreign-types | 0.5.0 | MIT/Apache-2.0 | macOS build |
| foreign-types-macros | 0.2.4 | MIT/Apache-2.0 | macOS build |
| foreign-types-shared | 0.3.1 | MIT/Apache-2.0 | macOS build |
| form_urlencoded | 1.2.2 | MIT OR Apache-2.0 | macOS build |
| futures-channel | 0.3.34 | MIT OR Apache-2.0 | other targets |
| futures-core | 0.3.34 | MIT OR Apache-2.0 | other targets |
| futures-executor | 0.3.34 | MIT OR Apache-2.0 | other targets |
| futures-io | 0.3.34 | MIT OR Apache-2.0 | other targets |
| futures-macro | 0.3.34 | MIT OR Apache-2.0 | other targets |
| futures-sink | 0.3.34 | MIT OR Apache-2.0 | other targets |
| futures-task | 0.3.34 | MIT OR Apache-2.0 | other targets |
| futures-util | 0.3.34 | MIT OR Apache-2.0 | other targets |
| gdk | 0.18.2 | MIT | other targets |
| gdk-pixbuf | 0.18.5 | MIT | other targets |
| gdk-pixbuf-sys | 0.18.0 | MIT | other targets |
| gdk-sys | 0.18.2 | MIT | other targets |
| gdkwayland-sys | 0.18.2 | MIT | other targets |
| gdkx11 | 0.18.2 | MIT | other targets |
| gdkx11-sys | 0.18.2 | MIT | other targets |
| generic-array | 0.14.7 | MIT | macOS build |
| getrandom | 0.3.4 | MIT OR Apache-2.0 | macOS build |
| getrandom | 0.4.3 | MIT OR Apache-2.0 | macOS build |
| gio | 0.18.4 | MIT | other targets |
| gio-sys | 0.18.1 | MIT | other targets |
| glib | 0.18.5 | MIT | other targets |
| glib-macros | 0.18.5 | MIT | other targets |
| glib-sys | 0.18.1 | MIT | other targets |
| glob | 0.3.4 | MIT OR Apache-2.0 | macOS build |
| gobject-sys | 0.18.0 | MIT | other targets |
| gtk | 0.18.2 | MIT | other targets |
| gtk-sys | 0.18.2 | MIT | other targets |
| gtk3-macros | 0.18.2 | MIT | other targets |
| hashbrown | 0.12.3 | MIT OR Apache-2.0 | macOS build |
| hashbrown | 0.17.1 | MIT OR Apache-2.0 | macOS build |
| heck | 0.4.1 | MIT OR Apache-2.0 | other targets |
| heck | 0.5.0 | MIT OR Apache-2.0 | macOS build |
| hex | 0.4.3 | MIT OR Apache-2.0 | macOS build |
| html5ever | 0.38.0 | MIT OR Apache-2.0 | macOS build |
| http | 1.5.0 | MIT OR Apache-2.0 | macOS build |
| http-body | 1.1.0 | MIT | other targets |
| http-body-util | 0.1.5 | MIT | other targets |
| httparse | 1.10.1 | MIT OR Apache-2.0 | other targets |
| hyper | 1.11.1 | MIT | other targets |
| hyper-util | 0.1.20 | MIT | other targets |
| iana-time-zone | 0.1.65 | MIT OR Apache-2.0 | macOS build |
| iana-time-zone-haiku | 0.1.2 | MIT OR Apache-2.0 | other targets |
| ico | 0.5.0 | MIT | macOS build |
| icu_collections | 2.3.0 | Unicode-3.0 | macOS build |
| icu_locale_core | 2.3.0 | Unicode-3.0 | macOS build |
| icu_normalizer | 2.3.0 | Unicode-3.0 | macOS build |
| icu_normalizer_data | 2.3.0 | Unicode-3.0 | macOS build |
| icu_properties | 2.3.0 | Unicode-3.0 | macOS build |
| icu_properties_data | 2.3.0 | Unicode-3.0 | macOS build |
| icu_provider | 2.3.1 | Unicode-3.0 | macOS build |
| ident_case | 1.0.1 | MIT/Apache-2.0 | macOS build |
| idna | 1.1.0 | MIT OR Apache-2.0 | macOS build |
| idna_adapter | 1.2.2 | Apache-2.0 OR MIT | macOS build |
| indexmap | 1.9.3 | Apache-2.0 OR MIT | macOS build |
| indexmap | 2.14.2 | Apache-2.0 OR MIT | macOS build |
| infer | 0.19.0 | MIT | macOS build |
| ipnet | 2.12.2 | MIT OR Apache-2.0 | other targets |
| is-docker | 0.2.0 | MIT | other targets |
| is-wsl | 0.4.0 | MIT | other targets |
| itoa | 1.0.18 | MIT OR Apache-2.0 | macOS build |
| javascriptcore-rs | 1.1.2 | MIT | other targets |
| javascriptcore-rs-sys | 1.1.1 | MIT | other targets |
| jiff | 0.2.37 | Unlicense OR MIT | macOS build |
| jiff-core | 0.1.1 | Unlicense OR MIT | macOS build |
| jiff-static | 0.2.37 | Unlicense OR MIT | other targets |
| jiff-tzdb | 0.1.8 | Unlicense OR MIT | other targets |
| jiff-tzdb-platform | 0.1.3 | Unlicense OR MIT | other targets |
| jni | 0.21.1 | MIT/Apache-2.0 | other targets |
| jni-sys | 0.3.1 | MIT OR Apache-2.0 | other targets |
| jni-sys | 0.4.1 | MIT OR Apache-2.0 | other targets |
| jni-sys-macros | 0.4.1 | MIT OR Apache-2.0 | other targets |
| js-sys | 0.3.105 | MIT OR Apache-2.0 | other targets |
| json-patch | 3.0.1 | MIT/Apache-2.0 | macOS build |
| jsonptr | 0.6.3 | MIT OR Apache-2.0 | macOS build |
| keyboard-types | 0.7.0 | MIT OR Apache-2.0 | macOS build |
| libappindicator | 0.9.0 | Apache-2.0 OR MIT | other targets |
| libappindicator-sys | 0.9.0 | Apache-2.0 OR MIT | other targets |
| libc | 0.2.189 | MIT OR Apache-2.0 | macOS build |
| libdbus-sys | 0.2.7 | Apache-2.0/MIT | other targets |
| libloading | 0.7.4 | ISC | other targets |
| libredox | 0.1.25 | MIT | other targets |
| litemap | 0.8.3 | Unicode-3.0 | macOS build |
| lock_api | 0.4.14 | MIT OR Apache-2.0 | macOS build |
| log | 0.4.34 | MIT OR Apache-2.0 | macOS build |
| markup5ever | 0.38.0 | MIT OR Apache-2.0 | macOS build |
| memchr | 2.8.3 | Unlicense OR MIT | macOS build |
| memoffset | 0.9.1 | MIT | other targets |
| mime | 0.3.17 | MIT OR Apache-2.0 | macOS build |
| miniz_oxide | 0.8.9 | MIT OR Zlib OR Apache-2.0 | macOS build |
| miniz_oxide | 0.9.1 | MIT OR Zlib OR Apache-2.0 | macOS build |
| mio | 1.2.3 | MIT | macOS build |
| muda | 0.19.3 | Apache-2.0 OR MIT | macOS build |
| multiversion | 0.9.0 | MIT OR Apache-2.0 | other targets |
| multiversion-macros | 0.9.0 | MIT OR Apache-2.0 | other targets |
| multiversion_no_op | 1.0.0 | Apache-2.0 OR MIT | macOS build |
| ndk | 0.9.0 | MIT OR Apache-2.0 | other targets |
| ndk-sys | 0.6.0+11769913 | MIT OR Apache-2.0 | other targets |
| new_debug_unreachable | 1.0.6 | MIT | macOS build |
| num-conv | 0.2.2 | MIT OR Apache-2.0 | macOS build |
| num-traits | 0.2.19 | MIT OR Apache-2.0 | macOS build |
| num_enum | 0.7.6 | BSD-3-Clause OR MIT OR Apache-2.0 | other targets |
| num_enum_derive | 0.7.6 | BSD-3-Clause OR MIT OR Apache-2.0 | other targets |
| objc2 | 0.6.4 | MIT | macOS build |
| objc2-app-kit | 0.3.2 | Zlib OR Apache-2.0 OR MIT | macOS build |
| objc2-cloud-kit | 0.3.2 | Zlib OR Apache-2.0 OR MIT | other targets |
| objc2-core-data | 0.3.2 | Zlib OR Apache-2.0 OR MIT | other targets |
| objc2-core-foundation | 0.3.2 | Zlib OR Apache-2.0 OR MIT | macOS build |
| objc2-core-graphics | 0.3.2 | Zlib OR Apache-2.0 OR MIT | macOS build |
| objc2-core-image | 0.3.2 | Zlib OR Apache-2.0 OR MIT | other targets |
| objc2-core-location | 0.3.2 | Zlib OR Apache-2.0 OR MIT | other targets |
| objc2-core-text | 0.3.2 | Zlib OR Apache-2.0 OR MIT | other targets |
| objc2-encode | 4.1.0 | MIT | macOS build |
| objc2-exception-helper | 0.1.1 | Zlib OR Apache-2.0 OR MIT | macOS build |
| objc2-foundation | 0.3.2 | MIT | macOS build |
| objc2-io-surface | 0.3.2 | Zlib OR Apache-2.0 OR MIT | macOS build |
| objc2-quartz-core | 0.3.2 | Zlib OR Apache-2.0 OR MIT | other targets |
| objc2-ui-kit | 0.3.2 | Zlib OR Apache-2.0 OR MIT | other targets |
| objc2-user-notifications | 0.3.2 | Zlib OR Apache-2.0 OR MIT | other targets |
| objc2-web-kit | 0.3.2 | Zlib OR Apache-2.0 OR MIT | macOS build |
| once_cell | 1.21.4 | MIT OR Apache-2.0 | macOS build |
| open | 5.4.4 | MIT | macOS build |
| option-ext | 0.2.0 | MPL-2.0 | macOS build |
| os_pipe | 1.2.3 | MIT | macOS build |
| pango | 0.18.3 | MIT | other targets |
| pango-sys | 0.18.0 | MIT | other targets |
| parking_lot | 0.12.5 | MIT OR Apache-2.0 | macOS build |
| parking_lot_core | 0.9.12 | MIT OR Apache-2.0 | macOS build |
| percent-encoding | 2.3.2 | MIT OR Apache-2.0 | macOS build |
| phf | 0.13.1 | MIT | macOS build |
| phf_codegen | 0.13.1 | MIT | macOS build |
| phf_generator | 0.13.1 | MIT | macOS build |
| phf_macros | 0.13.1 | MIT | macOS build |
| phf_shared | 0.13.1 | MIT | macOS build |
| pin-project-lite | 0.2.17 | Apache-2.0 OR MIT | macOS build |
| pkg-config | 0.3.34 | MIT OR Apache-2.0 | other targets |
| plist | 1.10.1 | MIT | macOS build |
| png | 0.17.16 | MIT OR Apache-2.0 | macOS build |
| png | 0.18.1 | MIT OR Apache-2.0 | macOS build |
| portable-atomic | 1.15.0 | Apache-2.0 OR MIT | other targets |
| portable-atomic-util | 0.2.8 | Apache-2.0 OR MIT | other targets |
| potential_utf | 0.1.6 | Unicode-3.0 | macOS build |
| powerfmt | 0.2.0 | MIT OR Apache-2.0 | macOS build |
| precomputed-hash | 0.1.1 | MIT | macOS build |
| proc-macro-crate | 1.3.1 | MIT OR Apache-2.0 | other targets |
| proc-macro-crate | 2.0.2 | MIT OR Apache-2.0 | other targets |
| proc-macro-crate | 3.5.0 | MIT OR Apache-2.0 | other targets |
| proc-macro-error | 1.0.4 | MIT OR Apache-2.0 | other targets |
| proc-macro-error-attr | 1.0.4 | MIT OR Apache-2.0 | other targets |
| proc-macro2 | 1.0.107 | MIT OR Apache-2.0 | macOS build |
| quick-xml | 0.42.0 | MIT | macOS build |
| quote | 1.0.47 | MIT OR Apache-2.0 | macOS build |
| r-efi | 5.3.0 | MIT OR Apache-2.0 OR LGPL-2.1-or-later | other targets |
| r-efi | 6.0.0 | MIT OR Apache-2.0 OR LGPL-2.1-or-later | other targets |
| raw-window-handle | 0.6.2 | MIT OR Apache-2.0 OR Zlib | macOS build |
| redox_syscall | 0.5.18 | MIT | other targets |
| redox_users | 0.5.3 | MIT | other targets |
| ref-cast | 1.0.27 | MIT OR Apache-2.0 | macOS build |
| ref-cast-impl | 1.0.27 | MIT OR Apache-2.0 | macOS build |
| regex | 1.13.1 | MIT OR Apache-2.0 | macOS build |
| regex-automata | 0.4.18 | MIT OR Apache-2.0 | macOS build |
| regex-syntax | 0.8.11 | MIT OR Apache-2.0 | macOS build |
| reqwest | 0.13.5 | MIT OR Apache-2.0 | other targets |
| rfd | 0.16.0 | MIT | macOS build |
| rustc-hash | 2.1.3 | Apache-2.0 OR MIT | macOS build |
| rustc_version | 0.4.1 | MIT OR Apache-2.0 | macOS build |
| rustversion | 1.0.23 | MIT OR Apache-2.0 | macOS build |
| same-file | 1.0.6 | Unlicense/MIT | macOS build |
| schemars | 0.8.22 | MIT | macOS build |
| schemars | 0.9.0 | MIT | macOS build |
| schemars | 1.2.2 | MIT | macOS build |
| schemars_derive | 0.8.22 | MIT | macOS build |
| scopeguard | 1.2.0 | MIT OR Apache-2.0 | macOS build |
| selectors | 0.36.1 | MPL-2.0 | macOS build |
| semver | 1.0.28 | MIT OR Apache-2.0 | macOS build |
| serde | 1.0.229 | MIT OR Apache-2.0 | macOS build |
| serde-untagged | 0.1.9 | MIT OR Apache-2.0 | macOS build |
| serde_core | 1.0.229 | MIT OR Apache-2.0 | macOS build |
| serde_derive | 1.0.229 | MIT OR Apache-2.0 | macOS build |
| serde_derive_internals | 0.29.1 | MIT OR Apache-2.0 | macOS build |
| serde_json | 1.0.151 | MIT OR Apache-2.0 | macOS build |
| serde_repr | 0.1.21 | MIT OR Apache-2.0 | macOS build |
| serde_spanned | 0.6.9 | MIT OR Apache-2.0 | other targets |
| serde_spanned | 1.1.1 | MIT OR Apache-2.0 | macOS build |
| serde_with | 3.23.0 | MIT OR Apache-2.0 | macOS build |
| serde_with_macros | 3.23.0 | MIT OR Apache-2.0 | macOS build |
| serialize-to-javascript | 0.1.2 | MIT OR Apache-2.0 | macOS build |
| serialize-to-javascript-impl | 0.1.2 | MIT OR Apache-2.0 | macOS build |
| servo_arc | 0.4.3 | MIT OR Apache-2.0 | macOS build |
| sha2 | 0.10.9 | MIT OR Apache-2.0 | macOS build |
| shared_child | 1.1.2 | MIT | macOS build |
| shlex | 2.0.1 | MIT OR Apache-2.0 | macOS build |
| sigchld | 0.2.5 | MIT | macOS build |
| signal-hook | 0.4.4 | MIT OR Apache-2.0 | macOS build |
| signal-hook-registry | 1.4.8 | MIT OR Apache-2.0 | macOS build |
| simd-adler32 | 0.3.10 | MIT | macOS build |
| simdutf8 | 0.1.5 | MIT OR Apache-2.0 | macOS build |
| siphasher | 1.0.3 | MIT/Apache-2.0 | macOS build |
| slab | 0.4.12 | MIT | other targets |
| smallvec | 1.16.1 | MIT OR Apache-2.0 | macOS build |
| socket2 | 0.6.5 | MIT OR Apache-2.0 | macOS build |
| softbuffer | 0.4.8 | MIT OR Apache-2.0 | other targets |
| soup3 | 0.5.0 | MIT | other targets |
| soup3-sys | 0.5.0 | MIT | other targets |
| stable_deref_trait | 1.2.1 | MIT OR Apache-2.0 | macOS build |
| string_cache | 0.9.0 | MIT OR Apache-2.0 | macOS build |
| string_cache_codegen | 0.6.1 | MIT OR Apache-2.0 | macOS build |
| strsim | 0.11.1 | MIT | macOS build |
| swift-rs | 1.0.8 | MIT OR Apache-2.0 | macOS build |
| syn | 1.0.109 | MIT OR Apache-2.0 | other targets |
| syn | 2.0.119 | MIT OR Apache-2.0 | macOS build |
| syn | 3.0.6 | MIT OR Apache-2.0 | macOS build |
| sync_wrapper | 1.0.2 | Apache-2.0 | other targets |
| synstructure | 0.14.0 | MIT | macOS build |
| sys-locale | 0.3.2 | MIT OR Apache-2.0 | macOS build |
| system-deps | 6.2.2 | MIT OR Apache-2.0 | other targets |
| tao | 0.35.3 | Apache-2.0 | macOS build |
| tao-macros | 0.1.4 | MIT OR Apache-2.0 | other targets |
| target-lexicon | 0.12.16 | Apache-2.0 WITH LLVM-exception | other targets |
| tauri | 2.11.6 | Apache-2.0 OR MIT | macOS build |
| tauri-build | 2.6.3 | Apache-2.0 OR MIT | macOS build |
| tauri-codegen | 2.6.3 | Apache-2.0 OR MIT | macOS build |
| tauri-macros | 2.6.3 | Apache-2.0 OR MIT | macOS build |
| tauri-plugin | 2.6.3 | Apache-2.0 OR MIT | macOS build |
| tauri-plugin-dialog | 2.7.3 | Apache-2.0 OR MIT | macOS build |
| tauri-plugin-fs | 2.5.2 | Apache-2.0 OR MIT | macOS build |
| tauri-plugin-shell | 2.3.6 | Apache-2.0 OR MIT | macOS build |
| tauri-runtime | 2.11.3 | Apache-2.0 OR MIT | macOS build |
| tauri-runtime-wry | 2.11.4 | Apache-2.0 OR MIT | macOS build |
| tauri-utils | 2.9.3 | Apache-2.0 OR MIT | macOS build |
| tauri-winres | 0.3.6 | MIT | macOS build |
| tendril | 0.5.1 | MIT OR Apache-2.0 | macOS build |
| thiserror | 1.0.69 | MIT OR Apache-2.0 | macOS build |
| thiserror | 2.0.20 | MIT OR Apache-2.0 | macOS build |
| thiserror-impl | 1.0.69 | MIT OR Apache-2.0 | macOS build |
| thiserror-impl | 2.0.20 | MIT OR Apache-2.0 | macOS build |
| time | 0.3.55 | MIT OR Apache-2.0 | macOS build |
| time-core | 0.1.9 | MIT OR Apache-2.0 | macOS build |
| time-macros | 0.2.32 | MIT OR Apache-2.0 | macOS build |
| tinystr | 0.8.4 | Unicode-3.0 | macOS build |
| tinyvec | 1.13.3 | Zlib OR Apache-2.0 OR MIT | macOS build |
| tokio | 1.53.1 | MIT | macOS build |
| tokio-util | 0.7.19 | MIT | other targets |
| toml | 0.8.2 | MIT OR Apache-2.0 | other targets |
| toml | 0.9.12+spec-1.1.0 | MIT OR Apache-2.0 | macOS build |
| toml | 1.1.6+spec-1.1.0 | MIT OR Apache-2.0 | macOS build |
| toml_datetime | 0.6.3 | MIT OR Apache-2.0 | other targets |
| toml_datetime | 0.7.5+spec-1.1.0 | MIT OR Apache-2.0 | macOS build |
| toml_datetime | 1.1.1+spec-1.1.0 | MIT OR Apache-2.0 | macOS build |
| toml_edit | 0.19.15 | MIT OR Apache-2.0 | other targets |
| toml_edit | 0.20.2 | MIT OR Apache-2.0 | other targets |
| toml_edit | 0.25.15+spec-1.1.0 | MIT OR Apache-2.0 | other targets |
| toml_parser | 1.1.3+spec-1.1.0 | MIT OR Apache-2.0 | macOS build |
| toml_writer | 1.1.2+spec-1.1.0 | MIT OR Apache-2.0 | macOS build |
| tower | 0.5.3 | MIT | other targets |
| tower-http | 0.6.11 | MIT | other targets |
| tower-layer | 0.3.3 | MIT | other targets |
| tower-service | 0.3.3 | MIT | other targets |
| tracing | 0.1.44 | MIT | other targets |
| tracing-core | 0.1.36 | MIT | other targets |
| tray-icon | 0.24.2 | MIT OR Apache-2.0 | macOS build |
| try-lock | 0.2.5 | MIT | other targets |
| typeid | 1.0.3 | MIT OR Apache-2.0 | macOS build |
| typenum | 1.20.1 | MIT OR Apache-2.0 | macOS build |
| unic-char-property | 0.9.0 | MIT/Apache-2.0 | macOS build |
| unic-char-range | 0.9.0 | MIT/Apache-2.0 | macOS build |
| unic-common | 0.9.0 | MIT/Apache-2.0 | macOS build |
| unic-ucd-ident | 0.9.0 | MIT/Apache-2.0 | macOS build |
| unic-ucd-version | 0.9.0 | MIT/Apache-2.0 | macOS build |
| unicode-ident | 1.0.26 | (MIT OR Apache-2.0) AND Unicode-3.0 | macOS build |
| unicode-segmentation | 1.13.3 | MIT OR Apache-2.0 | macOS build |
| url | 2.5.8 | MIT OR Apache-2.0 | macOS build |
| urlpattern | 0.3.0 | MIT | macOS build |
| utf8_iter | 1.0.4 | Apache-2.0 OR MIT | macOS build |
| uuid | 1.26.1 | Apache-2.0 OR MIT | macOS build |
| version-compare | 0.2.1 | MIT | other targets |
| version_check | 0.9.5 | MIT/Apache-2.0 | macOS build |
| vswhom | 0.1.0 | MIT | other targets |
| vswhom-sys | 0.1.3 | MIT | other targets |
| walkdir | 2.5.0 | Unlicense/MIT | macOS build |
| want | 0.3.1 | MIT | other targets |
| wasi | 0.11.1+wasi-snapshot-preview1 | Apache-2.0 WITH LLVM-exception OR Apache-2.0 OR MIT | other targets |
| wasip2 | 1.0.4+wasi-0.2.12 | Apache-2.0 WITH LLVM-exception OR Apache-2.0 OR MIT | other targets |
| wasm-bindgen | 0.2.128 | MIT OR Apache-2.0 | other targets |
| wasm-bindgen-futures | 0.4.78 | MIT OR Apache-2.0 | other targets |
| wasm-bindgen-macro | 0.2.128 | MIT OR Apache-2.0 | other targets |
| wasm-bindgen-macro-support | 0.2.128 | MIT OR Apache-2.0 | other targets |
| wasm-bindgen-shared | 0.2.128 | MIT OR Apache-2.0 | other targets |
| wasm-streams | 0.5.0 | MIT OR Apache-2.0 | other targets |
| web-sys | 0.3.105 | MIT OR Apache-2.0 | other targets |
| web_atoms | 0.2.6 | MIT OR Apache-2.0 | macOS build |
| webkit2gtk | 2.0.2 | MIT | other targets |
| webkit2gtk-sys | 2.0.2 | MIT | other targets |
| webview2-com | 0.38.2 | MIT | other targets |
| webview2-com-macros | 0.8.1 | MIT | other targets |
| webview2-com-sys | 0.38.2 | MIT | other targets |
| winapi | 0.3.9 | MIT/Apache-2.0 | other targets |
| winapi-i686-pc-windows-gnu | 0.4.0 | MIT/Apache-2.0 | other targets |
| winapi-util | 0.1.11 | Unlicense OR MIT | other targets |
| winapi-x86_64-pc-windows-gnu | 0.4.0 | MIT/Apache-2.0 | other targets |
| window-vibrancy | 0.6.0 | Apache-2.0 OR MIT | macOS build |
| windows | 0.61.3 | MIT OR Apache-2.0 | other targets |
| windows-collections | 0.2.0 | MIT OR Apache-2.0 | other targets |
| windows-core | 0.61.2 | MIT OR Apache-2.0 | other targets |
| windows-core | 0.62.2 | MIT OR Apache-2.0 | other targets |
| windows-future | 0.2.1 | MIT OR Apache-2.0 | other targets |
| windows-implement | 0.60.2 | MIT OR Apache-2.0 | other targets |
| windows-interface | 0.59.3 | MIT OR Apache-2.0 | other targets |
| windows-link | 0.1.3 | MIT OR Apache-2.0 | other targets |
| windows-link | 0.2.1 | MIT OR Apache-2.0 | other targets |
| windows-numerics | 0.2.0 | MIT OR Apache-2.0 | other targets |
| windows-result | 0.3.4 | MIT OR Apache-2.0 | other targets |
| windows-result | 0.4.1 | MIT OR Apache-2.0 | other targets |
| windows-strings | 0.4.2 | MIT OR Apache-2.0 | other targets |
| windows-strings | 0.5.1 | MIT OR Apache-2.0 | other targets |
| windows-sys | 0.45.0 | MIT OR Apache-2.0 | other targets |
| windows-sys | 0.59.0 | MIT OR Apache-2.0 | other targets |
| windows-sys | 0.60.2 | MIT OR Apache-2.0 | other targets |
| windows-sys | 0.61.2 | MIT OR Apache-2.0 | other targets |
| windows-targets | 0.42.2 | MIT OR Apache-2.0 | other targets |
| windows-targets | 0.52.6 | MIT OR Apache-2.0 | other targets |
| windows-targets | 0.53.5 | MIT OR Apache-2.0 | other targets |
| windows-threading | 0.1.0 | MIT OR Apache-2.0 | other targets |
| windows-version | 0.1.7 | MIT OR Apache-2.0 | other targets |
| windows_aarch64_gnullvm | 0.42.2 | MIT OR Apache-2.0 | other targets |
| windows_aarch64_gnullvm | 0.52.6 | MIT OR Apache-2.0 | other targets |
| windows_aarch64_gnullvm | 0.53.1 | MIT OR Apache-2.0 | other targets |
| windows_aarch64_msvc | 0.42.2 | MIT OR Apache-2.0 | other targets |
| windows_aarch64_msvc | 0.52.6 | MIT OR Apache-2.0 | other targets |
| windows_aarch64_msvc | 0.53.1 | MIT OR Apache-2.0 | other targets |
| windows_i686_gnu | 0.42.2 | MIT OR Apache-2.0 | other targets |
| windows_i686_gnu | 0.52.6 | MIT OR Apache-2.0 | other targets |
| windows_i686_gnu | 0.53.1 | MIT OR Apache-2.0 | other targets |
| windows_i686_gnullvm | 0.52.6 | MIT OR Apache-2.0 | other targets |
| windows_i686_gnullvm | 0.53.1 | MIT OR Apache-2.0 | other targets |
| windows_i686_msvc | 0.42.2 | MIT OR Apache-2.0 | other targets |
| windows_i686_msvc | 0.52.6 | MIT OR Apache-2.0 | other targets |
| windows_i686_msvc | 0.53.1 | MIT OR Apache-2.0 | other targets |
| windows_x86_64_gnu | 0.42.2 | MIT OR Apache-2.0 | other targets |
| windows_x86_64_gnu | 0.52.6 | MIT OR Apache-2.0 | other targets |
| windows_x86_64_gnu | 0.53.1 | MIT OR Apache-2.0 | other targets |
| windows_x86_64_gnullvm | 0.42.2 | MIT OR Apache-2.0 | other targets |
| windows_x86_64_gnullvm | 0.52.6 | MIT OR Apache-2.0 | other targets |
| windows_x86_64_gnullvm | 0.53.1 | MIT OR Apache-2.0 | other targets |
| windows_x86_64_msvc | 0.42.2 | MIT OR Apache-2.0 | other targets |
| windows_x86_64_msvc | 0.52.6 | MIT OR Apache-2.0 | other targets |
| windows_x86_64_msvc | 0.53.1 | MIT OR Apache-2.0 | other targets |
| winnow | 0.5.40 | MIT | other targets |
| winnow | 0.7.15 | MIT | macOS build |
| winnow | 1.0.4 | MIT | macOS build |
| winreg | 0.55.0 | MIT | other targets |
| wit-bindgen | 0.57.1 | Apache-2.0 WITH LLVM-exception OR Apache-2.0 OR MIT | other targets |
| writeable | 0.6.4 | Unicode-3.0 | macOS build |
| wry | 0.55.1 | Apache-2.0 OR MIT | macOS build |
| x11 | 2.21.0 | MIT | other targets |
| x11-dl | 2.21.0 | MIT | other targets |
| yoke | 0.8.3 | Unicode-3.0 | macOS build |
| yoke-derive | 0.8.3 | Unicode-3.0 | macOS build |
| zerofrom | 0.1.8 | Unicode-3.0 | macOS build |
| zerofrom-derive | 0.1.8 | Unicode-3.0 | macOS build |
| zerotrie | 0.2.5 | Unicode-3.0 | macOS build |
| zerovec | 0.11.8 | Unicode-3.0 | macOS build |
| zerovec-derive | 0.11.6 | Unicode-3.0 | macOS build |
| zlib-rs | 0.6.8 | Zlib | macOS build |
| zmij | 1.0.23 | MIT | macOS build |

</details>
