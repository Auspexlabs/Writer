# Writer engine reference

The full reference for the `writer` engine. For the project as a whole (the Mac app, downloads, license), see the
[README](../README.en.md) ([中文](../README.md)).

A document engine for AI agents. One binary reads, creates and edits Word, Excel, PowerPoint and Markdown
files, reads PDFs, converts between formats, and exposes everything through a CLI and an MCP server.

Every document becomes one tree. Every element has a path. Every command prints JSON. Edits are made in the
native file structure, so what you did not touch stays byte for byte the same.

## Status

P0 kernel, complete:

| Format | Read | Create | Edit | Fidelity samples |
|---|---|---|---|---|
| docx | yes | yes | yes | 14 files |
| md | yes | yes | yes | 87 files, byte-exact |
| pptx | yes | yes | yes | 12 decks |
| xlsx | yes | yes | yes | 10 workbooks |
| pdf | yes | no | no (export first) | generated |
| mm | yes | yes | yes | FreeMind mind maps: `topic` kind, `//topic[@id=…]` paths, whitespace kept; exports to and from md, docx, pptx |

Export targets: md, docx, pptx (headings become slides), xlsx (tables become sheets), html, json.
xlsx cells read and write the full Excel formatting set (font, fill, borders, alignment, number format), hyperlinks
and notes; a range applies formatting to every cell it covers.
Excel sheets also carry their layout (merges, column widths, row heights, frozen panes, autofilter) and real charts
(`/sheet[1]/chart[1]`: column, bar, line, pie, area, scatter, doughnut) that Excel opens as its own.
Also done: a local HTTP API and live preview (`serve`, `watch`), slide rendering at real positions, and the
editor app (`writer app`, see below) with its built-in assistant. Still to come: pixel-faithful page rendering
and PNG output (P1), the remaining conversions (P2), semantic operations such as templates and reflow (P3).

## Quick start

```bash
dotnet run --project src/Writer.Cli -- create report.docx
dotnet run --project src/Writer.Cli -- add report.docx /body --type heading --prop text="Q4 Report" --prop level=1
dotnet run --project src/Writer.Cli -- add report.docx /body --type paragraph --prop md="Revenue **grew** 25%."
dotnet run --project src/Writer.Cli -- add report.docx /body --type table --prop data='[["Region","Sales"],["East","120"]]'
dotnet run --project src/Writer.Cli -- add report.docx /body --type pagebreak
dotnet run --project src/Writer.Cli -- set report.docx / --prop page=Letter --prop footer="Page {page} of {pages}"
dotnet run --project src/Writer.Cli -- view report.docx outline
dotnet run --project src/Writer.Cli -- export report.docx --to report.md
```

Build a standalone binary with `./build.sh` (all platforms) or `./build.sh osx-arm64` (one), then use
`dist/<rid>/writer` directly.

## Commands

```
create <file> [--from template]                         new blank document, or a copy of a template
get <file> [path] [--depth n] [--raw]                   one node as JSON (path defaults to /)
query <file> <path>                                     every node matching a path
add <file> <parent> --type kind [--prop k=v]...         add an element (--index n | --after path | --before path)
add <file> <parent> --raw xml                           put an element back from its 'get --raw' output (pptx)
set <file> <path> [--prop k=v]... [--raw xml] [--all]   change properties, or replace the raw XML
remove <file> <path> [--all]                            delete elements
move <file> <path> --to parent [--index n]              move an element
copy <file> <path> --to parent [--index n]              copy a pptx slide, or a shape, picture or table, exactly
view <file> [outline|structure|text|html|json]          the whole document; structure is the short form
search <file> <text> [--ignore-case]                    blocks whose text contains the words, with paths
section <file> <heading path> [--md text | --remove]    a heading with its blocks: print, rewrite from markdown, remove
replace <file> [path] --find a --with b [--preview]     find and replace; Word paragraphs keep their formatting
formula <file> <cell|range path> <formula> [--check]    xlsx: write (a range fills like Excel) and report references
batch <file> --run "command" [--run ...]                several commands, one save; nothing written when one fails
export <file> --to out.ext                              md, docx, pptx, xlsx, html, json
help [format] [element] [--json]                        elements and properties, generated from the registry
mcp                                                     MCP server over stdio
serve [--dir folder] [--port n] [--no-token]            local HTTP API for a UI (default port 26315)
watch <file> [--port n]                                 live browser preview, refreshed on every change
app [--dir folder] [--port n] [--no-browser]            the editor app on a folder of documents
                                                        (serve/app: --list-depth n = folder levels /files walks, default 3)
```

## The app

```bash
export ANTHROPIC_API_KEY=sk-...      # optional: enables the assistant panel (WRITER_MODEL picks the model)
writer app --dir ~/Documents/reports
```

`writer app` opens the editor (`ui/`) in the browser on the given folder. Word, Excel, PowerPoint, Markdown and
PDF files in that folder appear as tabs; every edit is saved back through the engine within a second, so the
files keep everything the editor does not model (comments, charts, animations, styles). The assistant on the
right edits the open file with the same `writer` commands an agent would use; its changes are listed under its
reply and can be undone with one click. Dropping or opening a file copies it into the folder first.

The app is a static page under `ui/` served by the engine at `/app/`; `build.sh` copies it next to the binary, and
`WRITER_UI` or `--ui <folder>` points at another copy.
Both file icon sets (B and C) are included in `ui/assets/icons/{b,c}/` and served from `/app/assets/icons/`; the `iconStyle` setting picks one.

## Desktop app

The macOS and Windows desktop shell lives in [desktop/](../desktop/README.md): a Tauri window around this same editor with a bundled, self-contained Writer engine. `writer app --no-browser` is the launcher interface the desktop host uses (it prints the one-time bootstrap URL on stderr instead of opening a browser); plain `writer app` still opens the default browser.

## HTTP API (for a UI)

`writer serve` prints `{"url": "http://127.0.0.1:26315", "token": "..."}` on stderr and listens on
localhost only. Every request carries the token as `Authorization: Bearer <token>`; the event stream, which
cannot send headers, accepts `?token=`. A browser UI served from another local origin must be allow-listed:
`writer serve --allow-origin http://localhost:5173`. Other origins get 403, `POST /run` insists on
`Content-Type: application/json`, and the Host header must be the loopback address. `writer watch` opens the
preview with a one-time code that the first visit swaps for an HttpOnly cookie, so no token lands in the URL bar.

| Endpoint | Purpose |
|---|---|
| `POST /run` with `{"command":"view a.docx outline"}` or `{"argv":[...]}` | run any command; returns `{"code":0,"output":"..."}` or `{"code":n,"error":{...}}` |
| `GET /html?file=` | rendered preview (`view html`) |
| `GET /json?file=`, `/outline?file=`, `/text?file=` | the other views |
| `GET /events?file=` | server-sent events: `change` whenever the file is saved by anyone |
| `GET /?file=` | the built-in preview page that `watch` opens |
| `GET /files` | the documents in the workspace (`--dir`, `--list-depth` levels deep, the 500 newest), plus whether the assistant is configured |
| `GET /file?file=`, `PUT /file?file=` (`&from=` renames) | raw bytes of a workspace file, served as a download; uploads and renames |
| `GET /binary?file=&path=` | the bytes of an image node |
| `POST /chat` with `{"file","messages":[{"role","content"}]}` | the assistant: server-sent events `delta` (streamed text), `text`, `tool` (with `wrote`), `done`, `error`; needs a configured model |
| `GET /app/` | the editor app |

Relative file paths in every endpoint and in `/run` commands resolve against the workspace folder.

See [SKILL.md](../SKILL.md) for the agent workflow and path syntax.

## MCP

Register `writer mcp` with your agent, for example in a JSON config:

```json
{ "mcpServers": { "writer": { "command": "/path/to/writer", "args": ["mcp"] } } }
```

The server exposes one tool, `writer`, whose `command` argument is the CLI command line without the program
name. Its result is exactly what the CLI prints.

## Build and test

Requires the .NET 10 SDK.

```bash
dotnet test Writer.slnx
./build.sh
```

## Layout

```
src/Writer.Core      unified tree, registry, paths, JSON, views, mutations (no format libraries)
src/Writer.Formats   docx, md, pptx, xlsx, pdf adapters, HTML writer, exporter
src/Writer.Cli       command line, MCP server, HTTP server and the assistant loop
ui/                  the editor app: dc-runtime pages per format plus engine.js, the bridge to the HTTP API
tests/Writer.Tests   unit, adapter matrix, CLI and round-trip fidelity tests
source/OfficeCLI     local reference copy of OfficeCLI (Apache-2.0); gitignored, not in the repository or the build
```

Test fixtures under `tests/Writer.Tests/Fixtures` come from OfficeCLI's examples and keep their Apache-2.0
notice.
