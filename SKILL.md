---
name: writer
description: Read, create and edit Word (.docx), Excel (.xlsx), PowerPoint (.pptx) and Markdown (.md) files, and read PDFs, through one path-addressed CLI or MCP tool. Use whenever the user wants to inspect, write, fix or convert any of these documents.
---

# writer

One binary, one tree model for every document. Every element has a path, every command prints JSON, and
edits keep everything you did not touch exactly as it was (comments, tracked changes, charts, animations).

## Install

If `writer` is not on PATH and the user is on a Mac (macOS 14 or later, Apple silicon), install the Writer app. Neither
way needs any clicks:

```bash
brew install --cask auspexlabs/tap/writer    # the app in /Applications, and `writer` on PATH
```

Without Homebrew:

```bash
curl -fL -o /tmp/Writer.dmg https://github.com/Auspexlabs/writer/releases/latest/download/Writer-mac.dmg
hdiutil attach -nobrowse -quiet /tmp/Writer.dmg -mountpoint /tmp/writer-dmg
cp -R /tmp/writer-dmg/Writer.app /Applications/ && hdiutil detach -quiet /tmp/writer-dmg
# the CLI is /Applications/Writer.app/Contents/MacOS/writer
```

## Run

CLI: `writer <command> <file> [path] [options]`. MCP: register `writer mcp` (stdio) as a server; its `writer` tool
takes the same command line as a single `command` string (no program name). From a source build the binary is
`dist/<platform>/writer`.

To connect yourself, add Writer to your MCP client's config, or use the client's own add-server command. Use the full
path, because apps opened from the Dock do not see the shell's PATH. Restart the client if it loads servers only at
startup.

```json
{ "mcpServers": { "writer": { "command": "/Applications/Writer.app/Contents/MacOS/writer", "args": ["mcp"] } } }
```

## Workflow

1. `writer view <file> outline` — every element with its path and a text preview. Always start here.
2. `writer get <file> <path>` or `writer query <file> <path>` — inspect one node, or every match.
3. `writer add | set | remove | move | copy` — change the file. Each command saves atomically.
4. `writer view <file> html` — check the result reads the way you intended.
5. `writer help <format> <element>` — the properties an element accepts. Run this instead of guessing.

## Paths

| Path | Meaning |
|---|---|
| `/body/paragraph[2]` | 2nd paragraph in a docx or md body |
| `//heading` | every heading in the file |
| `//heading[3]` | 3rd heading in the whole file |
| `//paragraph[@text~="Q4"]` | paragraphs containing Q4 (`=` for an exact match) |
| `/body/table[1]/row[-1]/cell[2]` | last row of the first table, second cell |
| `/body/*[3]` | 3rd block of any kind |
| `/slide[2]/shape[1]/paragraph[1]/run[1]` | pptx: slide → shape → paragraph → run |
| `/sheet[Sales]/cell[B3]`, `/sheet[1]/range[A1:C10]`, `/sheet[1]/row[4]` | xlsx: cells by reference, ranges, rows by number |
| `/page[1]/text[2]` | pdf: text blocks on a page (read-only) |

## Commands

```
writer create out.docx [--from template.docx]
writer get file.docx /body/paragraph[2] [--depth 2] [--raw]
writer query file.docx "//paragraph[@text~=\"Revenue\"]"
writer add file.docx /body --type heading --prop text="Summary" --prop level=1 [--index n | --after path | --before path]
writer set file.docx /body/paragraph[2] --prop md="Revenue **grew** 25%" [--all] [--raw "<w:p>...</w:p>"]
writer remove file.docx /body/paragraph[3] [--all]
writer move file.docx /body/table[1] --to /body --index 1
writer copy deck.pptx /slide[2] --to / --after /slide[2]    # pptx: a slide, or a shape, picture or table --to /slide[n]
writer view file.docx outline|text|html|json
writer export file.docx --to file.md        # md, docx, html, json targets
writer help [docx|xlsx|pptx|md|pdf] [element] [--json]
```

Values: lengths accept `2cm`, `20mm`, `1in`, `12pt`, `96px`; colors accept `FF0000`, `#FF0000`, `red`, `none`;
`--prop md=` takes inline markdown and works in every format; `--prop html=` takes inline HTML (b, i, u, s, code,
a, br, span style color/font-size/background-color/font-family) and reads back the content as HTML; `--prop text=`
is plain text. Image `src` accepts a file path or a `data:image/png;base64,...` URL.

Pictures in docx, pptx and xlsx (sheets: `add f.xlsx /sheet[1] --type image --prop src=logo.png`) take Office's own
picture adjustments: `crop=l,t,r,b` (percent cut off each edge; the frame keeps the scale), `rotation`, `flipH`, `flipV`,
`brightness` and `contrast` (-100…100), `grayscale`, `transparency` (0…100), `line` + `lineWidth` (border), `shadow`,
`geometry` (`roundRect` for rounded corners, `ellipse`, `triangle`, …); `0`, `false`, `none` or `rect` take one away.
`reset=true` drops them all and brings back the original. `background=remove` cuts the subject out (Apple Vision, macOS
14+) and `compress=print|web|email` re-encodes at the resolution the frame needs; both need the `writer-vision` helper
(next to `writer`, `WRITER_VISION`, or on PATH). `bytes` reads the stored size.

## Per format

- **docx**: body → heading, paragraph, code, table/row/cell, image, pagebreak, toc (`add f.docx /body --type toc --prop title=目录`; `set` regenerates it from the headings); tables take `style` (`ThreeLineTable` is a 三线表), `borders` (`style` drops the override), `borderColor`, `width`, `widths`, `align`, rows `header`/`height`, cells `colspan`/`rowspan`/`fill`/`borders`/`valign`/`width`; runs inside paragraphs. Lists are
  paragraphs with `list=bullet|number` and `level`. Styles by id or name (`Heading 1`, `Quote`, `Title`). Page setup
  is on the document: `set f.docx / --prop page=Letter --prop orientation=landscape --prop margin=narrow --prop columns=2
  --prop footer="Page {page} of {pages}"` (`page` also takes `21cm x 29.7cm`, `margin` one or four lengths, `header` and
  `footer` inline HTML; empty removes them). Anything not modelled (footnotes, comments, fields, tracked changes) is
  preserved untouched.
- **md**: same elements as docx. Untouched blocks are written back byte for byte.
- **pptx**: document → slide (title, layout, background) → shape (text, geometry, x/y/w/h, fill, line, font),
  image, table. `add / --type slide --prop layout=Content --prop title="Agenda"`; `add /slide[2] --type shape --prop md="..."`.
  Slides also carry `notes`, `hidden`, `transition` (none, fade, push, wipe, split, cover, cut, dissolve, zoom, random; `other` when the file uses an effect the engine does not model) and `duration` (ms, PowerPoint's `p14:dur`; the older `spd` speed is not exposed). `background` reads through to the layout's, then the master's solid colour; `backgroundImage=true` means the effective background is a picture. `/slide[n]/decor[m]` lists the layout/master shapes, pictures and background picture the slide shows, in drawing order (read-only, before its own shapes, so `shape[k]` ordinals do not change).
- **xlsx**: `set f.xlsx /sheet[1]/cell[B3] --prop value=42` creates the cell if needed. Numbers, true/false and
  `yyyy-mm-dd` are typed automatically; `--prop type=string` forces text. `range[A1:C10]` reads and writes a
  block as JSON rows or CSV. `formula=SUM(B2:B9)` stores a formula (Excel calculates it on open).
  Formatting on a cell or a whole range: `bold`, `italic`, `underline`, `strike`, `size=12pt`, `font=Arial`,
  `color`, `fill`, `format=0.00`, `align=center`, `valign=middle`, `wrap`, `indent`, `border=thin` (all sides),
  `borders={"bottom":"medium"}` (given sides only), `borderColor=FF0000`; e.g.
  `set f.xlsx /sheet[1]/range[A1:D1] --prop bold=true --prop fill=D9E2F3`. `link=https://…` or `link=#Sheet2!A1`
  adds a hyperlink and `note="Reviewed"` a comment; an empty value removes either.
  Sheet layout lives on the sheet: `merges`, `widths`, `heights`, `freeze`, `filter`. Charts are children of the sheet:
  `add f.xlsx /sheet[1] --type chart --prop type=column --prop categories=A2:A6 --prop series='[{"name":"B1","values":"B2:B6"}]' --prop title=Sales`;
  `set` on `/sheet[1]/chart[1]` changes type, title, series, legend, stacked and x/y/w/h in place.
- **pdf**: read-only. `view text`, `query //text[@text~="..."]`, then `export --to file.md` to edit.
- **mm**: FreeMind mind map. document → one root `topic` (`/topic[1]`) → nested topics (text, note, collapsed, side, link, color, fill, icon); `//topic[@id=ID_123]` addresses a node by its file id. `add map.mm /topic[1] --type topic --prop text="Idea" [--index n]`; `export` to and from md, docx, pptx.

## Errors

Errors are JSON on stderr with a non-zero exit: `{"error":{"code":"PATH_NOT_FOUND","message":"...","hint":"..."}}`.
The `hint` says what to do next. Codes: USAGE, FILE_NOT_FOUND, PATH_NOT_FOUND, PATH_AMBIGUOUS (add `--all` or pick one),
VALIDATION, UNSUPPORTED_KIND, FORMAT_READONLY, FORMAT_ERROR, UNKNOWN_FORMAT, IO.
