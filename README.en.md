<p align="center">
  <!-- the dark icon on light pages and the light icon on dark pages, for contrast, as in the Windows About panel -->
  <picture>
    <source media="(prefers-color-scheme: dark)" srcset="ui/assets/writer-logo-light.svg">
    <img src="ui/assets/writer-logo-dark.svg" width="96" height="96" alt="Writer icon">
  </picture>
</p>

<h1 align="center">Writer</h1>

<p align="center">
  <b>Edit Word, Excel and PowerPoint files directly on your Mac.</b><br>
  Every change is saved back to the original file within a second, and whatever you did not touch stays as it was.
  A built-in AI assistant works with the model of your choice.
</p>

<p align="center">
  <a href="https://github.com/Auspexlabs/writer/releases/latest">Download for Mac</a> ·
  <a href="docs/engine.md">Engine reference</a> ·
  <a href="README.md">中文</a>
</p>

<p align="center">
  <img src="docs/images/hero.webp" alt="The Writer window with an event plan open. In the AI panel on the right the user asks to move the event to the following weekend and update every date; the assistant has changed the opening paragraph, the schedule table and one note, the changes are marked in the document, and buttons below offer to keep or undo them.">
</p>

The app's interface is in Chinese for now.

## What it is

This repository holds two things built from the same code:

- **The Writer app**: a document app for the Mac. Word, Excel, PowerPoint, Markdown and mind-map files open ready to
  edit, and changes go straight back to the original file; PDFs open for viewing. A built-in AI assistant works with the
  model of your choice.
- **The Writer engine**: one executable, `writer`, that reads, creates and edits .docx, .xlsx, .pptx, .md and .mm files,
  reads PDFs and converts between formats. Every document is a tree, every element has a path, every command prints
  JSON. It is a command-line tool, an MCP server and a local HTTP API at once, and it is what the app's editors and
  its assistant run on.

## Features

- **Six kinds of documents, one app.** Word (.docx), Excel (.xlsx), PowerPoint (.pptx), Markdown (.md) and mind maps
  (FreeMind .mm) are all editable. PDFs can be viewed, or converted to Word to keep working on them.
- **The file stays the file.** No import, no conversion: the file you double-click in Finder is the one you are
  editing. Each change is written back within a second, and comments, charts, animations and styles you did not touch
  are kept as they were.
- **An AI assistant with your own model.** Say what to change and the assistant edits the document. The changes are
  listed under its reply: keep them, or undo them with one click. It works with Anthropic, OpenAI, DeepSeek, Qwen, Kimi,
  Zhipu GLM, Doubao, any OpenAI-compatible service, and local models through Ollama or LM Studio. Your API key stays on
  your computer.
- **Picture tools.** Remove a photo's background (with the Mac's own Apple Vision, on your Mac), crop, rotate and
  compress pictures.
- **Made the Mac way.** Native menus and shortcuts, light and dark mode, pinch to zoom, slide shows. New documents are
  kept as drafts until you save them for the first time and choose where they go.

<table>
  <tr>
    <td width="50%"><img src="docs/images/sheet.webp" alt="A spreadsheet: an event budget with formulas for what is left and the progress, next to a column chart of budget against spending."><br>Spreadsheets: formulas, charts and several sheets</td>
    <td width="50%"><img src="docs/images/slides.webp" alt="A presentation: slide thumbnails on the left and the title slide in the middle."><br>Slides: edit them, set transitions, present</td>
  </tr>
  <tr>
    <td width="50%"><img src="docs/images/mindmap.webp" alt="A mind map with one central topic and four branches."><br>Mind maps: convert them to an outline, Word or slides</td>
    <td width="50%"><img src="docs/images/markdown.webp" alt="A Markdown file shown as formatted meeting notes, with the heading outline on the left."><br>Markdown: switch between the formatted view and the source</td>
  </tr>
</table>

## Download

Install with Homebrew:

```bash
brew install --cask auspexlabs/tap/writer
```

Or download the dmg from [GitHub Releases](https://github.com/Auspexlabs/writer/releases/latest). It is signed and
notarized by Apple: open it and drag Writer into Applications.

- Requires macOS 14 or later, on a Mac with Apple silicon for now.
- A Windows version is in development.
- The official Writer app is free for everyone, including businesses and organizations that use it for work.

## For developers

The app's document engine is the `writer` command. With the app installed you can use it from a terminal:

```bash
W=/Applications/Writer.app/Contents/MacOS/writer   # installed with Homebrew? just use writer

$W view plan.docx outline          # every element with its path
$W set plan.docx '/body/table[1]/row[2]/cell[1]' --prop text="17 Oct 09:00"
$W export plan.docx --to plan.md   # convert to another format

# register it as an MCP server: add this command in your AI tool
$W mcp
```

The MCP server has one tool, `writer`, which takes a command line like the ones above (without the program name) and
returns exactly what the CLI prints. `writer serve` is a local HTTP API, and `writer app` opens the editor in a browser.

- Every command, the path syntax, the HTTP API and MCP setup: [docs/engine.md](docs/engine.md)
- How an AI agent should use it: [SKILL.md](SKILL.md)

## Build from source

The engine needs the .NET 10 SDK; the editor's tests need Node 20.

```bash
dotnet test Writer.slnx        # engine tests
node --test ui/tests/          # editor logic tests
./build.sh osx-arm64           # single-file writer in dist/osx-arm64/ (no argument: macOS, Linux and Windows)
dist/osx-arm64/writer app --dir ~/Documents   # open a folder in the editor, in your browser
```

The Mac app is built with Tauri 2 and also needs a Rust toolchain and the Xcode command line tools:

```bash
cd desktop
npm ci
npm run dev                    # builds the bundled engine and opens a development window
```

Packaging, signing and diagnostics are in [desktop/README.md](desktop/README.md).

## Project layout

```
src/          the engine (.NET 10): Writer.Core, the document tree and paths; Writer.Formats, reading and
              writing each format; Writer.Cli, the command line, MCP server, HTTP server and AI assistant
ui/           the editors, one page per format; engine.js talks to the engine
desktop/      the Mac app (Tauri 2 windows around a bundled engine); Windows in development
website/      the website, as static pages
tests/        engine tests: unit, format adapters, command line and round-trip fidelity
docs/         the engine reference, design documents and plans, screenshots for this README
brand/        logos and file icons
SKILL.md      instructions for AI agents
```

## Contributing

Bug reports and ideas are welcome in GitHub Issues. Because Writer is dual-licensed, code contributions require
agreeing to a Contributor License Agreement (CLA); until the formal CLA process is in place, external pull requests are
not merged. Building, testing and commit conventions are in [CONTRIBUTING.md](CONTRIBUTING.md).

## License

- **Source code**: open source under the [AGPL-3.0](LICENSE) (version 3 only), with additional terms that section 7
  of the AGPL allows, in [NOTICE](NOTICE): whoever distributes Writer or a modified version, or offers a
  modified version over a network, must keep the "Writer by Auspex" attribution and the copyright notice in the About
  screen and the documentation; a modified version must be marked as modified and must not be presented as the official
  Writer app.
- **Commercial license**: to use Writer's source code in your own product or service without the obligations of the
  AGPL, you can ask Auspex for a commercial license. See [COMMERCIAL.md](COMMERCIAL.md).
- **The official app is free**: the Writer app published by Auspex is free for everyone, including businesses and
  organizations that use it for work. The commercial license is only about reusing the source code.
- **Trademarks**: the Writer name and logo and the Auspex name are not licensed under the AGPL.
- **Third-party components**: used under their own licenses; see [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).

This is a summary; the LICENSE and NOTICE files are what count.

## Contact

- Commercial licensing, support and feedback: [mosheng9@outlook.com](mailto:mosheng9@outlook.com)

---

Writer by Auspex<br>
Copyright (C) 2026 北京奥斯佩克斯网络科技中心 (Auspex)
