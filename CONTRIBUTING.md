# 参与贡献 · Contributing

[中文](#中文) · [English](#english)

## 中文

谢谢你愿意帮助 Writer。发现问题或者有想法，请在 GitHub Issues 里提出。

### 贡献者许可协议（CLA）

Writer 同时以 AGPL-3.0 和商业许可提供（见 [COMMERCIAL.md](COMMERCIAL.md)）。为了把贡献的代码也纳入这两种许可，贡献者需要同意一份贡献者许可协议（CLA），允许 Auspex 以其他许可证发布这些代码。

正式的 CLA 流程还在准备中，上线之前暂不合并外部的 Pull Request。欢迎先开 Issue 讨论。

### 构建和测试

需要 .NET 10 SDK 和 Node 20。做桌面应用还需要 Rust 工具链和 Xcode 命令行工具（Windows 上是 Visual Studio C++ 构建工具和 WebView2）。

```bash
dotnet test Writer.slnx        # 引擎：单元、格式适配器、命令行和往返保真测试
node --test ui/tests/          # 编辑器里不依赖浏览器的逻辑
./build.sh osx-arm64           # 单文件 writer，输出到 dist/osx-arm64/（不带参数则构建全部平台）
dotnet run --project src/Writer.Cli -- app --dir ~/Documents   # 用仓库里的 ui/ 在浏览器里打开编辑器
```

桌面应用：

```bash
cd desktop
npm ci
npm run dev                    # 构建内置引擎，打开开发窗口
npm run prepare:engine && (cd src-tauri && cargo test)   # 桌面壳的 Rust 单元测试
```

打包、签名和诊断见 [desktop/README.md](desktop/README.md)，引擎的命令和接口见 [docs/engine.md](docs/engine.md)。

几条约定：

- C# 代码开启了 `TreatWarningsAsErrors`，警告会让构建失败。
- 改了引擎的行为，在 `tests/Writer.Tests` 里补测试；改了编辑器的逻辑，在 `ui/tests` 里补测试。
- 官网（`website/`）不发任何外部请求，说明见 [website/README.md](website/README.md)。

### 提交信息

标题是一行英文，格式是 `范围: 改了什么`：

- 范围是改动所在的部分，比如 `engine`、`ui`、`desktop`、`website`、`docs`，也可以是具体的格式或功能，比如 `xlsx`、`storage`。
- 说明写改完以后的效果，小写开头，结尾不加句号。界面上的中文词可以照写，比如 `存储`、`另存为`。
- 需要解释的，空一行再写正文：为什么改、改了什么、怎样验证的（常以一段 `Tests: …` 结尾）。

例子（来自 `git log`）：

```
engine: copy and add --raw for pptx; the slide editor saves copies and undone deletions exactly
xlsx: sheets are saved in the editor's order; inserting, moving or removing a sheet keeps what refers to its position
ui: a comment author's avatar colour comes from their name, the same in every document
```

## English

Thank you for helping with Writer. Please report bugs and share ideas in GitHub Issues.

### Contributor License Agreement (CLA)

Writer is offered under both the AGPL-3.0 and a commercial license (see [COMMERCIAL.md](COMMERCIAL.md)). So that
contributions can be offered under both, contributors need to agree to a Contributor License Agreement (CLA) that lets
Auspex license their contributions under other terms as well.

A formal CLA process is being set up; until it is in place, external pull requests are not merged. You are welcome to
open an issue to discuss a change first.

### Build and test

You need the .NET 10 SDK and Node 20. For the desktop app you also need a Rust toolchain and the Xcode command line
tools (on Windows, the Visual Studio C++ build tools and WebView2).

```bash
dotnet test Writer.slnx        # engine: unit, format adapter, command-line and round-trip fidelity tests
node --test ui/tests/          # the editors' logic that runs without a browser
./build.sh osx-arm64           # single-file writer in dist/osx-arm64/ (no argument builds every platform)
dotnet run --project src/Writer.Cli -- app --dir ~/Documents   # the editor from the repository's ui/, in a browser
```

The desktop app:

```bash
cd desktop
npm ci
npm run dev                    # builds the bundled engine and opens a development window
npm run prepare:engine && (cd src-tauri && cargo test)   # the desktop shell's Rust unit tests
```

Packaging, signing and diagnostics are in [desktop/README.md](desktop/README.md); the engine's commands and APIs are in
[docs/engine.md](docs/engine.md).

A few conventions:

- C# builds with `TreatWarningsAsErrors`: a warning fails the build.
- A change to the engine's behaviour comes with a test in `tests/Writer.Tests`; a change to an editor's logic, with one
  in `ui/tests`.
- The website (`website/`) makes no external requests; see [website/README.md](website/README.md).

### Commit messages

The subject is one line in English, in the form `area: what changed`:

- The area is the part of the project you changed, such as `engine`, `ui`, `desktop`, `website` or `docs`, or a
  specific format or feature, such as `xlsx` or `storage`.
- The description says what is true after the change, starts in lower case and has no full stop at the end. Chinese UI
  terms can be quoted as they appear, such as `存储` or `另存为`.
- When a change needs explaining, add a body after a blank line: why, what changed and how it was checked (often a
  closing `Tests: …` paragraph).

Examples from `git log`:

```
engine: copy and add --raw for pptx; the slide editor saves copies and undone deletions exactly
xlsx: sheets are saved in the editor's order; inserting, moving or removing a sheet keeps what refers to its position
ui: a comment author's avatar colour comes from their name, the same in every document
```
