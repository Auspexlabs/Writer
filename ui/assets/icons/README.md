# Writer 文件类型图标

软件内置 B、C 两套 64 × 64 矢量图标。两套均包含以下同名文件：

| 文件 | 格式 |
| --- | --- |
| `word.svg` | Word / `.docx` |
| `excel.svg` | Excel / `.xlsx` |
| `markdown.svg` | Markdown / `.md` |
| `mindmap.svg` | 思维导图 / `.mm` |
| `ppt.svg` | 演示文稿 / `.pptx` |
| `pdf.svg` | PDF / `.pdf` |

`b/` 是雾彩玻璃方案（默认），`c/` 是深色几何方案，由设置里的 `iconStyle`（`b` / `c`）选择。应用通过 `/app/assets/icons/<set>/<name>.svg` 访问，`build.sh` 会将整个 `ui/` 目录复制到发布包。对应的设计母版在 `brand/file-icons/concepts/`。

`ui/assets/` 下还有独立 W 和完整 Writer 字标的 SVG 与透明 PNG：`writer-mark.svg`、`writer-logo.svg`、`writer-mark-1024.png`、`writer-logo-2048.png`。品牌原稿位于 `brand/`。
