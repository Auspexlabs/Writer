# 保真度审计（round-trip + 渲染对比）

- 日期：2026-09-25，代码基线 `main@9dfcf18`
- 工具：`node tools/fidelity/roundtrip.mjs`（一条命令，自己 build 引擎、起 `writer serve`、跑完退出；结果在 `tools/fidelity/out/report.md` 和 `report.json`）
- 样本：仓库 fixtures 37 个（docx 14、xlsx 10、pptx 13）；另跑了 OfficeCLI 的 84 个示例（`--fixtures source/OfficeCLI/examples --fixtures source/OfficeCLI/assets/showcase`，不入库），下文标 †

## 0. 结论

1. **纯打开→保存不丢任何东西。** 37 个仓库样本 save 一列全 0（`create --from` 走引擎的 Open/Save）。† OfficeCLI 的 7 个文件在 `[Content_Types].xml` / `.rels` 上有改写（见第 4 节 #20），是 Open XML SDK 的包规范化，没有零件和元素丢失。
2. **丢东西的是编辑器的写回路径**：把每个段落 / 形状 / 单元格改一遍再设回原值（等价于用户在 app 里重打一遍这段字后保存），docx 几乎无损，**pptx 丢格式、xlsx 会破坏公式**。
3. **渲染（`view html`）和 Quick Look 的第一页相比 SSIM 中位数 0.65**：docx 0.28–0.85（字体、段距、页面几何全不对），xlsx 0.10–0.72（墨迹重合率 ≈ 0：没有填充、字体、列宽、合并），pptx 0.51–0.99（最接近，缺图表、SmartArt、几何形状、表格样式）。
4. WPS 无法无头导出：`wpscli` 12.1.26046 只有 `pdf2word/pdf2excel/pdf2ppt/pdf2photo/…` 这一个方向，没有 `*2pdf`；`wpsoffice.app` 没有 AppleScript 字典（无 `.sdef`、无 `NSAppleScriptEnabled`），主程序也没有转换参数。渲染参照改用 macOS 自带的 Office 预览（`qlmanage -t`），它在 `docx/charts.docx` 上会挂死（工具 30 s 超时并 `qlmanage -r` 复位）。

## 1. 方法

**Round-trip.** 每个样本两遍：
- **save**：`create out --from fixture`（打开、保存、不改）。
- **touch**：`view json` 列出所有 docx 段落/标题（`html`）、pptx 形状（`html`）、xlsx 单元格（`value` 或 `formula`），每个先设成每个词后面加 `~` 的版本，再设回原值（每文件最多 40 个）。写回器只重写有差异的词，所以必须先改一次再改回来，每个 run 才真的经过了引擎的写路径。

两个结果和原件一起解压，XML 逐零件规范化后比较：命名空间按 URI 归一（前缀改名不算）、属性排序、`Relationships`/`Types` 子元素排序、去掉 `w:rsid*`、`w14:paraId/textId`、`xr:uid`，`docProps/core|app.xml` 不比。指标是**多重集差**：原件有而结果没有的元素、属性、属性值，以及丢的文字长度（按整个包算，共享字符串挪位置不算丢）、丢的媒体/零件、多出来的元素。

**排名** = 次数 × 出现的样本数 × 可见度（3 = 读者看得见的格式或图形，2 = 内容结构，1 = 元数据/机制）。

**渲染.** `qlmanage -t -s 1200` 出第一页参照；`view html` 用 headless Chrome 截同样的页面（A4 96 dpi，幻灯片按 `presentation.xml` 的尺寸，只留第一页/第一张），两张缩到 300×424（幻灯片 320 宽）灰度，算 8×8 窗口 SSIM 和深色像素重合率（ink）。注意：三个编辑器画的是 JSON 模型，不是这份 HTML，所以渲染分数衡量的是 HTML 视图；但**模型里没有的东西**（段距、行距、样式字体、列宽…）两边一样缺。

## 2. Round-trip 结果

save 列全为 0，下表只列 touch。

### docx（14）

| fixture | set | 变化零件 | 丢元素 | 丢属性 | 值变 | 多出的元素 | 大小 Δ |
|---|---|---|---|---|---|---|---|
| charts.docx | 16 | document.xml | 0 | 0 | 0 | 550 | +39 |
| content-controls.docx | 14 | document.xml | 0 | 0 | 0 | — | +80 |
| diagram.docx | 25 | document.xml | 0 | 0 | 0 | 1176 | +345 |
| document-formatting.docx | 6 | document.xml | 0 | 0 | 0 | 355 | +109 |
| fields.docx | 24 | document.xml | 0 | 0 | 0 | 710 | +260 |
| formulas.docx | 40 | document.xml | 0 | 0 | 0 | 515 | +565 |
| numbering.docx | 40 | document.xml | 0 | 0 | 0 | 977 | +371 |
| paragraph-formatting.docx | 40 | document.xml | 0 | 0 | 0 | 1000 | +385 |
| pictures.docx | 16 | document.xml | 0 | 0 | 0 | 1635 | +444 |
| revisions.docx | 36 | document.xml | 0 | 0 | 4（`w:ins/w:del@w:id` 重新编号） | 1084 | +714 |
| run-formatting.docx | 40 | document.xml | 0 | 0 | 0 | 437 | +294 |
| sections.docx | 22 | document.xml | 0 | 0 | 0 | 4370 | +1048 |
| tables.docx | 40 | document.xml | 0 | 0 | 0 | 64 | +139 |
| textbox.docx | 12 | document.xml | 0 | 0 | 0 | 310 | +95 |

docx 的写回（`DocxReplace`，按词做 LCS，只重写差异段）**不丢格式**。"多出的元素"全是 `w:r/w:rPr/w:t`：run 在每个改动边界被切开，恢复原文后不再合并，还留下只有 `w:rPr` 的空 run（run-formatting.docx：`w:r` 78 → 176，其中 49 个没有文字）。Word 里看不出来，但文件每保存一次都长一点。

### xlsx（10）

| fixture | set | 变化零件 | 丢元素 | 丢属性 | 值变 | 多出的元素 | 大小 Δ |
|---|---|---|---|---|---|---|---|
| cell-formatting.xlsx | 40 | sharedStrings, sheet1, sheet2 | 0 | 0 | 42 | 320（`x:si/x:t` 各 160） | +350 |
| charts-pie.xlsx | 0（没有单元格） | — | 0 | 0 | 0 | — | 0 |
| charts.xlsx | 40 | [Content_Types], workbook.rels, sheet1 | 0 | 0 | 12 | 2 | +451 |
| conditional-formatting.xlsx | 40 | 同上 +3 | 0 | 0 | 16 | 2 | +608 |
| data-validation.xlsx | 40 | 同上 +3 | 0 | 0 | 23 | 2 | +543 |
| pivot-tables.xlsx | 40 | 同上 | 0 | 0 | 28 | 2 | +621 |
| shapes.xlsx | 22 | 同上 | 0 | 0 | 22 | 2 | +548 |
| sheet-settings.xlsx | 40 | 同上 +1 | 0 | 0 | 26 | 2 | +487 |
| slicers.xlsx | 40 | 同上 | 0 | 0 | 31 | 2 | +482 |
| sparklines.xlsx | 40 | 同上 | 0 | 0 | 16 | 2 | +470 |

仓库样本里"值变"全是 `x:c@t`：`t="str"` / 内联字符串被改写成共享字符串 `t="s"`（214 个单元格 / 9 个文件），并且**每次写入都在 sharedStrings 里追加一条，不复用**（cell-formatting.xlsx：3 条 → 163 条）。† OfficeCLI 的 showcase 工作簿暴露了更严重的两条：

| † fixture | 现象 |
|---|---|
| budget-tracker.xlsx、gradebook.xlsx | 丢 `x:f` 6 个、`x:f@si` 11、`x:f@t` 11、`x:f@ref` 5：**共享公式组被打散**。主单元格设回 `formula=` 后 `<f>` 没了 `t="shared" ref si`，从属单元格的 `<f t="shared" si="n"/>` 指向不存在的主公式；引擎读从属单元格时 `formula` 为空，编辑器就按常量写回，公式变成了数字 |
| workbook-settings.xlsx、charts-boxwhisker.xlsx 等 4 个 | 丢 `x:v` 16 个：设公式时把缓存结果清掉，Excel 之外的所有阅读器（Quick Look、我们自己的 `view`、数据脚本）看到空格 |
| charts-boxwhisker.xlsx | `x:c@t="n"` 96 个被去掉（数字默认类型，语义不变，属噪音） |

### pptx（13）

| fixture | set | 变化零件 | 丢元素 | 丢属性 | 值变 | 大小 Δ |
|---|---|---|---|---|---|---|
| animations.pptx | 40 | slide1 | 2 | 2 | 0 | -176 |
| budget_review_v2.pptx | 40 | — | 0 | 0 | 0 | -130 |
| charts-pie.pptx | 8 | — | 0 | 0 | 0 | -254 |
| decor.pptx | 40 | — | 0 | 0 | 0 | +8 |
| diagram.pptx | 24 | — | 0 | 0 | 0 | -52 |
| Mars-Settlement-Guide.pptx | 40 | — | 0 | 0 | 0 | -238 |
| pictures-basic.pptx | 30 | — | 0 | 0 | 0 | -507 |
| presentation.pptx | 27 | slide1–6 | 7 | 14 | 0 | -162 |
| shapes-basic.pptx | 40 | — | 0 | 0 | 0 | -124 |
| tables-merged.pptx | 2 | — | 0 | 0 | 0 | -52 |
| tables-styled.pptx | 16 | — | 0 | 0 | 0 | +43 |
| textboxes-advanced.pptx | 30 | slide1,3,4,5 | 39 | 53 | 4 | -277 |
| transitions-morph.pptx | 4 | — | 0 | 0 | 0 | -88 |

文件变小不是好事：丢的是 `a:endParaRPr`（12 个 / 3 个文件）、第二段起的 `a:pPr`（对齐、`lnSpc/spcPct/spcPts`、`lvl`）、run 上的 `baseline`（上下标）、`cap`、`spc`、`kern`、`lang`，以及 run 合并（`a:r` 9 → 1）。† 在设计过的模板里更明显：AURA_COFFEE.pptx、太空探索历程.pptx、野生动物科技公司.pptx 三个共丢 `a:latin/a:ea@typeface` 59、`a:solidFill/a:srgbClr` 57、`a:rPr@sz` 57、`a:rPr@lang` 72——**第一个 run 之后每个 run 自己的字体、颜色、字号都没了**。

## 3. 渲染结果（第一页 vs Quick Look）

| fixture | SSIM | ink | 主要差异 |
|---|---|---|---|
| docx/charts.docx | — | — | Quick Look 挂死 |
| docx/content-controls.docx | 0.85 | 0.03 | 字体、段距 |
| docx/diagram.docx | 0.63 | 0.12 | SmartArt 不画 |
| docx/document-formatting.docx | 0.82 | 0.03 | 字体、段距、页边距 |
| docx/fields.docx | 0.64 | 0.03 | 字体、段距 |
| docx/formulas.docx | 0.66 | 0.04 | 公式（OMML）不画 |
| docx/numbering.docx | 0.59 | 0.06 | 编号格式、缩进 |
| docx/paragraph-formatting.docx | 0.59 | 0.07 | 段前后距、行距、首行缩进、边框底纹 |
| docx/pictures.docx | 0.49 | 0.09 | 图片锚点/环绕 |
| docx/revisions.docx | 0.57 | 0.05 | 字体、段距 |
| docx/run-formatting.docx | 0.78 | 0.05 | 字体 |
| docx/sections.docx | 0.28 | 0.08 | 分栏、横向节、行号、脚注全无 |
| docx/tables.docx | 0.41 | 0.15 | 表格边框、底纹、列宽 |
| docx/textbox.docx | 0.61 | 0.03 | 文本框位置 |
| pptx/animations.pptx | 0.79 | 0.99 | |
| pptx/budget_review_v2.pptx | 0.77 | 1.00 | |
| pptx/charts-pie.pptx | 0.91 | 0.47 | 图表只画标题 |
| pptx/decor.pptx | 0.80 | 0.95 | |
| pptx/diagram.pptx | 0.91 | 0.47 | SmartArt 不画 |
| pptx/Mars-Settlement-Guide.pptx | 0.81 | 0.96 | |
| pptx/pictures-basic.pptx | 0.84 | 0.90 | |
| pptx/presentation.pptx | 0.73 | 0.92 | 字间距、行距 |
| pptx/shapes-basic.pptx | 0.82 | 0.63 | rect/ellipse/roundRect 以外的几何全画成矩形 |
| pptx/tables-merged.pptx | 0.71 | 0.71 | 表格样式 |
| pptx/tables-styled.pptx | 0.67 | 0.04 | 表格样式（条纹、首行填充）不画 |
| pptx/textboxes-advanced.pptx | 0.51 | 0.17 | 段落行距、上下标、文本框内边距 |
| pptx/transitions-morph.pptx | 0.99 | 1.00 | |
| xlsx/cell-formatting.xlsx | 0.51 | 0.01 | 无填充、字体、字号、合并；列宽行高 |
| xlsx/charts-pie.xlsx | 0.91 | 0.00 | 图表只画标题（页面几乎空白所以 SSIM 高） |
| xlsx/charts.xlsx | 0.45 | 0.02 | 同上 |
| xlsx/conditional-formatting.xlsx | 0.59 | 0.02 | 条件格式不画 |
| xlsx/data-validation.xlsx | 0.61 | 0.00 | 填充、列宽 |
| xlsx/pivot-tables.xlsx | 0.10 | 0.05 | 透视表区域只有原始单元格 |
| xlsx/shapes.xlsx | 0.65 | 0.00 | 形状不画 |
| xlsx/sheet-settings.xlsx | 0.64 | 0.02 | 列宽行高、网格线 |
| xlsx/slicers.xlsx | 0.47 | 0.03 | 切片器、表格样式不画 |
| xlsx/sparklines.xlsx | 0.72 | 0.00 | 迷你图不画 |

## 4. Top 20（频率 × 可见度）

| # | 丢失 | 证据 | 可见度 | 责任代码 | 工作包 |
|---|---|---|---|---|---|
| 1 | xlsx **共享公式被打散 / 从属单元格变常量** | † budget-tracker.xlsx sheet1：`x:f` −6、`@si` −11、`@ref` −5 | 3，数据丢失 | `src/Writer.Formats/Xlsx/XlsxCells.cs:198-209` `SetFormula` 新建 `CellFormula(text)` 不带 `t/ref/si`；`XlsxNodes.cs:225` 只读 `CellFormula.Text`，从属单元格读成没有公式 | XLSX-1 |
| 2 | pptx **第一个 run 之后的字体、颜色、字号丢失**，run 合并 | † AURA_COFFEE.pptx slide2 等 3 个模板：`a:latin/a:ea` −59、`a:solidFill/a:srgbClr` −57、`a:rPr@sz` −57；textboxes-advanced.pptx slide5 `a:r` 9→1 | 3 | `src/Writer.Formats/Pptx/PptxText.cs:33-48` `SetBody` 删掉全部段落，用第一段的 `pPr` 和第一个 run 的 `rPr` 重建；`:88-103` `MakeRun` 只补 `RunSpec` 有的属性；入口 `PptxNodes.cs:375-381`（形状）、`:1080-1082`（表格单元格） | PPTX-1 |
| 3 | pptx 第二段起的**段落属性**（对齐、行距、级别、项目符号）丢失 | textboxes-advanced.pptx slide1：`a:pPr` −4、`@algn` −3、`a:lnSpc/spcPct/spcPts` −4；† textboxes-basic `@lvl` −1 | 3 | `PptxText.cs:35-36,44` | PPTX-1 |
| 4 | docx **段前后距、行距、缩进不在模型里也不渲染** | paragraph-formatting.docx SSIM 0.59；所有 docx 的 `<p>` 只有 `text-align` | 3 | 模型：`src/Writer.Formats/Docx/DocxNodes.cs:78-89` 只出 `style/list/level/align`；视图：`src/Writer.Formats/Html/HtmlWriter.cs:182-186` | DOCX-1 |
| 5 | docx **文档默认字体 / 主题字体被系统 sans-serif 替代** | 全部 docx（sections.docx 对照：Times 12pt 无段距 vs 系统字体 + 1em 段距） | 3 | `HtmlWriter.cs:12`（Css `body{font-family:-apple-system…}`），`DocxStyles.cs` 没有 docDefaults/主题字体的解析 | DOCX-1 |
| 6 | xlsx 视图**没有填充、字体、对齐、数字格式、合并、列宽行高** | cell-formatting.xlsx ink 0.006，sheet-settings 0.015；模型有这些属性（`XlsxCells`、`XlsxLayout`）但视图不用 | 3 | `HtmlWriter.cs:222-227`（sheet → `<h2>`+`Table`）、`:259-`（`Table` 只画 1px `#bbb` 边框） | XLSX-2 |
| 7 | xlsx **设公式时清掉缓存结果** | † workbook-settings.xlsx：`x:v` −16（1140、1045、3410…） | 3（Excel 外全空） | `XlsxCells.cs:204-205` `cell.CellValue = null` | XLSX-1 |
| 8 | pptx `a:endParaRPr` 丢失（空行高度、光标字体） | animations、presentation、textboxes-advanced：−12，`@lang` −12、`@sz` −10、`@b` −2 | 3 | `PptxText.cs:38-47` | PPTX-1 |
| 9 | pptx **上下标、全大写、字间距、字距**不建模，一碰就没 | textboxes-advanced slide5（H₂SO₄ 变 H2SO4）：`@baseline` −7、`@cap` −2；presentation `@spc` −4、`@kern` −1 | 3 | `src/Writer.Core/RunSpec.cs:8-10`、`Common/InlineHtml.cs:150-180`（无 sub/sup/caps/letter-spacing）、`PptxText.cs:88-103` | PPTX-1 |
| 10 | docx **页面几何**：纸张、页边距、分栏、页眉页脚、脚注、行号都不渲染 | sections.docx SSIM 0.28；`view html` 固定 820px 单栏 | 3 | `HtmlWriter.cs:12,26-41`（`Render` 无 sectPr 处理） | DOCX-1 |
| 11 | docx **标题字号来自浏览器 h1–h6 而不是样式表** | sections.docx：Word 里标题与正文同大小（样式如此），我们 2em 粗体 | 3 | `HtmlWriter.cs:176-181`，`DocxStyles.cs`（样式链 → 字号/字体/间距未解析） | DOCX-1 |
| 12 | docx **表格边框、底纹、列宽**按默认画 | tables.docx SSIM 0.41 | 3 | `HtmlWriter.cs:259-`（`Table`/`CellStyle` 只认 `borders/fill/width` 几个属性） | DOCX-1 |
| 13 | pptx / xlsx / docx **图表与 SmartArt 只画标题** | charts-pie.pptx ink 0.47、diagram.pptx 0.47、charts.xlsx 0.02、diagram.docx 0.12 | 3 | `HtmlWriter.cs:227-229`（`<figure class="chart">title</figure>`），幻灯片上的 graphicFrame 不在 `Slide()` 的三种情况里（`:49-85`） | PPTX-2 / XLSX-2 |
| 14 | pptx **几何、渐变、阴影、线宽**：非 rect/ellipse/roundRect 全是矩形，`line` 固定 1px | shapes-basic.pptx ink 0.63 | 3 | `HtmlWriter.cs:103-117` `ShapeStyle` | PPTX-2 |
| 15 | docx **图片只当行内 `<img>`**，浮动锚点、环绕、位置丢失 | pictures.docx SSIM 0.49 | 3 | `HtmlWriter.cs:211-216`；`DocxImage.cs` 不出锚点/环绕属性 | DOCX-1 |
| 16 | xlsx **共享字符串不复用**，每次写入追加一条；`t="str"`/内联串改成 `t="s"` | cell-formatting.xlsx：sst 3→163；`x:c@t` 值变 214 / 9 个文件（† 347 / 19） | 1（不可见，文件膨胀，diff 噪音） | `src/Writer.Formats/Xlsx/XlsxAdapter.cs:184-195` `XlsxStrings.Add`；`XlsxCells.cs:171-176` `SetString` | XLSX-2 |
| 17 | docx **run 被切碎并留下空 run**，每次保存都增长 | 13 个 docx 多出 `w:r` 4780 个；run-formatting.docx 78→176（49 个空） | 1 | `src/Writer.Formats/Docx/DocxReplace.cs:52-73` `Apply`、`:186-212` `SplitRun`；`DocxRuns.cs:133` `MakeRuns` | DOCX-2 |
| 18 | pptx **表格样式**（条纹、首行、`tblStyleId`）不渲染 | tables-styled.pptx ink 0.04 | 3 | `HtmlWriter.cs:259-`；`PptxNodes` 表格没有样式属性 | PPTX-2 |
| 19 | pptx 超链接 **tooltip 丢失**、`r:id` 重发 | † shapes-effects.pptx slide5：`a:hlinkClick@tooltip` −1 | 2 | `PptxText.cs:137-146` `SetLink` 新建 `HyperlinkOnClick` 只带 `Id` | PPTX-1 |
| 20 | 纯保存时的**包规范化**：`[Content_Types]` 把 `xml` 的 Default 改成主文档类型并删掉对应 Override、`.rels` Target 改成绝对路径、`model/gltf.binary`→`gltf-binary` | † academic-paper.docx 等 7 个 OfficeCLI 文件；仓库样本无 | 1（Office 都能开；只是 diff 噪音） | Open XML SDK 的包保存（`DocxAdapter.cs:67`、`PptxAdapter.cs:55` 调 `Save`） | 不修，工具已把这类零件的可见度记为 1 |

不算问题、也不进门禁的噪音：`w:ins/w:del@w:id` 重新编号（revisions.docx）、`x:c@t="n"` 被省略、`a:t@xml:space`。

## 5. 工作包

每个包的验收都是同一条命令：`node tools/fidelity/roundtrip.mjs` 退出码 0 且对应行的数字下降；改完后 `--update-baseline` 收紧基线。

### XLSX-1 公式（P0，先做）
- 读：从属单元格（`<f t="shared" si="n"/>`）要按 `si` 找到主公式并按偏移平移后报出 `formula`（或至少报出 `sharedWith` 让编辑器不写回常量）。`XlsxNodes.cs:225`。
- 写：`SetFormula` 里，如果新公式文本等于现有 `<f>` 的文本，什么都不动（保留 `t/ref/si` 和 `<v>`）；真的改主单元格时把整组拆成各自的普通公式再写。`XlsxCells.cs:198-209`。
- 缓存值：公式文本没变就保留 `<v>`；变了才清并置 `fullCalcOnLoad`。`XlsxCells.cs:204-205`。
- 建议把 † `budget-tracker.xlsx` 复制进 `tests/Writer.Tests/Fixtures/xlsx/`（同源 Apache-2.0），门禁才覆盖共享公式。

### XLSX-2 字符串与视图（P2）
- `XlsxStrings.Add` 先查已有条目再追加（`XlsxAdapter.cs:184-195`）；值没变时不改 `t`。
- `HtmlWriter` 的 sheet 分支用模型已有的 `fill/bold/italic/size/font/color/align/valign/wrap/format/merges/widths/heights` 画（`HtmlWriter.cs:222-227, 259-`），图表至少画成占位框而不是一行标题。

### PPTX-1 文本写回（P0/P1）
- `SetBody` 改成和 `DocxReplace` 一样的**按段落、按 run 的差异写回**：段落一一对应时保留各自的 `pPr` 和 `endParaRPr`；run 文本没变就原样保留 `rPr`，变了才从**相邻 run** 克隆 `rPr` 再叠加 `RunSpec` 的差异。`PptxText.cs:33-65`。这一项同时解决 #2、#3、#8 和 #9 的大半（`baseline/cap/spc/kern/lang` 跟着 `rPr` 克隆留下来）。
- `RunSpec`/`InlineHtml` 补 `sub`/`sup`（`baseline`、`w:vertAlign`）、`caps`、`letter-spacing`，编辑器才能表达上下标。`RunSpec.cs:8-10`、`InlineHtml.cs:150-180`。
- `SetLink` 保留 `tooltip` 等原有属性。`PptxText.cs:137-146`。

### PPTX-2 视图（P1）
- graphicFrame 图表 / SmartArt：至少按位置画占位框和标题（`HtmlWriter.cs:49-85`）。
- `ShapeStyle`：`prstGeom` 常见几何（triangle、diamond、arrows、star…）用 clip-path，`ln` 的宽度、渐变、阴影（`HtmlWriter.cs:103-117`）。
- 表格 `tblStyleId` 的首行/条纹（`HtmlWriter.cs:259-`）。
- 文本框 `bodyPr` 内边距（默认 0.1in/0.05in）和 `lnSpc` 行距，现在写死 `padding:6px 8px; line-height:1.2`（`HtmlWriter.cs:16-17`）。

### DOCX-1 版式（P1）
- 模型：段落补 `spacingBefore/After`、`lineSpacing`、`indent/firstLine/hanging`、`keepNext`；从样式链（docDefaults → 段落样式 → run 样式）解析出**生效**的字体、字号、间距，并把主题字体（`a:latin` majorFont/minorFont）翻译成实际字体名。`DocxNodes.cs:78-89`、`DocxStyles.cs`。
- 视图：`Render` 读 `sectPr`（纸张、页边距、分栏、页眉页脚）铺页面；标题和正文用解析出的样式而不是 h1–h6 默认值；表格用 `tblBorders/tcBorders/shd/tblW/gridCol`；图片用锚点定位和环绕。`HtmlWriter.cs:12, 26-41, 176-186, 211-216, 259-`。

### DOCX-2 run 卫生（P2）
- `Apply` 结束后合并相邻且 `rPr` 相同的 run，删除没有内容的 run。`DocxReplace.cs:52-73`。

## 6. 门禁怎么用

```bash
node tools/fidelity/roundtrip.mjs                 # 全部 fixtures，对照 tools/fidelity/baseline.json，退化则 exit 1
node tools/fidelity/roundtrip.mjs --only tables   # 只跑名字含 tables 的样本（不做门禁判断）
node tools/fidelity/roundtrip.mjs --render        # 加第一页渲染对比（Quick Look + headless Chrome，慢 3–4 倍）
node tools/fidelity/roundtrip.mjs --fixtures ../source/OfficeCLI/examples --fixtures ../source/OfficeCLI/assets/showcase   # 扩大样本
node tools/fidelity/roundtrip.mjs --update-baseline   # 修完一项后收紧基线
```

`baseline.json` 记每个样本 save/touch 的丢失总数和 touch 多出的元素数（`added`：切碎的 run、空 run，文件每存一次就长一点）；任一样本超过基线、或任一样本打不开，退出 1 并打印 `REGRESSION` 行。`--cap n` 调每文件的 touch 上限（默认 40），`--no-touch` 只跑 save，`--formats docx` 只跑一种格式。输出目录 `tools/fidelity/out/` 已在 `.gitignore`。
