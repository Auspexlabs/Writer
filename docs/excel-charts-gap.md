# Excel / WPS 表格 对比 Writer：图表与表格能力差距清单

> 历史快照：本文以恢复前代码为基线，XLSX 文件桥接与公式函数的描述已部分过时。

日期：2026-09-22　范围：Microsoft Excel（Microsoft 365 桌面版 + Excel 网页版）、WPS 表格（WPS Office / WPS 365，含 WPS AI）与本仓库的 Writer（引擎 `src/Writer.Formats/Xlsx` + 浏览器编辑器 `ui/SheetEditor.dc.html`）。

目标：用户要求“功能一定要全，百分百，甚至比他们还好用”。本清单按功能逐条列出三方现状，"我们现在"一列全部基于代码事实（备注列给出文件与函数），不做猜测；不确定处标 `?` 并汇总在第 6 节。

## 0. 口径

### 0.1 状态定义（"我们现在"列，每格以状态词开头，便于统计）

| 状态 | 含义 |
|---|---|
| 已有 | 编辑器可用且能落盘到 .xlsx（或本身是纯前端/纯引擎功能且完整） |
| 部分 | 有雏形：常见情况是 UI 能操作但不写入文件，或引擎能写但 UI 没有入口，或只覆盖子集 |
| 进行中 | 正在并行实现的波次（引擎图表读写、合并、列宽行高、冻结、自动筛选、超链接、完整单元格格式、约 150 个函数、功能区与桥接扩展、自由缩放） |
| 缺失 | 没有 |

### 0.2 Excel / WPS 列标记

`✓` 桌面与网页都有；`✓ / 网页✗` 仅桌面；`✓ / 网页仅查看` 网页能显示不能创建编辑；`会员` WPS 会员功能；`AI` 由 WPS AI / Copilot 提供；`?` 未能核实；`✗` 没有。

优先级：P0 = 缺了就不算"能用的电子表格"或直接损害文件保真；P1 = 高频办公功能；P2 = 进阶分析/排版；P3 = 长尾与生态。

### 0.3 备注列的文件缩写

| 缩写 | 文件 |
|---|---|
| XA | `src/Writer.Formats/Xlsx/XlsxAdapter.cs`（打开/新建/保存/工作表增删移） |
| XN | `src/Writer.Formats/Xlsx/XlsxNodes.cs`（document/sheet/row/cell/range 节点） |
| XC | `src/Writer.Formats/Xlsx/XlsxCells.cs`（值/类型/公式/CSV，`XlsxStyles` 样式） |
| REG | `src/Writer.Core/Registry.cs`（元素与属性注册表） |
| SE | `ui/SheetEditor.dc.html`（表格编辑器：功能区 `renderVals`、单元格模型、图表渲染） |
| CALC | `ui/sheet-engine.js`（浏览器公式引擎：`FN` 函数表、`fmt`、`shiftF`/`adjF`） |
| ENG | `ui/engine.js`（桥接：`openXlsx`/`saveXlsx`/`diffMark`） |
| IO | `ui/office-io.js`（浏览器侧 csv/xlsx 解析、`printDoc`、`exportCsv`） |
| IDX | `ui/index.dc.html`（应用壳：文件菜单、撤销栈、助手改动卡片） |
| CHAT | `src/Writer.Cli/Chat.cs`（助手系统提示） |
| TEST | `tests/Writer.Tests/XlsxTests.cs` 与 `tests/Writer.Tests/Fixtures/xlsx/*.xlsx` |

## 1. 我们现在的底座（一页速览）

- 文件层（引擎）：OpenXML SDK 打开/新建/保存；工作表增删改名移动；单元格 `value / type / formula / bold / color / fill / format`；行 `data`、区域 `values`（JSON 或 CSV）。保存是 `Package.Clone`，未触碰的部件（图表、透视表、切片器、迷你图、条件格式、数据验证、批注、图片、名称、VBA）字节级保留，10 个样例工作簿的往返测试覆盖（XA `Save`；TEST `Open_read_everything_save_changes_nothing`）。公式只存文本、置 `fullCalcOnLoad`，不算值（XC `SetFormula`，XA `RecalculateOnLoad`）。
- 编辑器（前端）：功能区 5 个页签 开始/插入/公式/数据/视图（SE `tabsDef`）；单元格模型 `s:{b,i,u,st,fill,color,align,va,fmt,dec,wrap,fs,bt,bb,bl,br}`；`colW`/`rowH`/`frR`/`frC`/`merges`/`charts`/`cf`/`filter`/`showF`/`noGrid`。网格固定 20 列 × 80 行（CALC `NC = 20, NR = 80`）。图表 3 种（柱/线/饼）以内联 SVG 渲染，可拖动、改标题、切换类型、删除。
- 桥接：打开时只读 值/公式/加粗/字色/填充/数字格式（ENG `cellModel`，`colW: {}`、`frR: 0` 硬编码）；保存是逐格差异（ENG `cellProps`），只写 值/公式/类型=文本/加粗/字色/填充/数字格式。列宽、行高、合并、冻结、图表、条件格式、筛选、斜体/下划线/删除线/字号/对齐/换行/边框均不落盘。
- 公式：浏览器引擎 46 个函数（CALC `FN`），支持跨表引用、`$` 绝对引用、通配符条件、错误传播、循环检测；不支持整列引用、数组、动态数组、名称。
- 助手：通过 `writer` 命令改文件（CHAT `SystemPrompt`：帮助文本来自 REG，所以助手只能做注册表里有的事：单元格值/公式/加粗/字色/填充/数字格式、工作表增删改）；改动以卡片列出，可一键撤销（ENG `diffMark`，IDX `settle`）。
- 导入导出：xlsx 读写；csv/tsv 前端解析打开；当前表导出 CSV；打印/PDF 走浏览器打印（IO `printDoc`，固定 A4 横向、无图表）；`view html/json`、`export --to md/docx/pptx/html/json`、其它格式的表格导出成工作簿。

## 2. 图表 checklist

### 2.1 图表类型（族与子类型）

Excel 类型清单来自微软《Office 中可用的图表类型》；WPS 类型来自 WPS 学院《表格如何插入图表》（柱状/折线/组合/饼形/条形/面积/散点/股价/雷达）与社区帖（树状图在新版"其他图表"中）。

| 功能 | Excel | WPS | 我们现在 | 差距 | 优先级 | 备注 |
|---|---|---|---|---|---|---|
| 簇状柱形图 | ✓ | ✓ | 部分：可插入、多系列并列渲染，不落盘 | 引擎读写进行中；无系列/轴设置 | P0 | SE `insertChart('bar')`、`chartVals` rects；ENG `saveXlsx` 不写 `charts` |
| 堆积柱形图 | ✓ | ✓ | 缺失 | | P1 | 渲染只需在 `chartVals` 累加 y0/y1；文件为 `c:grouping stacked` |
| 百分比堆积柱形图 | ✓ | ✓ | 缺失 | | P1 | `percentStacked` |
| 三维簇状/堆积/百分比堆积柱形图、三维柱形图 | ✓ / 网页仅查看 | ✓ | 缺失 | 需 3D 投影渲染 | P3 | 建议文件读写保留 `c:view3D`，前端 2.5D 近似 |
| 折线图 / 带数据标记的折线图 | ✓ | ✓ | 部分：折线 + 固定圆点标记，不落盘 | 无线型/标记选项 | P0 | SE `chartVals` lines/dots |
| 堆积折线图 / 带标记 | ✓ | ✓ | 缺失 | | P2 | |
| 百分比堆积折线图 / 带标记 | ✓ | ✓ | 缺失 | | P2 | |
| 三维折线图 | ✓ / 网页仅查看 | ✓ | 缺失 | | P3 | |
| 饼图 | ✓ | ✓ | 部分：渲染首个系列，>5% 扇区标百分比，不落盘 | 无爆炸/起始角/颜色 | P0 | SE `chartVals` slices/labels |
| 三维饼图 | ✓ / 网页仅查看 | ✓ | 缺失 | | P3 | |
| 复合饼图 / 复合条饼图 | ✓ | ✓ | 缺失 | | P2 | `ofPieChart` |
| 圆环图 | ✓ | ✓ | 进行中：引擎 doughnut；UI 无入口 | | P1 | 渲染=饼图挖孔 `holeSize` |
| 簇状条形图（横向） | ✓ | ✓ | 进行中：引擎 bar；UI 无横向 | UI 里 `'bar'` 实际是柱形 | P0 | SE `insertChart('bar')` 画竖条；引擎需区分 `barDir col/bar` |
| 堆积 / 百分比堆积条形图 | ✓ | ✓ | 缺失 | | P1 | |
| 三维条形图各子类 | ✓ / 网页仅查看 | ✓ | 缺失 | | P3 | |
| 面积图 | ✓ | ✓ | 进行中：引擎 area；UI 无入口 | | P1 | |
| 堆积面积 / 百分比堆积面积 | ✓ | ✓ | 缺失 | | P1 | |
| 三维面积图各子类 | ✓ / 网页仅查看 | ✓ | 缺失 | | P3 | |
| 散点图（仅标记） | ✓ | ✓ | 进行中：引擎 scatter；UI 无入口 | UI 目前没有数值型 X 轴 | P1 | SE `chartVals` 的 X 轴是等距类别轴 |
| 散点图 带平滑线（含/不含标记） | ✓ | ✓ | 缺失 | | P2 | |
| 散点图 带直线（含/不含标记） | ✓ | ✓ | 缺失 | | P2 | |
| 气泡图 / 三维气泡图 | ✓ | ✓ | 缺失 | | P2 | `bubbleChart` 第三列为气泡大小 |
| 股价图（盘高-盘低-收盘 / 开盘-盘高-盘低-收盘 / 成交量-盘高-盘低-收盘 / 成交量-开盘-盘高-盘低-收盘） | ✓ / 网页? | ✓ | 缺失 | | P3 | |
| 曲面图（三维曲面 / 三维线框 / 曲面俯视 / 线框俯视） | ✓ / 网页仅查看 | ✗? | 缺失 | | P3 | WPS 学院类型列表无曲面图 |
| 雷达图 / 带标记 / 填充雷达 | ✓ | ✓ | 缺失 | | P2 | |
| 树状图 | ✓ | ✓（新版"其他图表"） | 缺失 | chartEx（`cx:`）新格式 | P2 | OfficeCLI `Core/Chart/ChartExBuilder.cs` 可参考 |
| 旭日图 | ✓ | ? | 缺失 | chartEx | P2 | |
| 直方图 / 排列图（Pareto） | ✓ | ?（有帕累托教程，直方图未核实） | 缺失 | chartEx；需分箱 binCount/binSize/上下溢出 | P2 | |
| 箱形图 | ✓ | ? | 缺失 | chartEx；四分位法 inclusive/exclusive | P2 | |
| 瀑布图 | ✓ | ?（旧版需堆积柱模拟） | 缺失 | chartEx；增减/汇总点 | P1 | 财务场景高频 |
| 漏斗图 | ✓ | ? | 缺失 | chartEx | P2 | |
| 地图（着色地图） | ✓（M365 桌面/Mac/网页） | ✗ | 缺失 | 需地理数据服务 | P3 | |
| 组合图（簇状柱-折线 / 次坐标轴 / 堆积面积-簇状柱 / 自定义每系列类型） | ✓ / 网页部分 | ✓ | 缺失 | | P1 | 文件层是同一 `plotArea` 内多 chart 元素 |
| 迷你图（折线 / 柱形 / 盈亏，高低首尾负点标记，样式） | ✓ / 网页仅查看 | ✓ | 缺失 | 文件层保留 | P2 | fixtures `sparklines.xlsx` 往返保留 |
| 数据透视图 | ✓ | ✓ | 缺失 | 依赖透视表 | P2 | |
| 图表工作表（Chart sheet） | ✓ / 网页? | ✓? | 缺失 | | P3 | |
| 动态图表动画（Office 2024） | ✓ 桌面 | ✗ | 缺失 | | P3 | |

### 2.2 图表元素与格式

| 功能 | Excel | WPS | 我们现在 | 差距 | 优先级 | 备注 |
|---|---|---|---|---|---|---|
| 图表标题（显示/隐藏、文本、位置、覆盖式、引用单元格 `=Sheet1!A1`） | ✓ | ✓ | 部分：有标题，双击/菜单改文本，不落盘 | 不可隐藏/定位/引用单元格 | P0 | SE `chartVals.onRename`；引擎 titles 进行中 |
| 坐标轴显示/隐藏（主横/主纵） | ✓ | ✓ | 部分：固定显示数值刻度与类别标签 | 不可隐藏 | P1 | SE `chartVals` labels |
| 坐标轴标题 | ✓ | ✓ | 缺失 | | P1 | |
| 坐标轴格式（最小/最大/主次单位、对数、逆序、交叉点、刻度线、标签位置、数字格式、日期轴/文本轴） | ✓ | ✓ | 部分：自动"好看"刻度（1/2/2.5/5 步长 ≤5 格），万/亿缩写 | 无任何手动设置 | P1 | SE `chartVals` stepv、`short` |
| 网格线（主/次，横/纵） | ✓ | ✓ | 部分：固定主横网格线 | 不可切换 | P1 | |
| 图例（位置上/下/左/右/覆盖，显示/隐藏，编辑图例项） | ✓ | ✓ | 部分：固定底部图例 | 引擎 legend 进行中 | P1 | SE 模板 legend flex 区 |
| 数据标签（值/类别/系列名/百分比/引导线/位置/来自单元格的值） | ✓ | ✓ | 部分：仅饼图百分比 | | P1 | |
| 数据表（含/不含图例项标示） | ✓ | ✓ | 缺失 | | P2 | |
| 趋势线（线性/指数/对数/多项式/幂/移动平均，前推后推，显示公式与 R²） | ✓ | ✓ | 缺失 | | P2 | |
| 误差线（标准误差/百分比/标准偏差/固定值/自定义） | ✓ | ✓ | 缺失 | | P2 | |
| 线条：垂直线 / 高低点连线 / 系列线；涨跌柱线 | ✓ | ✓ | 缺失 | | P3 | |
| 次坐标轴（把系列绘制在次轴） | ✓ | ✓ | 缺失 | | P1 | |
| 系列格式（填充/边框/间隙宽度/系列重叠/平滑线/标记形状大小/负值反转/饼图爆炸与起始角/圆环内径/气泡大小） | ✓ | ✓ | 部分：固定 6 色调色板、固定间隙 | 无系列级设置 | P1 | SE `PAL`，`chartVals` bw = gw*0.7 |
| 数据点单独格式（按点着色 varyColors、单点填充） | ✓ | ✓ | 缺失 | | P2 | |
| 图表区/绘图区格式（填充、边框、阴影、发光、圆角） | ✓ | ✓ | 部分：固定白底圆角卡片 | | P2 | SE 模板 chart div |
| 图表内字体（标题/轴/图例字号字色） | ✓ | ✓ | 缺失：固定 11px | | P2 | |
| 更改颜色（主题配色，单色/彩色） | ✓ | ✓ | 缺失 | | P1 | |
| 图表样式（预设样式库） | ✓ / 网页部分 | ✓ | 缺失 | | P1 | |
| 快速布局 | ✓ / 网页? | ✓ | 缺失 | | P2 | |
| 切换行/列 | ✓ | ✓ | 缺失：插入时自动判定表头行/标签列，之后不能切换 | | P1 | SE `insertChart` hdr/lab 推断 |
| 选择数据（系列增删改、类别范围、系列顺序、隐藏与空单元格处理） | ✓ | ✓ | 缺失：插入后数据源不可改 | | P0 | SE chart 对象 `dr1/dr2/cols/lab` 只在 `insDel` 时位移 |
| 筛选图表数据（图表筛选器勾选系列/类别） | ✓ / 网页✗ | ✗? | 缺失 | | P2 | |
| 更改图表类型（同族/跨族、按系列） | ✓ | ✓ | 部分：柱/线/饼三态切换 | | P0 | SE `chartVals.types` |
| 移动图表（拖动） | ✓ | ✓ | 已有：拖标题栏移动 | 位置不落盘（引擎 position 进行中） | P0 | SE `onDrag`/`onWM`/`onWU` |
| 调整大小 / 锁定纵横比 | ✓ | ✓ | 缺失：固定 460×300 | | P0 | SE `insertChart` w/h |
| 属性：随单元格移动和调整大小 / 不随 | ✓ | ✓ | 部分：插入删除行列时数据源跟随，位置为像素坐标不跟随 | | P2 | SE `insDel` charts 段 |
| 标题/系列名自动取表头 | ✓ | ✓ | 已有 | | — | SE `insertChart` title、`chartVals` series.name |
| 空单元格与隐藏行列的显示方式（空距/零/连线） | ✓ | ✓ | 部分：非数字按 0；筛选隐藏行仍计入 | | P2 | SE `chartVals` vals |
| 替代文字（无障碍） | ✓ | ✓ | 缺失 | | P3 | |
| 图表模板（另存为 .crtx / 应用模板） | ✓ / 网页✗ | ✓ | 缺失 | | P3 | |
| 推荐的图表 | ✓ / 网页✗ | AI 推荐 | 缺失 | AI 机会（第 4 节） | P1 | |
| 快速分析（选区右下角一键图表/迷你图/格式） | ✓ 桌面 | ✗ | 缺失 | | P2 | |
| 三维旋转 / 透视 | ✓ 桌面 | ✓ | 缺失 | | P3 | |
| 图表复制为图片 / 另存为图片 | ✓ | ✓ | 缺失 | | P2 | SVG 已在 DOM，导出 PNG 成本低 |
| 图表复制到 Word / PPT（嵌入或链接） | ✓ | ✓ | 缺失 | 跨格式导出机会 | P1 | 引擎已能写 pptx/docx |
| 图表打印 / 单独打印图表 | ✓ | ✓ | 缺失：打印输出不含图表 | | P2 | IO `printDoc` xlsx 分支只输出表格 |
| 图表字体随文档主题 | ✓ | ✓ | 缺失 | | P3 | |
| 图表列表/导航 | ✗（选择窗格） | ✗ | 已有：侧栏图表列表，点击定位 | 比他们好 | — | SE `chartList` |
| AI 侧栏改图表 | Copilot（需授权） | AI | 部分：助手能改单元格，不能建图/改图 | 注册表无 chart 元素 | P0 | CHAT `SystemPrompt` 只暴露 REG 里的元素 |

### 2.3 图表的文件往返

| 功能 | Excel | WPS | 我们现在 | 差距 | 优先级 | 备注 |
|---|---|---|---|---|---|---|
| 读取文件中已有图表并显示 | ✓ | ✓ | 进行中：引擎读；桥接与 UI 还不读 | | P0 | ENG `openXlsx` 不读 charts |
| 新建图表写入文件（column/bar/line/pie/area/scatter/doughnut） | ✓ | ✓ | 进行中 | | P0 | |
| 修改已有图表（数据源/标题/图例/位置） | ✓ | ✓ | 进行中 | | P0 | |
| 删除文件中的图表（含 drawing 关系与部件清理） | ✓ | ✓ | 缺失：UI 删除仅内存 | | P0 | SE `onDelete` |
| 未修改图表原样保留（含 chartEx、3D、模板样式） | — | — | 已有：`Package.Clone` | | — | XA `Save`；TEST 覆盖 `charts.xlsx`、`charts-pie.xlsx` |
| chartEx 新图表（树状/旭日/直方/箱形/瀑布/漏斗）读写 | ✓ | 部分 | 缺失 | | P2 | OfficeCLI `ChartExBuilder`、`chartex-style.xml` 可参考 |
| 图表缓存值（`c:numCache`/`c:strCache`）写入 | ✓ | ✓ | 进行中：是否随图表写入一并落盘未核实 | 不写缓存则非 Excel 阅读器显示空图 | P1 | 见第 6 节 |
| 图表锚定（twoCellAnchor 随单元格 / oneCellAnchor / absolute） | ✓ | ✓ | 进行中（position） | | P1 | |

### 2.4 WPS 特有

| 功能 | Excel | WPS | 我们现在 | 差距 | 优先级 | 备注 |
|---|---|---|---|---|---|---|
| 在线图表 / 稻壳图表模板库 | ✗ | 会员 | 缺失 | 用"样式预设 + AI"替代 | P2 | |
| AI 生成图表（右键"图表"即推荐、WPS AI 按钮洞察、对话生成、自然语言调整图表） | Copilot 类似 | AI | 部分：助手可对话但无图表工具 | | P0 | 第 4 节机会 1、7 |
| 图表美化（预设模板一键切换） | 图表样式 | ✓ | 缺失 | | P1 | |
| 图表以嵌入工作簿方式粘贴到文字/演示，可直接编辑数据源 | 嵌入对象 | ✓ | 缺失 | | P1 | 第 4 节机会 4 |
| 多维表格统计图表（实时关联数据库式表） | ✗ | ✓ | 缺失 | 非传统电子表格，不做 | P3 | |

## 3. 表格能力 checklist

### 3.1 单元格与格式

| 功能 | Excel | WPS | 我们现在 | 差距 | 优先级 | 备注 |
|---|---|---|---|---|---|---|
| 字体族 | ✓ | ✓ | 进行中：引擎 font；UI 无字体下拉 | | P0 | XC 只读写 Bold/Color；REG cell 无 font |
| 字号 | ✓ | ✓ | 部分：UI 9–36 下拉（`s.fs`），不落盘 | 引擎 size 进行中 | P0 | SE `S(cs.fs …)`；ENG `cellProps` 不写 fs |
| 加粗 | ✓ | ✓ | 已有：UI + 文件 | | — | XN `SetProp("bold")`；ENG `cellProps.bold` |
| 斜体 / 下划线 / 删除线 | ✓ | ✓ | 部分：UI 有（`s.i/u/st`），不落盘 | 引擎进行中 | P0 | SE `tb('i'/'u'/'st')` |
| 双下划线 / 上标 / 下标 | ✓ | ✓ | 缺失 | | P3 | |
| 字体颜色（调色板/主题色/自定义） | ✓ | ✓ | 已有：UI 取色器 + 文件 RGB | 读取不识别主题色/索引色（读为默认色） | — | SE `C('A' …)`；XC `XlsxStyles.Rgb` 仅 rgb |
| 填充色（纯色） | ✓ | ✓ | 已有：UI 7 预设 + 自定义 hex；文件 solid fill | 无主题色 | — | SE `FILLS`；XN `SetProp("fill")` |
| 图案填充 / 渐变填充 | ✓ | ✓ | 缺失 | | P3 | |
| 边框（13 种线型、颜色、上下左右/内部/外部/斜线、绘制边框） | ✓ | ✓ | 部分：UI 6 预设（所有/外侧/上/下/粗下/无）黑色，不落盘 | 引擎 borders 进行中；无颜色/线型/内部/斜线 | P0 | SE `borders()`；`s.bt/bb/bl/br` |
| 水平对齐（左/中/右/填充/两端/跨列居中/分散） | ✓ | ✓ | 部分：左/中/右 UI，不落盘 | 引擎 align 进行中 | P0 | SE `AL()` |
| 垂直对齐（上/中/下） | ✓ | ✓ | 部分：UI 不落盘 | 引擎 valign 进行中 | P0 | SE `M('va' …)` |
| 自动换行 | ✓ | ✓ | 部分：UI 不落盘 | 引擎 wrap 进行中 | P0 | SE `tb('wrap')` |
| 缩进（增/减） | ✓ | ✓ | 进行中：引擎 indent；UI 无 | | P1 | |
| 文字方向 / 旋转角度 / 竖排 | ✓ | ✓ | 缺失 | | P2 | |
| 缩小字体填充 | ✓ | ✓ | 缺失 | | P3 | |
| 合并（合并居中 / 跨越合并 / 合并单元格 / 取消） | ✓ | ✓ | 部分：UI 合并居中与取消，不落盘 | 引擎 merges 进行中；无跨越合并 | P0 | SE `merge()`；IO `xlsxSheets` 浏览器直读路径读 `mergeCell` 但 ENG `openXlsx` 不读 |
| 数字格式类别（常规/数值/货币/会计/日期/时间/百分比/分数/科学记数/文本/特殊/自定义代码） | ✓ | ✓ | 部分：UI 6 类（常规/数字/货币¥/百分比/日期/文本）+ 小数位；文件层能写任意格式代码 | UI↔文件映射有损 | P0 | ENG `fmtOf`/`codeOf`：`0.00`→`#,##0.00`（多出千位符）、任何日期码→`yyyy-mm-dd`、货币恒为 `"¥"`、会计/分数/科学/时间/特殊→常规；XC `NumberFormatId` 支持内置 + 自定义 ≥164 |
| 千位分隔 / 增减小数位 / 货币 / 百分比 按钮 | ✓ | ✓ | 已有 | | — | SE `decimals()` |
| 条件格式：突出显示（大于/小于/介于/等于/文本包含/发生日期/重复值） | ✓ | ✓ | 部分：大于/小于/等于/文本包含/重复值，固定红底，不落盘 | 无介于/日期；无格式选择 | P1 | SE `addCF`、`cfs` 渲染；REG 无 cf |
| 条件格式：最前/最后（前 N / 前 N% / 高于低于平均） | ✓ | ✓ | 缺失 | | P1 | |
| 条件格式：数据条（渐变/实心/负值/坐标轴） | ✓ | ✓ | 部分：单色渐变，不落盘 | | P1 | |
| 条件格式：色阶（双色/三色） | ✓ | ✓ | 部分：三色固定，不落盘 | | P1 | |
| 条件格式：图标集 | ✓ | ✓ | 缺失 | | P1 | |
| 条件格式：公式规则、规则管理器（优先级/如果为真则停止/应用范围） | ✓ | ✓ | 缺失：仅清除所选/全部 | | P1 | |
| 条件格式落盘 / 读取文件规则 | ✓ | ✓ | 缺失：文件层原样保留 | | P1 | fixtures `conditional-formatting.xlsx` 往返保留 |
| 单元格样式（内置样式库 / 新建 / 合并） | ✓ | ✓ | 缺失 | | P2 | |
| 主题（颜色/字体/效果） | ✓ | ✓ | 缺失 | 主题色单元格读为无色 | P2 | XC `XlsxStyles.Rgb` |
| 格式刷（单击 / 双击锁定） | ✓ | ✓ | 部分：单次 | | — | SE `painter` |
| 清除（内容/格式/全部/批注/超链接） | ✓ | ✓ | 部分：内容/格式/全部 | | — | SE `clear()` |
| 富文本（单元格内多段字体） | ✓ | ✓ | 缺失：读取合并为纯文本；写回丢失 runs | | P2 | XC `Display` 用 `InnerText` |
| 超链接（URL/本文档位置/邮件/文件，屏幕提示，Ctrl+K） | ✓ | ✓ | 进行中：引擎；UI 无 | | P1 | |
| 批注 / 注释 | ✓ | ✓ | 缺失：文件层保留 | | P1 | 见 3.6 |
| 单元格内图片 / `IMAGE` 函数 | ✓ M365 | ✗? | 缺失 | | P3 | |
| 复选框（单元格控件） | ✓ M365 | ? | 缺失 | | P3 | |
| 数据类型（股票/地理/组织） | ✓ M365 | ✗ | 缺失 | | P3 | |
| 特殊符号 / 公式（方程） | ✓ | ✓ | 缺失 | | P3 | |
| 单元格锁定 / 隐藏公式（配合保护） | ✓ | ✓ | 缺失 | | P2 | |
| 选择性粘贴（数值/格式/公式/转置/运算/跳过空单元/列宽） | ✓ / 网页部分 | ✓ | 缺失：内部剪贴板整体粘贴，外部只收 TSV 值 | | P1 | SE `paste()` |
| 拖放移动/复制单元格 | ✓ | ✓ | 缺失 | | P2 | |
| 填充柄（数值序列/复制/公式相对引用/日期与工作日序列/自定义序列/填充选项菜单/双击填充到底） | ✓ | ✓ | 部分：线性数值序列、公式偏移、文本尾数递增、Ctrl+D/R | 无日期序列、双击填充、选项菜单 | P1 | SE `doFill`/`fillDir` |
| 自动完成（同列文本） | ✓ | ✓ | 缺失 | | P2 | |
| 输入自动识别（%、¥、千分位） | ✓ | ✓ | 已有 | 无日期/时间/分数识别 | P1 | SE `parseInput` |
| 单元格内换行（Alt+Enter） | ✓ | ✓ | 已有 | | — | SE `onInKey` |
| 中文输入法组合 | ✓ | ✓ | 已有 | | — | SE `isComposing` |
| 右键上下文菜单（单元格） | ✓ | ✓ | 缺失：只有工作表标签有右键菜单 | | P1 | SE `onMD` 对 `button===2` 直接返回 |

### 3.2 行列与工作表

| 功能 | Excel | WPS | 我们现在 | 差距 | 优先级 | 备注 |
|---|---|---|---|---|---|---|
| 网格大小 | 16,384 列 × 1,048,576 行 | 同 Excel | 部分：固定 20 列 × 80 行；文件里超出的单元格会读入但不显示、不可编辑、原样保留 | 硬伤 | P0 | CALC `NC = 20, NR = 80`；SE 所有循环与 `setSel` 夹取用 `E.NC/E.NR`；需虚拟滚动 |
| 列宽（拖动 / 精确值 / 自动调整 / 默认宽） | ✓ | ✓ | 部分：拖动、双击自适应；不落盘 | 引擎进行中；无精确输入 | P0 | SE `onResize`/`onAuto`；ENG `openXlsx` 写死 `colW: {}` |
| 行高（拖动 / 精确值 / 自动） | ✓ | ✓ | 部分：拖动；不落盘 | 引擎进行中 | P0 | SE `rowH`、`onResize` |
| 插入/删除 整行整列（多行列、公式引用跟随、跨表引用跟随） | ✓ | ✓ | 已有：UI；落盘为逐格差异 | 引擎无"插入行"语义，大表会发很多命令 | P1 | SE `insDel`；CALC `adjF`；ENG `saveXlsx` |
| 插入/删除 单元格（右移/下移） | ✓ | ✓ | 缺失 | | P2 | |
| 隐藏 / 取消隐藏 行列 | ✓ | ✓ | 缺失：只有筛选造成的隐藏 | | P1 | SE `hidden` 仅来自 `filter` |
| 冻结窗格（首行 / 首列 / 任意位置） | ✓ | ✓ | 部分：仅首行/首列开关，不落盘 | 引擎进行中；任意位置缺 | P0 | SE `frR/frC` 取值 0/1 |
| 拆分窗口 | ✓ / 网页✗ | ✓ | 缺失 | | P3 | |
| 分组 / 分级显示 / 自动建立分级 / 显示级别 | ✓ | ✓ | 缺失 | | P2 | |
| 网格线显示、行号列标显示 | ✓ | ✓ | 部分：网格线开关不落盘 | | P2 | SE `noGrid` |
| 工作表 新建/删除/重命名/移动/复制 | ✓ | ✓ | 部分：全有；重命名会同步跨表公式；但左移右移的顺序不写回文件 | 顺序不落盘 | P1 | SE `tabctx`、`renameSheet`；ENG `saveXlsx` 无 move；XN `MoveTo` 引擎已支持 |
| 跨工作簿移动/复制工作表 | ✓ | ✓ | 缺失 | | P3 | |
| 工作表标签颜色 | ✓ | ✓ | 缺失 | | P2 | |
| 隐藏 / 取消隐藏工作表（含 veryHidden） | ✓ | ✓ | 缺失 | | P2 | |
| 工作表保护（密码、允许的操作列表） | ✓ / 网页仅生效 | ✓ | 缺失 | | P2 | |
| 工作簿保护（结构 / 窗口） | ✓ | ✓ | 缺失 | | P3 | |
| 允许编辑区域 / 区域权限 | ✓ | ✓ | 缺失 | | P3 | |
| 工作表背景图片 | ✓ 桌面 | ✓ | 缺失 | | P3 | |
| 从右到左工作表 | ✓ | ✓ | 缺失 | | P3 | |
| 名称框跳转 / 定位（Ctrl+G）/ 定位条件 | ✓ | ✓ | 部分：名称框输入 A1 或区域跳转 | 无定位条件 | P2 | SE `onNameKey` |
| 缩放（10–400% 自由、缩放到选区） | ✓ | ✓ | 部分：50–200% 7 档 | 自由缩放进行中 | P1 | SE 状态栏 `<select>` 与视图页签 4 档 |
| 新建窗口 / 并排查看 / 同步滚动 | ✓ 桌面 | ✓ | 缺失 | | P3 | |
| 工作表概览侧栏（使用范围、单元格数、图表数） | ✗ | ✗ | 已有 | 比他们好 | — | SE `sheetThumbs` |
| 全选、整行整列选择、Shift/Ctrl 扩展 | ✓ | ✓ | 部分：全选、点表头选整列整行、Shift 扩展 | 无 Ctrl 多区域选择、无 Ctrl+Space/Shift+Space | P1 | SE `colHeads.onMD`、`selAll` |
| 键盘导航（方向键、Ctrl+方向、Enter/Tab、F2、Home、PageUp/Down、Delete、Esc） | ✓ | ✓ | 已有 | 无 Ctrl+Home/End、Ctrl+Shift+End | P2 | SE `onInKey`、`jump` |
| 状态栏（就绪/编辑、缩放、聚合） | ✓ | ✓ | 已有 | | — | SE `statText` |

### 3.3 公式与函数

#### 3.3.1 函数分类覆盖

Excel 数量取自微软《Excel 函数（按类别）》页面（含兼容性函数约 500 个，微软口径"400 多个"）；WPS 与 Excel 基本同名同集（新版含 XLOOKUP/FILTER/UNIQUE 等动态数组函数）。我们 = CALC `FN` 表，46 个。

| 类别 | Excel 数量 | WPS | 我们现在 | 差距（示例） | 优先级 | 备注 |
|---|---|---|---|---|---|---|
| 数学与三角 | 84 | ✓ | 部分：13（SUM PRODUCT ROUND ROUNDUP ROUNDDOWN INT ABS MOD POWER SQRT PI RAND SUMIF） | SUMIFS SUMPRODUCT SUBTOTAL AGGREGATE CEILING FLOOR MROUND TRUNC RANDBETWEEN SEQUENCE QUOTIENT SIGN EXP LN LOG 三角函数 | P0 | ~150 函数扩展进行中 |
| 统计 | 99 | ✓ | 部分：9（AVERAGE MIN MAX COUNT COUNTA COUNTBLANK MEDIAN COUNTIF AVERAGEIF） | COUNTIFS AVERAGEIFS MAXIFS MINIFS LARGE SMALL RANK.EQ STDEV.S VAR.S MODE PERCENTILE QUARTILE CORREL FORECAST.* TREND FREQUENCY | P0 | |
| 逻辑 | 15 | ✓ | 部分：5（IF IFERROR AND OR NOT） | IFS SWITCH IFNA XOR TRUE FALSE LET LAMBDA MAP BYROW BYCOL REDUCE SCAN MAKEARRAY | P0 | |
| 查找与引用 | 35 | ✓（新版） | 部分：4（VLOOKUP HLOOKUP INDEX MATCH） | XLOOKUP XMATCH FILTER SORT SORTBY UNIQUE OFFSET INDIRECT ROW COLUMN ROWS COLUMNS CHOOSE LOOKUP TRANSPOSE HYPERLINK ADDRESS TAKE DROP VSTACK HSTACK GROUPBY PIVOTBY | P0 | |
| 文本 | 39 | ✓ | 部分：10（LEN LEFT RIGHT MID UPPER LOWER TRIM CONCAT CONCATENATE TEXT） | TEXTJOIN TEXTSPLIT TEXTBEFORE TEXTAFTER SUBSTITUTE REPLACE FIND SEARCH REPT PROPER VALUE EXACT CHAR CODE CLEAN NUMBERVALUE REGEXTEST/EXTRACT/REPLACE | P0 | `TEXT` 仅支持 `%`/`,`/小数位三种格式 |
| 日期与时间 | 21 | ✓ | 部分：2（TODAY NOW，且返回字符串） | DATE YEAR MONTH DAY HOUR MINUTE SECOND EDATE EOMONTH DATEDIF NETWORKDAYS WORKDAY WEEKDAY WEEKNUM DAYS YEARFRAC DATEVALUE TIME | P0 | CALC `TODAY` 返回 `yyyy/m/d` 字符串，无法参与日期运算 |
| 信息 | 18 | ✓ | 部分：3（ISBLANK ISNUMBER ISERROR） | ISTEXT ISNA ISERR ISEVEN ISODD ISFORMULA NA N TYPE CELL SHEET | P1 | |
| 财务 | 66 | ✓ | 缺失 | PMT FV PV NPV IRR XNPV XIRR RATE NPER IPMT PPMT SLN DDB | P1 | |
| 数据库 | 12 | ✓ | 缺失 | DSUM DCOUNT DAVERAGE DGET | P2 | |
| 工程 | 54 | ✓ | 缺失 | CONVERT DEC2BIN HEX2DEC BITAND | P3 | |
| 多维数据集 | 7 | ✗ | 缺失 | | P3 | 依赖数据模型，不做 |
| Web | 3 | ✗? | 缺失 | ENCODEURL FILTERXML WEBSERVICE | P3 | |
| 兼容性 | 41 | ✓ | 部分：1（CONCATENATE） | RANK STDEV VAR PERCENTILE QUARTILE MODE FORECAST | P2 | 旧名别名映射即可 |
| 用户自定义 | 3 | ✗ | 缺失 | | P3 | |
| 合计 | ≈500 | — | 部分：46 | | | CALC `FUNCS` |

#### 3.3.2 公式能力

| 功能 | Excel | WPS | 我们现在 | 差距 | 优先级 | 备注 |
|---|---|---|---|---|---|---|
| 公式存入文件（任意函数） | ✓ | ✓ | 已有：存 `<f>`，置 fullCalcOnLoad，Excel 打开重算 | 引擎不写缓存值 `<v>`，非 Excel 阅读器看到空 | — | XC `SetFormula`；XA `RecalculateOnLoad` |
| 前端实时计算 | ✓ | ✓ | 部分：46 函数 | | P0 | CALC `Calc` |
| 动态数组 / 溢出（FILTER SORT UNIQUE SEQUENCE、`#` 溢出引用、#SPILL!） | ✓ | ✓（新版） | 缺失：区域当标量时取首格（隐式交集） | | P1 | CALC `sc()` |
| 传统数组公式（Ctrl+Shift+Enter）、数组常量 `{1,2;3,4}` | ✓ | ✓ | 缺失：词法不识别 `{}` | | P1 | CALC `TK` |
| 整列 / 整行引用（A:A、1:1） | ✓ | ✓ | 缺失：只识别 A1 形式 | | P0 | CALC `TK` 第三分支 |
| 跨表引用（Sheet!A1、'带 空格'!A1、中文表名） | ✓ | ✓ | 已有 | | — | CALC `refNode` |
| 跨工作簿外部引用 `[Book.xlsx]Sheet!A1` | ✓ / 网页仅查看 | ✓ | 缺失 | | P3 | |
| 三维引用 `Sheet1:Sheet3!A1` | ✓ 桌面 | ✓ | 缺失 | | P3 | |
| 结构化引用 `Table1[列]` | ✓ | ✓ | 缺失 | 依赖表格对象 | P2 | |
| 名称管理器 / 定义名称 / 根据所选内容创建 / 用于公式 | ✓ / 网页仅使用 | ✓ | 缺失：文件层 definedNames 原样保留 | | P1 | |
| 运算符：算术 / 比较 / 文本连接 / 百分号 / 幂 | ✓ | ✓ | 已有 | | — | CALC `parse` |
| 运算符：引用联合 `,`、交集 空格、隐式交集 `@` | ✓ | ✓ | 缺失 | | P3 | |
| 绝对/相对/混合引用（$）与 F4 切换 | ✓ | ✓ | 部分：解析并在复制/填充/插入删除时保持 | 无 F4 | P1 | CALC `shiftF` d1/d2 |
| 公式自动补全 / 参数提示 / 函数说明 | ✓ | ✓ | 部分：fx 菜单按分组插入函数名 | 无参数提示 | P1 | SE `fnMenu`、`insertFn`；CALC `FUNC_GROUPS` |
| 引用着色（编辑公式时引用区域彩色框） | ✓ | ✓ | 缺失 | | P1 | |
| 点选引用（编辑公式时点击单元格插入引用） | ✓ | ✓ | 已有 | 不能拖选区域 | — | SE `onMD` 判断 `cur.val` 末尾运算符 |
| 自动求和（求和/平均/计数/最大/最小，自动推断区域，多列一次） | ✓ | ✓ | 已有 | | — | SE `autoFn` |
| 显示公式（Ctrl+`） | ✓ | ✓ | 已有：不落盘 | | — | SE `showF` |
| 公式审核：追踪引用 / 从属单元格、移去箭头 | ✓ / 网页✗ | ✓ | 缺失 | | P2 | |
| 公式求值（逐步计算） | ✓ 桌面 | ✓ | 缺失 | AI 机会 2 | P2 | |
| 错误检查（错误指示器、原因、修复建议） | ✓ 桌面 | ✓ | 部分：错误值红字显示 | 无提示/修复 | P1 | SE `color = isErr ? …` |
| 监视窗口 | ✓ 桌面 | ✓? | 缺失 | | P3 | |
| 循环引用提示 / 迭代计算（最多次数、最大误差） | ✓ | ✓ | 部分：循环得 `#CIRC!`（Excel 是 0 + 状态栏警告与定位）；无迭代 | | P1 | CALC `value()` stack |
| 计算选项（自动/手动/除数据表外、F9 重算） | ✓ | ✓ | 缺失：始终自动 | | P2 | |
| 精度（15 位有效数字）与显示精度 | ✓ | ✓ | 部分：JS double，常规显示取 10 位有效 | | P2 | CALC `fmt` `toPrecision(10)` |
| 日期序列值体系（1900 系统、日期加减、时间小数） | ✓ | ✓ | 部分：`fmt('date')` 按序列显示；但文件读入的日期变成文本 `2024-01-05`，TODAY/NOW 返回字符串，日期运算不可用 | | P0 | XC `Display` 输出 `yyyy-MM-dd` 文本；ENG `cellModel` 取显示文本；CALC `TODAY` |
| 文本型数字隐式转换、布尔参与运算 | ✓ | ✓ | 已有 | | — | CALC `num` |
| 错误类型（#DIV/0! #N/A #NAME? #NULL! #NUM! #REF! #VALUE! #SPILL! #CALC!） | ✓ | ✓ | 部分：前 7 种缺 #NULL!，多自定义 #CIRC! | | P2 | CALC `fail` |
| 通配符条件（`*` `?`、比较运算符文本） | ✓ | ✓ | 已有 | | — | CALC `crit` |
| LAMBDA / LET 自定义函数 | ✓ | ✓? | 缺失 | | P2 | |
| Python in Excel / Office.js 自定义函数 / VBA UDF / WPS JS 宏 | ✓ 桌面 | JS 宏 | 缺失 | 助手替代 | P3 | |
| 公式栏（展开多行、名称框、fx） | ✓ | ✓ | 部分：单行公式栏 + 名称框 + fx | | P2 | SE 模板第 52–55 行 |
| 公式大写自动化、逗号分隔 | ✓ | ✓ | 已有 | | — | CALC `upperF` |

### 3.4 数据工具

| 功能 | Excel | WPS | 我们现在 | 差距 | 优先级 | 备注 |
|---|---|---|---|---|---|---|
| 排序：单列升/降序（自动识别表头） | ✓ | ✓ | 已有 | | — | SE `sort`、`hasHeader` |
| 排序：多列多级、按颜色/图标、自定义序列、按行、区分大小写 | ✓ | ✓ | 缺失 | | P1 | |
| 自动筛选：按值勾选、（全选）、空白 | ✓ | ✓ | 部分：UI 有，不落盘 | 引擎 autofilter 进行中 | P0 | SE `toggleFilter`、`flt` 弹层 |
| 筛选：文本/数字/日期条件、前 10 项、按颜色、搜索框 | ✓ | ✓ | 缺失 | | P1 | |
| 高级筛选（条件区域 / 复制到 / 唯一记录） | ✓ 桌面 | ✓ | 缺失 | | P2 | |
| 重新应用 / 清除筛选 | ✓ | ✓ | 部分：清除 | | — | SE `清除筛选` |
| 删除重复项（可选列） | ✓ | ✓ | 部分：整行全列判重，不可选列 | | P1 | SE `dedupe` |
| 分列（分隔符 / 固定宽度 / 列数据格式） | ✓ / 网页✗ | ✓（含智能分列） | 缺失 | | P1 | |
| 快速填充（Flash Fill） | ✓ 桌面 | ✓ 智能填充 Ctrl+E | 缺失 | AI 机会 | P1 | |
| 数据验证（整数/小数/序列/日期/时间/文本长度/自定义；输入信息；出错警告；圈释无效数据） | ✓ | ✓ 有效性 | 缺失：文件层保留 | | P1 | fixtures `data-validation.xlsx` 往返保留 |
| 合并计算（按位置/按分类） | ✓ 桌面 | ✓ | 缺失 | | P2 | |
| 模拟分析：单变量求解 | ✓ / 网页✗ | ✓ | 缺失 | | P2 | |
| 模拟分析：模拟运算表、方案管理器 | ✓ 桌面 | ✓ | 缺失 | | P3 | |
| 规划求解 | ✓ 桌面加载项 | ✓ | 缺失 | | P3 | |
| 预测工作表（FORECAST.ETS） | ✓ 桌面 | ✗? | 缺失 | | P3 | |
| 分析工具库（回归、t 检验、直方图…） | ✓ 桌面加载项 | ✗ | 缺失 | | P3 | |
| 数据透视表（创建、字段列表、布局、值汇总/显示方式、分组、刷新、样式、计算字段） | ✓ | ✓ | 缺失：文件层保留 | | P2 | fixtures `pivot-tables.xlsx`；OfficeCLI 有实现参考 |
| 数据透视图 | ✓ | ✓ | 缺失 | | P2 | |
| 推荐的数据透视表 | ✓ 桌面 | ✗ | 缺失 | AI 机会 | P2 | |
| 切片器（表 / 透视表） | ✓ / 网页仅使用 | ✓ | 缺失：文件层保留 | | P2 | fixtures `slicers.xlsx` |
| 日程表（时间线筛选） | ✓ | ✗? | 缺失 | | P3 | |
| 表格对象（套用表格格式、表样式、汇总行、镶边、筛选按钮、自动扩展、转换为区域） | ✓ | ✓ | 缺失：文件层保留 | | P1 | |
| 获取数据 / Power Query（CSV/文本/Excel/JSON/XML/网页/数据库/API，查询编辑器，刷新） | ✓ 桌面 / 网页部分 | 部分（导入文本/数据库） | 缺失 | | P3 | |
| 现有连接 / 全部刷新 | ✓ 桌面 | ✓ | 缺失 | | P3 | |
| Power Pivot / 数据模型 / DAX | ✓ 桌面 | ✗ | 缺失 | | P3 | 不做 |
| 分类汇总 | ✓ 桌面 | ✓ | 缺失 | | P2 | |
| 从图片插入数据（OCR） | ✓ M365 | 会员 | 缺失 | | P3 | |
| 快速分析（选区一键格式/图表/汇总） | ✓ 桌面 | ✗ | 缺失 | | P2 | |
| 数据分析问答（Ideas / Analyze Data） | ✓ M365 | AI | 部分：助手可读区域后回答 | 无引用高亮、无图表输出 | P1 | CHAT 提示只含 view/get；第 4 节机会 6 |
| 查找与替换（区分大小写/全字匹配/公式或值/按格式/全部查找列表/跨工作簿） | ✓ / 网页无替换 | ✓ | 部分：子串查找（值与公式文本）、逐个/全部替换（不改公式） | 无选项、仅当前表 | P1 | SE `findNext`/`replace` |
| 定位条件（常量/公式/空值/可见单元格/行列差异） | ✓ 桌面 | ✓ | 缺失 | | P2 | |
| 状态栏聚合（求和/平均/计数/数值计数/最大/最小，可定制） | ✓ | ✓ | 部分：平均/计数/求和 | | — | SE `statText` |
| 转置粘贴 | ✓ | ✓ | 缺失 | | P1 | |
| 文本转数字 / 数字转文本 | 提示修复 | ✓ | 缺失 | 一键清理机会 | P1 | |
| 删除空行空列、批量去空格 | ✗（手工） | 智能工具箱 | 缺失 | 一键清理机会 | P1 | |
| 复制/剪切/粘贴（内部含格式与公式偏移；系统剪贴板 TSV） | ✓ | ✓ | 已有 | 外部粘贴无 HTML 表格 | — | SE `copy`/`paste`、`onCopy`/`onPaste` |
| 撤销 / 重做（多级） | ✓ | ✓ | 已有：80 步 | | — | IDX `hist` |

### 3.5 视图与打印

| 功能 | Excel | WPS | 我们现在 | 差距 | 优先级 | 备注 |
|---|---|---|---|---|---|---|
| 普通 / 分页预览 / 页面布局视图 | ✓ / 网页? | ✓ | 缺失 | | P2 | |
| 打印区域（设置/取消/添加） | ✓ | ✓ | 缺失 | | P1 | |
| 打印标题（顶端标题行 / 左端标题列） | ✓ 桌面 | ✓ | 缺失 | | P1 | |
| 页面设置（纸张/方向/边距/居中/缩放到 N 页） | ✓ | ✓ | 缺失：固定 A4 横向 12mm | | P1 | IO `printDoc` xlsx 分支 |
| 页眉页脚（预设/自定义/页码/日期/文件名/图片；首页与奇偶页不同） | ✓ / 网页✗ | ✓ | 缺失 | 文件层保留 | P1 | |
| 分页符（插入/删除/重置） | ✓ 桌面 | ✓ | 缺失 | | P2 | |
| 打印选项（网格线、行号列标、批注、错误值、单色、草稿） | ✓ | ✓ | 缺失 | | P2 | |
| 打印范围（选定区域 / 活动表 / 整个工作簿、份数、先列后行） | ✓ | ✓ | 部分：打印所有表的使用范围 | | P1 | IO `printDoc` |
| 打印预览 | ✓ | ✓ | 部分：浏览器打印对话框 | | — | |
| 导出 PDF（引擎级，含页面设置与图表） | ✓ | ✓ | 部分：浏览器"打印为 PDF"，无图表 | | P1 | IDX `download('pdf')` |
| 自定义视图 | ✓ 桌面 | ✓ | 缺失 | | P3 | |
| 工作表视图（Sheet Views，协作个人筛选排序） | ✓ | ✗ | 缺失 | | P3 | |
| 显示/隐藏 编辑栏 / 标题 / 网格线 | ✓ | ✓ | 部分：网格线 | | P3 | |
| 阅读模式（十字高亮） | ✗ | ✓ | 缺失 | | P2 | |
| 护眼模式 / 深色模式 | ✗ / 深色 UI | ✓ | 缺失 | | P3 | |
| 缩放到选定区域 | ✓ | ✓ | 缺失 | | P3 | |
| 全屏 | ✓ | ✓ | 缺失 | | P3 | |

### 3.6 协作与审阅

| 功能 | Excel | WPS | 我们现在 | 差距 | 优先级 | 备注 |
|---|---|---|---|---|---|---|
| 批注（线程评论、@提及、回复、解决） | ✓ | ✓ 评论 | 缺失：文件层保留 | | P1 | |
| 注释（旧式 Note，显示/隐藏、批量显示） | ✓ | ✓ 批注 | 缺失 | | P2 | |
| 修订 / 跟踪更改（旧式共享工作簿） | ✓ 桌面 | ✓ 修订 | 缺失 | | P3 | |
| 显示更改（Show Changes 面板，按人/范围/时间） | ✓ | 历史版本 | 部分：助手改动卡片列出修改单元格数、可撤销/保留（只覆盖 AI 改动） | | P1 | ENG `diffMark`；IDX `settle` |
| 版本历史 / 恢复 | ✓ 云 | ✓ 云 | 缺失 | | P2 | |
| 实时共同编辑（多人光标、在线状态） | ✓ | ✓ 金山文档 | 缺失：单用户；引擎有文件 change SSE 可做刷新 | | P3 | docs/engine.md `/events` |
| 共享链接与权限（查看/编辑/指定人/有效期/水印/下载限制） | ✓ | ✓ | 缺失 | | P3 | |
| 文件加密（打开密码 / 修改密码） | ✓ 桌面 | ✓ | 缺失 | | P3 | |
| 标记为最终 / 数字签名 / 敏感度标签 / IRM | ✓ 桌面 | 部分 | 缺失 | | P3 | |
| 拼写检查 / 翻译 / 朗读 / 辅助功能检查器 | ✓ | ✓ | 缺失 | | P3 | |
| 工作簿统计 / 文档检查器 | ✓ | ✓ | 部分：侧栏统计 | | — | SE `sheetThumbs` |
| 自动保存 | ✓ 云 | ✓ | 已有：每次修改约 1 秒内写回引擎，文件保持未建模内容 | | — | docs/engine.md；ENG `save` |
| 助手改动隔离（差异卡、一键撤销、保留） | Copilot 无差异卡 | WPS AI 可撤销 | 已有 | 比他们好的基础 | — | IDX `settle`；SE `x.ai` 高亮 |

### 3.7 导入导出

| 功能 | Excel | WPS | 我们现在 | 差距 | 优先级 | 备注 |
|---|---|---|---|---|---|---|
| 打开 / 保存 .xlsx | ✓ | ✓ | 已有：OpenXML；未触碰部件字节保真 | | — | XA；TEST 10 个样例 |
| .xls（97-2003） | ✓ | ✓ | 缺失 | | P3 | |
| .xlsm（宏保留） | ✓ | ✓ | 缺失：扩展名不在适配器列表 | | P3 | XA `Extensions = [".xlsx"]` |
| .xlsb / .ods / .et / .csv 直接打开 | ✓ / ✓ / ✗ / ✓ | ✓ | 缺失 | | P3 | |
| CSV / TSV 导入（编码、分隔符、类型推断） | ✓ | ✓ | 部分：前端解析 csv/tsv 为表格；引擎 `range values` 接受 CSV 文本 | 无编码/分隔符选项；引擎不能直接打开 csv | P1 | IO `csvCells`；XC `ParseCsv`；ENG `upload` |
| CSV / TSV / TXT 导出 | ✓ | ✓ | 部分：当前表导出 CSV（UTF-8 BOM） | 无 TSV、无编码选择、无全簿 | P1 | IO `exportCsv` |
| PDF 导出 | ✓ | ✓ | 部分：浏览器打印 | 见 3.5 | P1 | |
| 图片导出（区域 / 图表 另存为图片） | 复制为图片 | ✓ 输出为图片 | 缺失 | | P2 | |
| HTML / 网页导出 | ✓ 桌面 | ✓ | 部分：引擎 `view html` / `export --to html`（表头 + 单元格显示值） | 样式是否随出未核实 | P2 | `HtmlWriter` case "sheet" |
| JSON / Markdown 导出 | ✗ | ✗ | 已有：`view json`、`export --to md` | 比他们多 | — | `Exporter`、`Views` |
| 从 docx/md/pptx 表格生成工作簿 | ✗ | ✗ | 已有：`export --to x.xlsx` | 比他们多 | — | `Exporter.Workbook` |
| 从网页 / 数据库 / JSON / XML 导入 | ✓ 桌面 | 部分 | 缺失 | | P3 | |
| 粘贴 HTML 表格（网页/Word 表格保留格式） | ✓ | ✓ | 部分：只读 `text/plain` TSV | | P1 | SE `onPaste` |
| 插入图片（浮动） | ✓ | ✓ | 缺失：文件层保留 | | P2 | |
| 形状 / 文本框 / SmartArt / 图标 | ✓ | ✓ | 缺失：保留 | | P3 | fixtures `shapes.xlsx` |
| 嵌入对象（OLE） | ✓ 桌面 | ✓ | 缺失：保留 | | P3 | |
| 宏 VBA / Office 脚本 / JS 宏 | ✓ 桌面 / Office Scripts | JS 宏 | 缺失：VBA 部件保留不执行 | 助手替代 | P3 | |
| 加载项 / 插件 | ✓ | ✓ | 缺失 | | P3 | |
| CLI / MCP / HTTP 编程接口 | ✗（Office.js/Graph） | ✗ | 已有 | 比他们多 | — | docs/engine.md |

### 3.8 WPS 特有便利功能

| 功能 | Excel | WPS | 我们现在 | 差距 | 优先级 | 备注 |
|---|---|---|---|---|---|---|
| 智能工具箱（60 项：填充序列、录入日期、填充空白、合并拆分、一键对比、表格目录、工作表排序重命名、大小写、加减乘除、多区域粘贴、保留/去除内容） | ✗ | 会员 | 缺失 | 助手一键化 | P2 | WPS 社区帖 |
| 智能填充 Ctrl+E（模式识别拆分/合并列） | 快速填充 | ✓ | 缺失 | | P1 | |
| 数据对比（两区域/两表差异标记） | ✗（Inquire 加载项） | ✓ | 缺失 | AI 机会 | P2 | |
| 拆分表格（按内容拆成多表/多簿）/ 合并表格（多表/多簿合并） | ✗ | ✓ | 缺失 | | P2 | |
| 高亮重复项 / 拒绝录入重复项 | 条件格式 / 数据验证 | ✓ | 部分：条件格式重复值 | | P1 | SE `addCF('dup')` |
| 阅读模式（十字高亮） | ✗ | ✓ | 缺失 | | P2 | |
| 智能分列 | 分列 | ✓ | 缺失 | | P1 | |
| 表格目录 / 工作表批量重命名与排序 | ✗ | ✓ | 缺失 | | P3 | |
| 输出为图片 / 输出为 PDF（一键） | 部分 | ✓ | 部分 | | P2 | |
| 图片转表格（OCR） | ✓ | 会员 | 缺失 | | P3 | |
| WPS AI：AI 写公式（自然语言 → 公式，双击 Ctrl 唤起） | Copilot | AI | 部分：助手可写 `formula`，无预览/校验 | | P0 | CHAT；第 4 节机会 2 |
| WPS AI：公式解读 | Copilot | AI | 部分：助手可读公式回答 | | P1 | |
| WPS AI：数据分析（概览/趋势/排名/异常） | Copilot | AI | 部分 | | P1 | |
| WPS AI：生成图表（推荐 + 一键 + 文字摘要） | Copilot | AI | 缺失 | | P0 | |
| WPS AI：一键生成报告 | ✗ | AI | 缺失 | 跨格式导出机会 | P2 | |
| 云同步 / 手机平板端 | ✓ | ✓ | 缺失：本地文件夹 | | P3 | |
| 多维表格（数据库式） / 智能表单 | Lists / Forms | ✓ | 缺失 | 不做 | P3 | |

## 4. 比他们更好用的机会

每条一行规格，标注层：引擎 / 前端 / 两者。

1. 自然语言建图与改图（两者）：注册表新增 `chart` 元素（属性 `type/data/categories/series/title/legend/dataLabels/axisTitles/secondaryAxis/anchor/size`），助手系统提示自动带上 `help xlsx chart`；用户说"把 B2:D9 做成按月折线图放在 F2"→ `add /sheet[1] --type chart …`，"改成堆积柱形并加数据标签"→ `set /sheet[1]/chart[1] …`；改动进入差异卡可撤销。
2. 公式解释与生成带预览（两者）：选中单元格→"解释"：助手读取公式与引用值，逐步说明；"帮我写公式"：先写到预览态（前端临时单元格显示结果），用户确认后 `set --prop formula=`；错误值（#N/A、#REF!）单元格出现"修复"提示并给出替代写法（XLOOKUP 精确匹配、IFERROR）。
3. 一键清理（引擎 + 前端按钮）：新增语义命令 `clean <file> <sheet>`：去首尾空格、文本型数字转数字、统一日期格式、删空行空列、去重、表头加粗冻结、列宽自适应；执行前给出变更计数预览，执行后进差异卡。
4. 跨格式导出（引擎）：`export <xlsx> --range A1:D9 --to deck.pptx --slide 3` 把区域/图表以原生表格或 DrawingML 图表写进 pptx/docx（引擎已能写这两种格式，图表部件同源）；反向 `export --to report.docx` 生成带图表的报告。
5. 助手改动的结构级差异与逐项撤销（前端）：`diffMark` 从"N 个单元格"扩到"新增图表 X / 合并 A1:C1 / 插入 2 行 / 新增条件格式规则"，每项可单独撤销；同时保留现有整体撤销。
6. 数据问答带引用高亮（两者）：助手回答"第三季度哪个区域最高"时返回引用的单元格路径，前端用 `ai` 高亮并可"写成公式"落到表里，把一次性答案变成可复算的公式。
7. 推荐图表带理由（前端 + 助手）：选区右键或插入页签显示 3 个候选图表缩略图与一句理由（"3 个系列随时间变化→折线"），点一下插入，后续用对话微调。
8. 自然语言条件格式与数据验证（两者）：注册表加 `cf` 与 `validation` 元素；"低于 60 分标红""这列只能选是/否"直接落成文件规则，Excel/WPS 打开一致。
9. 跨文件、跨格式的一次操作（引擎已有基础）：同一个助手能把 sales.xlsx 汇总表贴进 报告.docx、再生成 汇报.pptx，Excel/WPS 都做不到"一个对话跨三种文件"。
10. 批量生成（助手 + 引擎 range/sheet 命令）："给每个地区建一张表并各画一张图"，助手循环执行；提供 `add --type sheet --from-template` 减少命令数。
11. 变更审计时间线（前端）：把助手每条命令（路径 + 属性）记成时间线，可回放到任意点，比 Excel 的"显示更改"多了"为什么改"（对话上下文）。
12. 公式引擎与 Excel 对齐的可验证性（前端）：为 CALC 建一套 Excel 结果对照用例（fixtures 里的公式表），保证同一公式在我们与 Excel 结果一致；这是"比他们好用"的信任基础。

## 5. 实施路线图（进行中波次之后）

| 波次 | 内容 | 层 | 工作量 | 解锁的清单项 |
|---|---|---|---|---|
| 进行中 | 引擎读写图表（column/bar/line/pie/area/scatter/doughnut、标题、系列、图例、位置）、合并、列宽行高、冻结、自动筛选、超链接、完整单元格格式、约 150 函数；功能区与桥接跟进；自由缩放 | 两者 | 大 | 2.1 基础族、2.3、3.1 多数 P0、3.2 列宽行高冻结 |
| 第 1 波：能当真表格用（P0 收尾） | 网格去掉 20×80 上限（虚拟滚动、按需渲染、`NC/NR` 动态）；日期以序列值贯穿（XC Display 输出序列 + 格式、CALC 日期函数 20 个）；数字格式代码原样往返（UI 显示格式代码、`fmtOf/codeOf` 不再改写）；整列整行引用；图表：选择数据、切换行列、调整大小、删除落盘；工作表顺序/标签颜色/隐藏落盘；引擎"插入删除行列"语义命令；`XlsxStyles.Apply` 去重 cellXfs | 两者 | 大 | 3.2 网格、3.3 日期/整列、3.1 数字格式、2.2 选择数据/大小 |
| 第 2 波：高频办公（P1） | 条件格式全类型读写 + 规则管理器；数据验证；批注/注释；超链接 UI；表格对象（表样式、汇总行、结构化引用）；多列排序、条件筛选；分列、快速填充；选择性粘贴/转置/HTML 粘贴；查找替换选项；隐藏行列；引用着色与参数提示；名称管理器；图表元素：轴标题、数据标签、图例位置、网格线开关、更改颜色/样式、次坐标轴、组合图、堆积/百分比堆积、圆环/面积/散点/气泡/雷达 UI、瀑布图；打印区域/标题/页面设置/页眉页脚 + 引擎级 PDF；富文本单元格；动态数组与数组公式；财务/信息函数；右键菜单；机会 1、2、3、5、8 | 两者 | 大 | 2.1/2.2 P1、3.1 条件格式、3.4 多数 P1、3.5 打印、3.6 批注 |
| 第 3 波：进阶分析（P2） | 数据透视表 + 透视图 + 切片器（引擎读写与 UI）；迷你图；分组分级显示与分类汇总；合并计算、单变量求解；公式审核（追踪、求值、监视）；单元格样式与主题；工作表/工作簿保护与区域权限；趋势线、误差线、数据表、单点格式、chartEx 六类图表；图片插入；分页预览；版本历史；图表/区域导出图片；机会 4、6、7、9、10 | 两者 | 大 | 2.1 chartEx、2.2 P2、3.4 透视、3.5 视图 |
| 第 4 波：长尾与生态（P3） | 股价/曲面/地图/3D 图表读写与近似渲染；工程/多维/Web 函数；模拟运算表、方案管理器、规划求解、预测表；Power Query 式获取数据与刷新；实时协作与共享权限；.xls/.xlsm/.ods；加密与签名；拆分窗口、多窗口、自定义视图；机会 11、12 | 两者 | 大 | 其余 P3 |

工作量口径：小 = 一天内一层可完成；中 = 数天、一层为主；大 = 一周以上或两层联动。上表各波次内部可拆小项，单项估计：图表元素每项 小~中；条件格式/数据验证读写 各 中；透视表 大；虚拟滚动网格 中~大；日期序列化 中；数字格式往返 小；引擎 PDF 中~大。

## 6. 统计与不确定之处

### 6.1 状态统计（按"我们现在"列首词计数）

| 范围 | 已有 | 部分 | 进行中 | 缺失 | 合计 |
|---|---|---|---|---|---|
| 图表（2.1–2.4） | 4 | 16 | 9 | 60 | 89 |
| 表格（3.1–3.8） | 29 | 63 | 3 | 127 | 222 |
| 合计 | 33 | 79 | 12 | 187 | 311 |

说明："部分"里有 14 行（图表 2、表格 12）注明"引擎进行中"（UI 已有、落盘等进行中波次），进行中波次完成后它们会升为"已有"；"缺失"里 71 行（图表 16、表格 55）是 P3 长尾，扣除后 P0–P2 的缺失为 116 行。统计脚本见 6.3。

### 6.2 未能完全核实的点

- 我们的：图表缓存值（`c:numCache`）是否随进行中的引擎图表写入一起落盘——若不写，除 Excel 外的阅读器（含 WPS 的部分预览、macOS 快速查看）会显示空图。
- 我们的：`export --to html` / `view html` 对工作表是否输出加粗/填充等样式，只看到 `HtmlWriter` 按 `row.data` 输出显示值。
- 我们的：浏览器直读路径 `IO.xlsxSheets`（读列宽与合并）与引擎路径 `ENG.openXlsx`（不读）两条路径何时各自生效；上传走 `ENG.upload` 后一律经引擎打开，因此实际用户路径应是不读列宽/合并。
- Excel 网页版：股价图、曲面图、气泡图能否在网页版新建；快速布局在网页版是否存在。微软页面只明确"3D 图表网页版仅查看""推荐图表与高级格式仅桌面"。
- WPS：旭日图、直方图、箱形图、瀑布图、漏斗图是否已成为原生图表类型（社区 2024 帖证实树状图在"其他图表"中；官方 2025 文章提到旭日/瀑布/漏斗但未列为插入项）；曲面图大概率没有；日程表、监视窗口、复选框控件未核实。
- WPS：动态数组与 LET/LAMBDA 的支持范围（新版已有 XLOOKUP/FILTER/UNIQUE 类函数）。

### 6.3 统计脚本

```bash
# 在仓库根目录运行；分别统计图表段（2.1–2.4）与表格段（3.1–3.8）"我们现在"列的状态词
python3 - <<'EOF'
import re, collections
sec = None; cnt = collections.Counter(); rows = collections.Counter()
for line in open('docs/excel-charts-gap.md', encoding='utf-8'):
    if line.startswith('## 2.'): sec = 'chart'
    elif line.startswith('## 3.'): sec = 'sheet'
    elif line.startswith('## 4.'): sec = None
    if not sec or not line.startswith('| '): continue
    cells = [c.strip() for c in line.strip().strip('|').split('|')]
    if len(cells) < 7 or cells[0] in ('功能', '类别', '状态', '缩写') or cells[0].startswith('---'): continue
    k = re.split(r'[：:（ ]', cells[3])[0]
    if k in ('已有', '部分', '进行中', '缺失'): cnt[(sec, k)] += 1; rows[sec] += 1
for s in ('chart', 'sheet'): print(s, {k: cnt[(s, k)] for k in ('已有', '部分', '进行中', '缺失')}, 'total', rows[s])
EOF
```

## 来源

- Microsoft 支持《Office 中可用的图表类型》 https://support.microsoft.com/en-us/office/available-chart-types-in-office-a6187218-807e-4103-9e0a-27cdb19afb90
- Microsoft 支持《在浏览器中和在 Excel 中使用工作簿的区别》 https://support.microsoft.com/en-us/office/differences-between-using-a-workbook-in-the-browser-and-in-excel-f0dc28ed-b85d-4e1d-be6d-5878005db3b6
- Microsoft Learn《Excel 网页版服务说明》 https://learn.microsoft.com/en-us/office365/servicedescriptions/office-online-service-description/excel-online
- Microsoft 支持《Excel 函数（按类别）》 https://support.microsoft.com/en-us/office/excel-functions-by-category-5f91f4e9-7b42-46d2-9bd1-63f26a86c0eb
- WPS 学院《表格如何插入图表和设计图表》 https://www.wps.cn/learning/course/detail/id/541.html
- WPS 官网《WPS AI 生成图表实操指南》 https://www.wps.cn/article/ban-gong-xiao-lv-shu-ju-ke-shi-hua-2026-NK6lasDc.html
- WPS 365 博客《WPS AI 表格功能升级：写公式+公式解读+数据分析+图表四合一》 https://plus.wps.cn/blog/p109403.html
- WPS 社区《智能工具箱》 https://bbs.wps.cn/topic/10418 ；《矩形树状图》 https://bbs.wps.cn/topic/38772 ；《堆积瀑布图》 https://bbs.wps.cn/topic/17262
- WPS 365 博客《分页预览》 https://plus.wps.cn/blog/p95215.html ；《切片器常见问题》 https://plus.wps.cn/blog/p118992.html ；《数据透视图常见问题》 https://plus.wps.cn/blog/p119477.html ；《打印设置全攻略》 https://plus.wps.cn/blog/p117459.html
- 知乎《关于 WPS 图表的创建、类型及构成》 https://zhuanlan.zhihu.com/p/716387068
- 本仓库参考实现：`source/OfficeCLI/src/officecli/Core/Chart/*`、`Handlers/Excel/*` 与 `tests/Writer.Tests/Fixtures/md/excel_*.md`（OOXML 级图表/条件格式/验证/透视/切片器/迷你图属性词汇）
