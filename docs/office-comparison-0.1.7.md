# Writer 0.1.7 与 Word、Excel、PowerPoint 的深入对比

**实施进度更新：用户已要求全部补齐并加入 PDF 离线。已落地项与剩余范围见 [实施状态](office-implementation-status.md)。本文以下是修复前审计基线，不能把历史“未修复”当成最新状态，也不能把已完成的一批修复当成全部完成。**

核查日期：2026-09-27。代码基线：`f44087cdfa8fbcd6399eba286b1b43a562f4b10f`，加本次 Excel 默认对齐校验修复。对比对象以 Microsoft 365 桌面版为主；Windows、Mac、网页版和订阅服务的能力有差异，不能混为一套。尤其 Power Pivot、部分插件及云端能力不能当作所有 Mac Office 都有的功能。

**结论：基础编辑已经比较广，仍有明显的复杂文档兼容、计算与数据分析、专业排版、演示制作差距。当前不能称为与 Office 功能相同。** 安装体积小，既来自架构和系统组件复用，也来自尚未覆盖的功能；体积本身不能证明功能相当或运行内存更低。

本文核对了编辑器、保存桥接层、原生文件读写及微软官方功能说明。另运行实际逻辑探针与自动化回归。没有将所有项目在真实 Microsoft Office 和桌面 WebView 中逐项人工验收；“有实现”不等于所有文件都兼容。

**第二轮补充：已用临时原生文件复现 1904 日期系统的公式缓存写错，另确认 Excel 打印样式缺失、PPT PNG 富文本与图片效果降级。** 具体使用场景和可执行验收条件见第 9–11 节。这轮继续做对比与复现，没有修改这些问题的产品实现，也没有更新安装包。

**第三轮补充：日常操作层面的缺口见第 12 节。** 新核对了制表位排版、书签精度、交叉引用选项、表格公式、公式审核、定位条件、工作表成组编辑、拆分窗格、对象选择与锁定、全稿字体替换；也确认表格排序／文本表格转换／图片抠图等已有实现，避免误列为缺失。

## 1. 如何理解支持程度

| 标记 | 含义 |
|---|---|
| 已有 | 找到编辑和落盘实现；复杂文件的外观仍可能需要样本验证 |
| 部分 | 有功能入口，但范围、计算、显示或输出有限制 |
| 保留／显示 | 文件对象可保留或显示；不能据此称为可完整编辑 |
| 未实现 | 当前源码中没有找到对应的完整使用流程 |
| 待验证 | 不能仅靠源码或现有测试证明正确，需真实文件或实际应用验证 |

优先级：**P0** 保存失败／可能丢数据；**P1** 日常使用的正确性、排版或核心业务阻塞；**P2** 高级功能；**P3** 长尾功能或另一个产品方向。优先级是工程建议，不是承诺的发布日期。本文不使用“达到 Office 的百分之多少”，因为没有合理的统一分母。

## 2. 此次实际复现的问题

| 问题 | 实际结果 | 当前状态 |
|---|---|---|
| Excel 合并居中后撤销，保存报 `align: 'general' is not valid` | 前端撤销发送合法的 Excel 默认对齐值，通用属性校验却拒绝它 | **本次本地修复**：仅对 XLSX 的单元格／区域允许 `general`；原生写入器会清除显式对齐，恢复自动对齐。已验证保存重开。安装版尚未因此自动更新 |
| Excel 时间验证／自定义公式验证 | 已有规则设为 09:00–17:00，输入 08:00 未被拒绝；规则 `A1>0`，输入 -3 未被拒绝 | **未修复，P1**。规则保留在模型中，编辑器 `dvBad` 明确跳过这两类 |
| Excel 条件格式打印 | A1=20、规则 >10 填充 `#123456`，编辑器显示该色，打印 HTML 不含该色 | **未修复，P1**。打印没有使用条件格式求值结果；不只是打印机颜色设置问题 |
| Excel 超大显式范围求和 | `SUM(A1:A1048576)` 返回 `#REF!`；同一文档 `SUM(A1:A10)` 正常 | **未修复，P1**。地址支持到最后一行，但范围求值有百万格上限；`COUNTBLANK(A:A)` 等专门处理的函数正常 |
| Excel 1904 日期系统公式 | A1 为 `2024-03-05`，B1=`A1+1`，日期格式；保存重开 B1 为 `2028-03-07`，应为 `2024-03-06`。1900 系统对照样本正确 | **未修复，P0**。原始 A1 能正确读取；错误出在计算与缓存写入，不能用已有普通日期读写测试证明日期公式正确 |
| Excel 普通打印样式 | 字符串形式的四边细线／双线在打印中变成无边框；删除线、文字旋转、格式码颜色未进入输出；工作簿默认 Georgia 18 被打印为 Arial 11 | **未修复，P1**。逐边形式的双线对照样本可以输出，说明是不同样式路径处理不一致 |
| Word 多栏分页 | 1 栏启用分页，2、3 栏不启用同一分页模式 | **部分支持，P1**。保存分栏设置不等于编辑器能准确逐页显示多栏文档 |
| Word 页眉表格 | 原生页眉同时放普通段落与表格；返回给编辑器的 HTML 只有段落。修改页眉段落后，表格仍在 DOCX 中 | **保留但不显示，P1**。不能当作已丢数据，也不能当作编辑／打印支持完整 |
| PPT 动画名称与实际播放效果 | `dissolve` 采用淡入淡出帧；不认识的进入动画也回退为淡入淡出；Morph 的纯旋转变化没有对应旋转帧 | **部分支持，P1/P2**。原文件保留与播放效果完整是两件事 |
| PPT 导出 PNG | 同一文本框里 64px 红色粗体片段被与其余文字一起按 32px 普通深色绘制；图片裁剪、灰度没有用于 Canvas 绘制 | **未修复，P1**。调用实际 PNG 绘制函数并记录 Canvas 命令确认；尚未做真实像素截图比对。背景图片对照样本有正常绘制调用，不应误报全部图片都丢失 |

复现脚本：[office-0.1.7-probes.mjs](audits/office-0.1.7-probes.mjs)，结果：[office-0.1.7-probes.json](audits/office-0.1.7-probes.json)。这是 Node 下运行实际产品逻辑的结果，不是真实 Office 的截图对比。

第二轮脚本：[office-0.1.7-workflows.mjs](audits/office-0.1.7-workflows.mjs)，结果：[office-0.1.7-workflows.json](audits/office-0.1.7-workflows.json)。使用真实 CLI 读写临时 DOCX/XLSX，并记录实际打印函数和 PNG 绘制函数的输出；不修改用户文件。

### 截图问题修复与验证

- 根因：`ui/engine.js:2194` 撤销格式时发出 `align=general`，`src/Writer.Core/Registry.cs` 的 XLSX 格式属性却只允许 left/center/right/justify。
- 修复：XLSX 单独声明对齐枚举，加入 general。Word/PPT 的对齐约束保持各自语义；没有把 general 强行映射成 left，因为数字的自动对齐应仍为右对齐。
- 原生回归：单格及区域经过实际属性校验，设置多种格式、恢复默认、保存、重开、再次居中。修复前两项测试均复现同一异常，修复后通过。
- 新增跨层回归 `ui/tests/xlsx-history.test.mjs`：真实编辑器逻辑 → 保存桥接 → 本地引擎 → `.xlsx` 文件 → 重开。覆盖合并居中、跨列合并、取消合并、撤销、重做，以及批量格式恢复。检查文字、公式、区域外单元格及格式，不只是检查命令生成。
- 本次完整回归：原生 **930/930**，UI／桥接 **1074/1074**，均无跳过；CLI 构建无警告、无错误。测试通过不代表下面列出的未实现能力已经补齐。

## 3. Word：主要缺口是复杂排版与专业文档工作流

已有基础：字体与中西文字体、段落和自定义样式、多级编号、表格合并拆分、图片与环绕裁剪、形状、公式、脚注尾注、目录题注、批注与回复、常见修订、撤销、水印、页面色、简单图表、主题、基本邮件合并及比较报告。不能再沿用旧清单，把这些全部写成“没有”。微软对桌面版和网页版的区别见[官方功能表](https://support.microsoft.com/en-us/word/word-features-comparison-word-for-the-web-vs-desktop)。

| 能力 | 我们现在及欠缺 | 使用影响／优先级 | 主要代码证据 |
|---|---|---|---|
| 分页与多栏 | 单栏有分页及局部重排；多栏不走同一分页模式 | 报刊、双栏论文的页数和布局不能按 Word 精度承诺，P1 | `ui/WordEditor.dc.html:1148` |
| 分节混合纸张 | 原生分节可保存；编辑器与打印仍有全局页面尺寸路径 | 横竖混排、不同纸张的长文件须专项验收，P1 | `ui/word-print.js:5-42`、`DocxSection.cs` |
| 跨页大表／复杂脚注 | 基础表格与脚注已有；完整跨页排版尚未验证 | 合同、论文、报表打印最容易暴露差异，P1 | `docs/releases/0.1.7.md` 支持边界 |
| 复杂页眉页脚 | 分节文本与常见页码支持；图片、其他域等部分内容显示为只读占位并保留 | 带公司标志、表格或复杂域的页眉不能按普通正文完整编辑，P1 | `DocxSection.cs:323-350` |
| 行号与断字 | 能设置并保存，但界面提示在 Word 中显示／生效 | Writer 中看到的排版不完全等于打印或 Word，P1/P2 | `ui/WordEditor.dc.html:2486-2487` |
| 比较与合并审阅 | 输出文字／格式差异的独立比较报告 | 尚不是带完整修订的合并文档，律师审稿需区分，P1 | `WordEditor.dc.html:1724-1740`、`engine.js:3065` |
| 修订种类 | 支持插入删除及部分文字／段落格式修订 | 移动内容、表格结构等修订不能视作全覆盖，复杂稿待验证，P1 | `DocxRevisions.cs:73,97,150` |
| 域计算系统 | 有 REF、MERGEFIELD 及目录／编号等专用实现 | 缺完整 IF、ASK、FILLIN、INCLUDETEXT、索引和法律引文目录等创建更新流程，P2 | `DocxFields.cs:11-15`、Registry |
| 表单与内容控件 | 内容控件中的部分内容可投影编辑 | 缺复选框、下拉列表、日期控件、数据绑定的表单设计流程，P2 | `DocxBlocks.cs:15,44`、Registry |
| 邮件合并 | CSV／TSV／JSON 数据生成文档 | 缺完整数据源向导、标签信封版式、批量邮件及条件规则，P2 | `WordEditor.dc.html:1740`、`word-mailmerge.js` |
| Word 图表 | 可新建柱、条、线、面积、饼、环形的基础单系列图表 | 多系列组合、复杂坐标轴、科学图表编辑不足，P2 | `DocxCharts.cs:18-37` |
| SmartArt | 原对象保留、显示支持有限 | 不能编辑节点、层级、布局并保持原生 SmartArt 语义，P2 | `Registry.cs:571-588`、`office-draw.js` |
| OLE／MathType | 保留嵌入对象；WMF 已有本地预览路径 | 不能像 Office 双击进入嵌入应用编辑；MT Extra 字体／EMF 外观待验证，P1/P2 | `DocxObjects` 相关实现、发布边界 |
| 原生公式 | OMML／LaTeX 常见分式、根号、矩阵等已有 | 不等于完整 Word 公式／LaTeX 语言；不支持的结构会降级，P2 | `DocxMath.cs:84-87` |
| 查找替换 | 文字、大小写、全词、跨文字片段匹配已有 | 缺 Word 完整通配符、按格式和特殊标记查找替换，P2 | `ui/text-find.js`、`WordEditor.dc.html:2208` |
| 引文与参考文献 | APA、MLA、Chicago、IEEE、GB/T 7714—2015 顺序编码已有 | 缺完整样式生态和各种细分格式；不能说全部国标格式都支持，P2 | 引文实现、0.1.7 发布边界 |
| 中文高级排版 | 中西文字体、常见缩进编号已有 | 注音、文档网格、装订镜像、竖排等未见完整流程；中英混排外观待样本验证，P1/P2 | Word 编辑器和 Registry 能力核查 |
| 拼写／语法／朗读听写 | 浏览器拼写和通用 AI 助手可用 | 没有与 Microsoft Editor、系统听写朗读相当的一体化工作流，P2 | `WordEditor.dc.html:628`、AI 入口 |
| 主题与模板 | 本地中英文模板、主题和样式集已有 | 缺大型素材库、任意主题资源管理及完整企业模板制作，P2 | `ui/templates/index.json`、`DocxDesign.cs` |
| 限制编辑 | 有只读限制，已有受保护稿可能需在 Word 解锁 | 不等于文件密码加密或完整区域权限，P1/P2 | `DocxDesign.cs:50` |
| 超长文件性能 | 普通段落输入能复用前面分页结果 | 结构／列表／脚注变化仍会完整分页，不能承诺千页复杂稿始终流畅，P1 | 0.1.7 发布边界、分页实现 |

专业流程的参照分别是微软的[法律黑线比较](https://support.microsoft.com/en-us/word/compare-document-differences-using-the-legal-blackline-option)、[可填写表单](https://support.microsoft.com/en-us/word/create-a-form-in-word-that-users-can-complete-or-print)、[邮件合并](https://support.microsoft.com/en-us/word/use-mail-merge-for-bulk-email-letters-labels-and-envelopes)。这些不能仅凭我们存在“比较”“合并”按钮就判定相当。

## 4. Excel：差距最大，尤其是计算可靠性与数据分析深度

已有基础：完整地址边界的虚拟网格、普通公式、筛选与当前区域排序去重、格式与条件格式、图表、冻结、打印范围、选择性粘贴、动态数组、名称、原生表格、结构化引用、分组和基础透视表。**80 行 × 20 列的硬限制已取消。** 微软工作表上限为 1,048,576 行 × 16,384 列，见[官方规格](https://support.microsoft.com/en-us/excel/excel-specifications-and-limits)；支持这个地址范围不意味着可以同时物化、计算或绘制其中所有格子。

| 能力 | 我们现在及欠缺 | 使用影响／优先级 | 主要代码证据 |
|---|---|---|---|
| 合并／撤销／保存 | 本次发现并修复 general 校验不一致；补真实文件回归 | 本地已修，发布版仍需更新才能收到，P0 | Registry、`xlsx-history.test.mjs` |
| 超大网格 | 最后一个地址 `XFD1048576` 可解析，网格虚拟化 | 部分范围计算仍有限制，不能据此称为无规模限制，P1 | `sheet-grid.js:2`、`sheet-engine.js:18,1020` |
| 基础函数覆盖 | 当前注册 276 个名称，包含别名；已有 LET、FILTER、SORT、UNIQUE 等 | 不是 276 个函数均已达到 Excel 所有边界语义，P1 | `sheet-engine.js` FUNCS、探针 |
| 高阶公式 | 未注册 LAMBDA、MAP、REDUCE、SCAN、BYROW、BYCOL、MAKEARRAY | 新式复杂模板会返回错误，P1/P2 | 函数表与实际求值探针 |
| 部分动态数组 | HSTACK、VSTACK、TAKE、DROP、CHOOSECOLS、TEXTSPLIT 已有 | TOCOL、TOROW、WRAPROWS 等尚缺，P2 | 函数表与探针 |
| 专业统计／金融／预测 | 常见统计金融函数有一部分 | LINEST、LOGEST、T.TEST、CHISQ.TEST、FORECAST.ETS、PRICE、YIELD 等未注册，P1/P2 | 函数表与探针 |
| 数据服务公式 | 未注册 GETPIVOTDATA、CUBEVALUE、STOCKHISTORY、IMAGE、PY | 商业分析、关联数据、股票数据、Python 工作流不完整，P2/P3 | 函数表与探针；部分微软功能依赖订阅／平台 |
| 三维引用 | `SUM(Sheet1:Sheet3!A1)` 实测 `#NAME?` | 跨月份工作表汇总受阻，P1 | 引用解析及探针 |
| 外部工作簿链接 | 外部文件引用无法自动解析刷新，探针返回 `#REF!` | 多文件财务模型不能当作完整支持，P1 | 引用解析及探针 |
| 名称作用域 | 全工作簿名称已有；局部名称在原生文件中保留，但暴露与计算路径跳过 LocalSheetId | 同名局部名称可能无法正确计算／编辑，P1 | `XlsxBookFeatures.cs:11-14`、`XlsxCalculation.cs:92` |
| 迭代计算与兼容语义 | 普通依赖计算、缓存和循环错误检测已有 | 未见完整迭代计算配置；空白／错误值／日期等边界仍需 Excel 对照样本，P1 | Calc、原生计算路径 |
| 1900／1904 日期系统 | 普通日期值的读写支持两套；公式计算内部使用 1900 序号，写回没有按 1904 转换 | 已复现日期公式缓存偏移 1462 天，P0；普通日期支持与公式支持必须分开验收 | `XlsxCells.cs:85-90`、`XlsxCalculation.cs:97-110,136`、`sheet-engine.js:179` |
| 透视表 | 一个行字段、可选一个列字段、一个值字段及常用聚合 | 多字段嵌套、计算字段、日期分组等缺口大；复杂已有透视结构保留不等于能完整刷新，P1/P2 | `XlsxPivots.cs:99-104`、`sheet-engine.js:1223` |
| 切片器／时间线／透视图 | 未见完整交互与创作流程 | 交互仪表盘能力不足，P2 | 透视表 UI、Registry 核查 |
| Power Query | 有普通导入；没有查询编辑器、连接器、M 语言变换与刷新链 | 周期性清洗、合并数据源无法等价替代，P2 | 导入与表格工具核查 |
| Power Pivot／DAX | 未见关系数据模型、DAX、模型度量值 | 企业分析能力缺失；属于 Windows 高级能力对照，P2/P3 | 数据模型与 Registry 核查 |
| 假设分析／规划求解 | 未见目标求解、方案管理、模拟运算表、Solver 或 Analysis ToolPak 工作流 | 预算优化、敏感性分析需其他工具，P2 | 编辑器数据工具核查 |
| 排序 | 当前区域、多关键字排序已有，UI 最多三关键字 | 按颜色／图标／自定义序列、横向排序等不足，P2 | `SheetEditor.dc.html:803` |
| 筛选 | 值列表与比较条件已有 | 完整高级条件区域提取、颜色与动态日期筛选不足，P2 | SheetEditor 筛选实现 |
| 数据验证 | 列表、整数、小数、日期、文本长度支持；时间、自定义未检查 | 会允许违反原文件规则的输入；公式／空白也有跳过路径，P1 | `SheetEditor.dc.html:872`、探针 |
| 条件格式 | 编辑器有基础规则、色阶、数据条、图标集、公式 | 打印缺失已复现；复杂优先级／stopIfTrue 等兼容性仍需补，P1 | `renderVals`、`sheet-print.js` |
| 图表类型 | 可创建柱、条、线、饼、面积、散点、环形、组合图 | 某些导入图表可保留或显示，但雷达、瀑布、直方图、树图等不能当作完整可创建编辑，P2 | `XlsxCharts.cs:209`、`Registry.cs:252` |
| 图表分析选项 | 有基础系列、标签、标题和图例 | 完整坐标轴设置、误差线、趋势线等缺少完整编辑流程；迷你图未见实现，P2 | 图表 UI 与 Registry |
| 分列与快速整理 | 有分隔符拆分、普通填充 | 分列使用字符串拆分，缺完整引号语义、固定宽度、逐列类型向导；无完整快速填充推断，P2 | `SheetEditor.dc.html:825` |
| 工作表保护 | 有保护开关，界面整体限制编辑 | 未见允许编辑未锁定单元格、细分权限和完整密码流程，P1/P2 | `XlsxBookFeatures.cs:75`、SheetEditor |
| 批注与协作 | 有传统单元格备注 | 不是现代线程批注、@提及、多人同时编辑和云端版本历史，P2/P3 | `XlsxNotes.cs` |
| 打印 | 打印范围、重复标题、纸张缩放、图片图表已有 | 已确认条件格式、部分边框、删除线、旋转、格式码颜色及工作簿默认字体有遗漏；手动分页等也不足，P1 | `sheet-print.js:20-26,91-92`、第二轮探针 |
| 大表速度 | 虚拟网格与共享快照已实现，百万值有既往样本 | 百万公式／跨表依赖不等价；原生计算引擎有 30 秒、512 MB 和递归限制，需单独压测，P1 | `XlsxCalculation.cs:99`、发布边界 |

函数以[微软函数目录](https://support.microsoft.com/en-us/excel/excel-functions-alphabetical)为参照，不把函数数量直接换算成功能覆盖率。[Power Query](https://support.microsoft.com/en-us/excel/about-power-query-in-excel)是完整的数据获取变换系统，[Power Pivot](https://support.microsoft.com/en-us/excel/power-pivot-powerful-data-analysis-and-data-modeling-in-excel)涉及数据建模；二者不能由“能导入 CSV、能建透视表”替代。平台覆盖应结合[数据源支持表](https://support.microsoft.com/en-us/excel/power-query-data-sources-in-excel-versions)判断。[假设分析](https://support.microsoft.com/en-us/excel/introduction-to-what-if-analysis)也有独立的使用流程。

补充核实：SUMIFS、COUNTIFS、SUMPRODUCT、XLOOKUP、XMATCH、VLOOKUP、INDEX/MATCH、IFS、SWITCH、TEXTJOIN、SEQUENCE、XIRR、XNPV、NETWORKDAYS、WORKDAY、DATEDIF、SUBTOTAL、AGGREGATE 都已注册；不能把它们笼统列成缺失。这里验证的是存在实现，不代表全部边界条件已与 Excel 对照。微软两种日期系统相差 1462 天，见[日期系统说明](https://support.microsoft.com/en-us/excel/date-systems-in-excel)；本次原生文件探针发现的偏移与该差值相同。

## 5. PowerPoint：基础制作可用，复杂对象和播放还原不足

已有基础：实际字号和文字格式、图片裁剪、形状连接线、表格、组合层次、对齐、主题、布局与母版继承、备注、演讲者视图、分节、基本动画与路径、切换、音视频、排练和保存墨迹。微软各平台的功能差别见[官方平台对照](https://support.microsoft.com/en-us/powerpoint/compare-powerpoint-features-on-different-platforms)。

| 能力 | 我们现在及欠缺 | 使用影响／优先级 | 主要代码证据 |
|---|---|---|---|
| 图表编辑 | 导入的图表对象可保留、部分可绘制 | 缺原生插入图表、数据表与系列编辑的完整流程，P1/P2 | SlideEditor 插入工具、Registry |
| SmartArt | 原对象保留／显示 | 缺节点增删、层级重组、布局切换，P2 | 对象模型与只读属性 |
| 母版编辑 | 可修改母版文字／形状的部分属性 | 缺完整新建母版／版式、占位符、图片图表及全画布管理，P1/P2 | `PptxMasterEdit.cs:12,30` |
| 备注／讲义母版 | 普通备注与演讲者视图已有 | 缺完整备注母版、讲义母版制作流程，P2 | SlideEditor、打印路径 |
| 大纲 | 有导航列表 | 不是能直接编辑和重排正文结构的大纲编辑器，P2 | `SlideEditor.dc.html:32` |
| 动画效果与时序 | 可创作 15 个预设效果，支持路径、时长、延迟、同时／之后 | 复杂触发器、重复、自动反向、逐字逐词、图表分项等不完整，P1/P2 | `office-io.js:133` 起、播放实现 |
| 已有动画的播放 | 不支持的效果采用回退动画 | 文件未丢失原始设置，也可能现场播放不一样，P1 | `fxFrames`、探针 |
| 切换与平滑 | 有基本切换和轻量 Morph；按对象特征匹配，主要平移缩放 | dissolve 实际回退淡入淡出；缺完整旋转、文字和形状形态变换，P1/P2 | `transitionFrames`、`morphFrames`、探针 |
| 幻灯片缩放导航 | zoom 切换效果已有 | 不等于摘要缩放、节缩放和幻灯片缩放对象，P2 | 切换实现与对象注册表 |
| 音视频剪辑与编排 | 可插入并使用浏览器控件播放 | 缺剪裁、书签、淡入淡出、跨页播放和完整时间线；编码兼容待测，P1/P2 | `PptxMedia`、`SlideEditor.dc.html:529` |
| 录制与视频输出 | 有放映／排练，不具备完整录制流水线 | 缺旁白、摄像头、录屏整合及导出 MP4，P2 | 导出与媒体 UI 核查 |
| 墨迹 | 可保存为透明图片 | 不是可重新编辑笔画的原生墨迹，也没有完整墨迹回放／转形状，P2 | `SlideEditor.dc.html:541` |
| 批注 | 传统批注可编辑；现代批注保留并只读显示 | 现代回复、@提及、解决与云端协作不足，P2 | `PptxComments.cs:26-30` |
| 打印与 PDF | 基本逐页幻灯片打印已有 | 缺 2／3／6／9 页讲义、完整备注页等常用输出版式，P1/P2 | `office-io.js:529` 起 |
| 导出 PNG 的外观 | 有 Canvas 导出流程，背景图片可以进入绘制 | 局部文字格式被压平；图片裁剪和灰度未用于绘制。导出图片目前不能承诺与编辑器一致，P1 | `office-io.js:413-489`、第二轮 Canvas 命令探针 |
| 高级形状／3D | 常见形状、连接线、组合已有 | 未见完整布尔合并、自由曲线节点编辑、3D 及 SVG 组件编辑流程，P2/P3 | 形状 UI 与 Registry |
| 交互式演示 | 常规链接和演示流程已有 | 自定义放映、缩放对象、复杂动作触发等未见完整流程，P2 | SlideEditor 放映／动作核查 |
| 主题模板／智能设计 | 本地模板、主题、AI 美化已有 | 缺 Office 全套模板资源与云端设计服务；不能说完全没有 AI，P2/P3 | 模板目录与 AI 入口 |
| 播放与资源性能 | 普通内容有本地编辑／播放实现 | 大量视频、字体替换、巨型演示及 Mac/Windows 播放一致性仍待实测，P1 | 本次未做完整桌面性能验收 |

参照：[Morph 的详细能力](https://support.microsoft.com/en-us/powerpoint/morph-transition-tips-and-tricks)、[缩放导航](https://support.microsoft.com/en-us/powerpoint/use-zoom-for-powerpoint-to-bring-your-presentation-to-life)、[录制演示](https://support.microsoft.com/en-us/powerpoint/record-your-presentation)、[输出视频](https://support.microsoft.com/en-us/powerpoint/turn-your-presentation-into-a-video)。另外，微软已从 Microsoft 365 PowerPoint Windows 2502 起移除比较与合并功能，不能继续把它算成对标当前版必补的缺口，见[官方说明](https://support.microsoft.com/en-us/powerpoint/track-changes-in-your-presentation)。

## 6. 三个编辑器共同的差距

| 领域 | 当前边界 | 建议 |
|---|---|---|
| 文件兼容 | DOCX/XLSX/PPTX 有原生读写和未编辑部分保留；保存保留、显示正确、可编辑是三种能力 | 建立真实 Office 文件集，分别记录三种结果，P1 |
| 宏文件 | DOCM/XLSM/PPTM 通过转换打开，转换副本去除 VBA 项目；原文件不因此覆盖 | 明确提示另存普通格式会失去宏，不应宣称宏兼容，P1 |
| 旧格式 | DOC/XLS/PPT 等走兼容转换，非原格式无损编辑保存；未见 XLSB 原生适配器 | 依照格式分别列导入和导出能力，P1/P2 |
| 密码／权限／签名 | 编辑保护不等于文件加密；加密 Office 文件需先解锁。缺完整 IRM、敏感度标签、文档数字签名工作流 | 企业用户的重要边界，P1/P2；应用安装包签名与文档签名无关 |
| 多人协作 | 本地文件变化监听及撤销存在；未见多人实时协作、云端权限和跨设备版本历史 | 是独立的大功能方向，P3 |
| 插件与自动化 | 自有 CLI/MCP/AI 能操作文档 | 不兼容 VBA、COM、完整 Office Add-ins／Office Scripts 生态，P2/P3 |
| 无障碍 | 已有部分替代文字能力 | 缺完整可访问性检查、阅读顺序校验与带结构标签 PDF 验证，P1/P2 |
| 实际资源占用 | 包大小、磁盘占用和运行 RAM 是不同指标；WebView、媒体、公式量都影响 RAM | 需要统一样本测空闲、普通文件、巨型文件的完整进程树，不能从约 160 MB 安装体积推断 RAM |

兼容转换的证据：`src/Writer.Formats/Compat/OoxmlVariants.cs:11-38`、`CompatAdapter.cs:15-18`。微软订阅云功能与买断版也不完全相同，见[套件说明](https://support.microsoft.com/en-us/office/system-requirements/office-suites-for-individuals-and-families)。

## 7. 推荐实施顺序

1. **先保证编辑可信。** 修复 1904 日期公式缓存错误，把此次合并撤销修复交付到实际安装版本；同类场景以“操作 → 撤销／重做 → 保存 → 重开”验收。覆盖清除格式、粘贴、行列增删、排序去重、表格命令，不能只检查菜单或内存变化。
2. **补已确认的日常正确性问题。** Excel 时间／自定义验证、打印样式、超大范围求值；Word 多栏与分节打印；PPT PNG 输出保真，以及明示或改善动画降级。
3. **补高价值业务功能。** Excel 多字段透视和常用缺失公式；Word 专业比较审阅、表单与复杂页眉页脚；PPT 图表／SmartArt、讲义输出和媒体编辑。
4. **再决定高级生态投入。** Power Query／DAX、宏兼容、多用户协作、企业权限、录制与视频输出都是单独的大项目，不是几个按钮能补齐。

## 8. 本次验证记录与未覆盖范围

```sh
dotnet build src/Writer.Cli/Writer.Cli.csproj --no-restore --nologo -v quiet
dotnet test tests/Writer.Tests/Writer.Tests.csproj --no-build --nologo -v quiet
node --test ui/tests
node docs/audits/office-0.1.7-probes.mjs
git diff --check
```

- 第一轮合并撤销修复时构建成功；原生 930 项、UI／桥接 1074 项通过，0 失败，0 跳过。第二轮未改产品源码，未重复跑整套回归；新增工作流探针单独运行成功。
- 截图问题先通过新增测试复现，再修复；新增跨层测试使用真实引擎和临时 `.xlsx` 文件，结束后清理。
- 未做此次桌面点击录像、真实 Microsoft Office 外观逐页比对、所有字体／编码组合、百万公式压力测试，也未制作或公证新的安装包。
- 历史 `office-gap.md`、`office365-gap.md` 包含旧状态；本报告按当前代码重新分类。已发布 0.1.7 的说明见 `docs/releases/0.1.7.md`。本次本地修复与已安装／已发布版本须分别看待。

## 9. 按实际使用场景判断能否替代

下表是基于已核查实现和复现结果的适用性判断，不是每种文件都已人工验收。`部分` 表示必须检查限制；`缺口` 表示当前工作流会被阻塞；`待验` 表示本轮证据不足。保存原始对象不等于具备相应编辑能力。

| 用户实际要做的事 | 打开／显示 | 编辑／计算 | 保存重开 | 打印／播放／导出 | 当前结论 |
|---|---|---|---|---|---|
| Word：普通通知、简历、短报告 | 已有 | 已有 | 有回归覆盖 | 复杂版式待验 | 已有较完整基础，仍需检查目标模板 |
| Word：公司信笺，页眉有标志与表格 | 部分；表格不显示 | 复杂页眉部分只读 | 原生表格保留已实测 | 编辑器打印路径缺内容 | 正式信笺不宜承诺等价输出 |
| Word：双栏论文、横竖混排的长报告 | 部分 | 格式设置可写入 | 原生设置已有 | 多栏／混合尺寸不完整 | 主要缺口在排版与输出 |
| Word：多人审合同，合并对方修订 | 常见修订已有 | 部分 | 常见路径有回归 | 比较结果是独立报告 | 缺专业修订合并流程 |
| Word：申请表，复选框／日期／下拉框 | 部分内容可见 | 无完整控件设计 | 原件保留不等于可编辑 | 待验 | 缺表单制作能力 |
| Word：批量制作信件、标签、信封 | 基础合并已有 | CSV 等可用 | 文档输出已有 | 标签信封流程不足 | 批量文档部分可用，完整邮寄流程欠缺 |
| Excel：普通清单、简单预算、查找汇总 | 已有 | 常见函数已实现 | 合并撤销本地已修 | 打印样式有明确缺陷 | 基础编辑广，输出可靠性须先补 |
| Excel：旧 Mac 工作簿中的日期公式 | 普通日期读取正确 | 1904 系统不完整 | 已复现错误日期缓存 | 输出可能沿用错误结果 | **P0，当前不能视为可靠** |
| Excel：多月份工作表、多文件汇总 | 表和公式可见 | 三维／外部引用不足 | 正确结果不能保证 | 结果可靠性不足 | 财务模型存在核心阻塞 |
| Excel：按地区、产品、月份做交互报表 | 原数据已有 | 透视字段数量受限 | 复杂透视不能完整维护 | 缺切片器等交互 | 距完整业务分析仍有明显差距 |
| Excel：每月把多份来源表清洗、合并、刷新 | 普通导入已有 | 缺查询／变换链 | 无完整查询流程 | 无完整刷新链 | Power Query 类工作流尚缺 |
| PPT：文字、图片、形状的常规汇报 | 已有 | 已有 | 常见对象有回归 | 普通播放已有；PNG 降级 | 基础制作可用，交付格式要单独验收 |
| PPT：带图表、SmartArt 的业务汇报 | 保留／部分显示 | 缺完整数据和结构编辑 | 原对象保留 | 复杂对象输出待验 | 不能作为完整编辑替代 |
| PPT：同一句话强调字号颜色，导出宣传图片 | 编辑器支持 | 支持局部格式 | 原生保存路径已有 | PNG 压平局部格式已复现 | 当前 PNG 不等于画布所见 |
| PPT：精确动画、视频书签触发的培训课件 | 部分 | 部分动画与媒体 | 原始标记可保留 | 触发器／效果降级 | 缺完整课件播放能力 |
| PPT：录旁白视频、输出 MP4、打印带备注讲义 | 部分基础内容已有 | 缺录制制作流程 | 普通演示稿可保存 | 视频输出／讲义版式不足 | 整套交付工作流未齐 |

从场景看，Word 的主要投入应放在“复杂排版与审阅”，Excel 在“结果正确与分析能力”，PPT 在“对象编辑与输出／播放保真”。三者都不能只按照功能区按钮数量来比较。

## 10. 可直接进入开发的优先清单

这是由对比得出的工作拆分；本轮没有开始实现这些新项。每项都有具体完成条件，避免“功能加上了”却不能完整使用。

| 顺序 | 工作 | 涉及位置 | 完成判据 |
|---|---|---|---|
| P0-1 | Excel 两套日期系统贯通计算与写回 | `sheet-engine.js`、`XlsxCalculation.cs`、日期模型 | 同一日期在 1900／1904 文件中，直接值、引用、加减、DATE／YEAR／TEXT、保存缓存均正确；原始值不变。首先让本报告的 2024→2028 样本通过 |
| P0-2 | 合并撤销修复进入交付版本 | Registry、XLSX 历史回归、发布流程 | 安装版实际完成合并→撤销→自动保存→关闭→重开，不再出现 general 报错；打包时仍须签名／公证 |
| P1-1 | Excel 编辑与打印共用样式含义 | `SheetEditor.dc.html`、`sheet-print.js` | 条件格式、细线／双线、单边线、删除线、旋转、数字格式颜色、默认字体字号均在打印中正确；保留现有合并、隐藏行、重复标题与图片图表功能 |
| P1-2 | Excel 数据验证完整执行 | `dvBad`、输入／粘贴／编辑路径 | 时间与自定义规则实际生效；检验单格输入、粘贴、公式和空白；支持规则引用的相对／绝对地址，错误提示与原文件一致 |
| P1-3 | 大范围求值采用稀疏路径 | Calc 的引用与聚合 | 显式范围 `A1:A1048576` 与整列引用给出正确 SUM/COUNT/AVERAGE；空白和错误语义一致，不为一列空格创建百万个对象 |
| P1-4 | Word 页眉页脚块级内容建模 | `DocxSection.cs`、Word 页眉 UI、打印 | 表格／图片／域能显示并打印；编辑一个段落后未编辑的表格、图片、关系不损坏；每节及首页／奇偶页分别检查 |
| P1-5 | Word 多栏与分节排版贯通 | 页面模型、分页器、`word-print.js` | 两栏、三栏、横竖混排、不同页边距、脚注和跨页表格在编辑、页数、打印中一致；用固定字体的 Word/PDF 样本比对 |
| P1-6 | PPT PNG 与画布使用相同样式规则 | `paintSlide`、文字／图片渲染 | 局部字号、字体、颜色、粗斜体、裁剪、灰度等不降级；分别对文本、图片、表格、原生对象做像素或图像差异验收 |
| P1-7 | PPT 动画支持范围与播放一致 | `fxFrames`、`transitionFrames`、原生动画模型 | 名称对应真实效果；至少补 dissolve 和 Morph 旋转；未支持触发器不能被无提示当作正常播放；结合真实播放样本 |
| P2-1 | Excel 多字段透视与常用缺失公式 | 透视模型、缓存、计算器、UI | 至少可做“地区→产品”两级行、月份列、销量＋金额两项值；保存重开仍可刷新；GETPIVOTDATA 可引用结果 |
| P2-2 | Word 专业审阅与表单 | 修订／比较、内容控件模型 | 比较结果可按差异接受拒绝并保存；控件可创建、填写、保留类型；复杂结构修订有明确定义 |
| P2-3 | PPT 图表／SmartArt 与常用交付 | 原生对象模型、UI、导出 | 修改数据／节点后仍是原生可编辑对象；补讲义与备注页，再扩展媒体剪裁、录制和视频输出 |

跨文件引用、Power Query、DAX、多人协作、宏生态宜分别设独立项目级范围；不要把它们塞成上述正确性修复中的“小功能”。这里没有给小时级估时，因为尚未做各子系统设计和真实样本规模评估。

## 11. 对标对象的平台与云服务边界

| 能力 | 对比时应怎样计算 |
|---|---|
| Power Query | Windows、Microsoft 365 Mac、网页版均有不同范围的能力。Mac 不能简单标成“没有”；微软当前文档列出订阅版 Mac 的查询编辑器，连接器和刷新范围仍须按平台表核对。Writer 目前缺完整查询链 |
| Power Pivot／DAX | 完整模型编辑器和高级分析以 Windows 对照为主，不能当作所有 Mac 套件的共同能力 |
| Python in Excel | 微软当前说明覆盖 Windows、Mac、Web，计算在微软云端，需要网络；与本地嵌入一个 Python 解释器不是同一工作流。Writer 目前没有该工作流 |
| 宏与 Office 插件 | 桌面、Web、Mac 的插件机制和限制不同。自有 CLI/MCP 是 Writer 的能力，但不能替代现成 VBA／插件的兼容性 |
| 云协作与 Copilot | 受账号、套餐、网络及组织服务影响，应列为独立对标维度，不把它们当作本地安装包中全部静态功能 |
| 演示媒体触发 | 微软可按点击对象或媒体书签触发动画；Writer 目前的 click／with／after 时序不能等价覆盖它 |

平台依据：[Mac Power Query 编辑器说明](https://support.microsoft.com/en-us/excel/import-and-shape-data-in-excel-for-mac-power-query)、[Power Query 各平台数据源](https://support.microsoft.com/en-us/excel/power-query-data-sources-in-excel-versions)、[Power Query 与 Power Pivot 能力](https://support.microsoft.com/en-au/excel/learn-to-use-power-query-and-power-pivot-in-excel)、[Python in Excel 与云计算要求](https://support.microsoft.com/en-us/excel/python/introduction-to-python-in-excel)、[PowerPoint 动画触发器](https://support.microsoft.com/en-us/powerpoint/trigger-an-animation-effect)。Mac Power Query 页面内仍含旧的“不支持编辑器”措辞，与同页明确写出的 16.69 起正式支持章节不一致；本报告采用带版本条件的现行说明，不将旧句子作为“Mac 没有”的依据。

## 12. 第三轮：日常操作还缺什么

本节依据当前编辑器命令、状态模型、保存属性及微软文档核查。“未见完整流程”是代码审计结论，没有伪装成逐个按钮的桌面人工测试；本轮没有修改产品源码或重跑此前已通过的整套回归。

### Word

| 日常能力 | 现状与差距 | 具体影响 | 证据／建议 |
|---|---|---|---|
| 自定义制表位真正参与排版 | 标尺可设左／中／右／小数点制表位，也能落盘；画布仍采用默认半英寸 Tab，未按这些设置布局 | 报价单中的金额、姓名与签字栏在 Writer 与 Word 中可能对不齐 | `WordEditor.dc.html:15,1975-1988`、`engine.js:1542-1568`；P1 |
| 任意文字范围书签 | 当前 UI 给光标所在整段设置单个 `data-w-bookmark` | 无完整“选中几个字或图片，精确标记该范围”的书签创建流程 | `WordEditor.dc.html:1694-1714`；P2 |
| 完整交叉引用选项 | 已有标题／题注／书签引用，写入 `REF name \\h`；字段校验仅接受这个简化语法 | 缺引用页码、编号的不同格式、上方／下方等完整选择流程 | `WordEditor.dc.html:1752-1756`、`DocxFields.cs:11-15`；P2 |
| Word 表格公式 | 表格插入、合并拆分、排序已有；没有完整的单元格公式输入与计算更新流程 | 无法像 Word 那样在文档表格里使用 `SUM(ABOVE)` 等自动合计并更新 | Registry 的单元格 formula 限 XLSX，DocxFields 未实现表格算术域；P2 |
| 精确图片环绕 | 基本环绕已有；tight／through 的导入显示也主要落到矩形 CSS 浮动，UI 常用选择为嵌入、四周、前后 | 缺轮廓环绕点编辑；复杂图文混排不能保证与 Word 一致 | `WordEditor.dc.html:2589`、`picture.js:238-258`；P1/P2 |

微软参照：[书签可标记文字、图片或位置](https://support.microsoft.com/en-gb/word/add-or-delete-bookmarks-in-a-word-document-or-outlook-message)、[交叉引用选项](https://support.microsoft.com/en-us/word/create-a-cross-reference)、[表格公式](https://support.microsoft.com/en-US/Word/use-a-formula-in-a-word-table)、[制表位设置](https://support.microsoft.com/en-us/word/set-tabs-in-a-table)。

**明确已有，不应列为缺失：** 文本转表格／表格转文本、按当前列排序、重复标题行、目录样式、书签跳转、基本交叉引用。证据包括 `WordEditor.dc.html:2077-2104`、`DocxTable.cs:700,746`、`WordEditor.dc.html:2244-2254`。但“有表格排序”目前是按光标列进行，不能自动等同于 Office 所有排序选项。

### Excel

| 日常能力 | 现状与差距 | 具体影响 | 证据／建议 |
|---|---|---|---|
| 公式引用追踪 | 有公式编辑、彩色引用和显示公式；未见追踪引用单元格／从属单元格的箭头和导航流程 | 查一个总额为何算错，要人工逐层找公式来源 | SheetEditor 命令与公式工具核查；P1/P2 |
| 逐步公式求值 | Calc 能计算结果，未见逐步进入嵌套表达式的 UI | 复杂 IF、查找和组合计算不容易定位错误步骤 | `sheet-engine.js` 与 SheetEditor 命令核查；P2 |
| 定位条件 | 有地址／名称跳转及文字查找；未见“仅公式、仅空白、仅常量、仅错误值”等选择对话框 | 不能一键选空格补值、找硬编码或批量选公式保护 | `SheetEditor.dc.html` 的 `goName`、`finder` 与命令核查；P2 |
| 不连续区域同时选择 | 主选择模型为一个 anchor 与一个 selection，`rng()` 生成单个矩形 | 缺按 Ctrl/Cmd 组合多个分散区域后统一编辑格式的完整流程 | `SheetEditor.dc.html:399,494`；P2 |
| 多张工作表成组编辑 | 当前通过一个 `doc.active` 选择工作表，普通编辑操作针对该表 | 不能选中多个月份的表，一次改相同位置的标题或公式 | `SheetEditor.dc.html:350,378`、工作表标签菜单；P2 |
| 可独立滚动的拆分窗格 | 冻结窗格已实现；读取仅投影 frozen／frozenSplit，写入 frozen；没有普通 split 的交互 | 不能在同一张长表中同时独立查看两段远处数据 | `XlsxLayout.cs:190-220`、SheetEditor 视图工具；P2 |

微软参照：[引用追踪](https://support.microsoft.com/en-us/excel/display-the-relationships-between-formulas-and-cells)、[逐步求值](https://support.microsoft.com/en-us/excel/evaluate-a-nested-formula-one-step-at-a-time)、[定位条件](https://support.microsoft.com/en-us/excel/find-and-select-cells-that-meet-specific-conditions-in-excel)、[工作表成组](https://support.microsoft.com/en-us/excel/group-worksheets)、[Mac 拆分窗格](https://support.microsoft.com/en-us/excel/view-multiple-panes-sheets-or-workbooks)。这些条目以相应页面列出的平台为准，不推断所有平台功能完全一样。

**明确已有：** 名称跳转、选择性粘贴、转置、向下／向右填充、冻结、分组与分类汇总、工作表复制／隐藏、工作簿范围查找替换。这里所说的“工作表成组编辑”与已有“行列分组折叠”是不同操作。

### PowerPoint

| 日常能力 | 现状与差距 | 具体影响 | 证据／建议 |
|---|---|---|---|
| 完整对象选择窗格 | 已有画布选择和前后层级操作；未见按名称列出全部对象、选择被遮挡对象、临时隐藏的完整面板 | 页面复杂时难以选中底下的图、形状或文字框 | SlideEditor 模板与命令核查；P1/P2 |
| 锁定对象防误触 | 现有“锁定”是锁定纵横比，未见锁住对象移动／编辑的完整流程 | 调正文案时容易移动背景装饰或固定元素 | `SlideEditor.dc.html:790-791,980`；P2 |
| 全稿替换指定字体 | 支持修改所选文字字体与主题／母版部分属性；未见“把全稿某字体换成另一字体”命令 | 企业模板换字体、处理缺字字体需逐处调整 | SlideEditor 文字格式、母版及查找命令核查；P2 |
| 形状布尔运算与编辑顶点 | 常规形状、组合、对齐已有；缺联合、相交、剪除、拆分及完整顶点编辑流程 | 做自定义图标、流程图、不规则遮罩需要外部工具 | SlideEditor 形状命令、Registry；P2，前轮条目的具体使用差异 |

微软参照：[选择窗格及对象锁定](https://support.microsoft.com/en-us/powerpoint/use-the-selection-pane-to-manage-objects-in-documents)、[全稿字体替换](https://support.microsoft.com/en-us/powerpoint/change-the-fonts-in-a-presentation)、[形状合并与编辑顶点](https://support.microsoft.com/en-us/powerpoint/draw-a-picture-by-combining-and-merging-shapes)。

**明确已有：** 多选组合、前后层级、对齐／分布、查找替换、图片裁剪与图片效果；共享图片工具还提供依赖桌面辅助程序的抠图与压缩入口，抠图标注 Apple Vision／macOS 14 以上。不能把“缺完整图片编辑器”扩大成“没有抠图”。对应 `ui/picture.js:14,177,205`；其平台可用性仍应与安装版分别核验。

### 这些差距对补功能顺序的影响

三套软件既缺高级业务能力，也缺常用的精细操作。建议在修复前两轮确认的日期、保存、打印和 PNG 问题后，优先补 **Word 制表位排版、Excel 公式查错与定位条件、PPT 对象选择窗格**。它们直接影响日常工作效率，再向更完整的专业功能扩展。
