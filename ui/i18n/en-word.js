// English UI text, keyed by the Chinese source string (rules in ui/i18n.js). Loaded only when the UI is in English.
Object.assign(window.I18N_EN = window.I18N_EN || {}, {
  // top bar / global
  '文件': 'File', 'AI 助手': 'Assistant', '沉浸书写 Ctrl+.': 'Immersive Writing Ctrl+.',
  '查找': 'Find', '替换为': 'Replace With', '查找下一个': 'Find Next', '替换': 'Replace', '全部替换': 'Replace All', '关闭': 'Close',
  '＋ 新页面': '+ New Page', '标题': 'Headings', '用「标题 1/2/3」样式的段落会显示在这里。': 'Paragraphs styled "Heading 1/2/3" will appear here.',
  '中文（简体）': 'Chinese (Simplified)', '＋': '＋', '取消': 'Cancel', '确定': 'OK',
  '缩小 Ctrl+−': 'Zoom Out Ctrl+−', '放大 Ctrl+＋': 'Zoom In Ctrl+＋', '输入缩放比例后回车（20–400）': 'Type a zoom percentage and press Return (20–400)',
  '（空标题）': '(Untitled Heading)',

  // style-picker labels (font list, block styles, highlight colors)
  '思源宋体': 'Noto Serif SC', '思源黑体': 'Noto Sans SC', '等宽': 'Monospace',
  '正文': 'Body', '标题 1': 'Heading 1', '标题 2': 'Heading 2', '标题 3': 'Heading 3', '引用': 'Quote', '代码': 'Code',
  '标题@@title': 'Title', '副标题': 'Subtitle', '列表段落': 'List Paragraph', '无间隔': 'No Spacing', '题注': 'Caption', '强调': 'Emphasis', '明显强调': 'Strong', '强烈强调': 'Intense Emphasis', '不明显强调': 'Subtle Emphasis',
  '样式': 'Styles', '字符样式': 'Character Styles', '清除字符样式': 'Clear Character Style', '修改样式以匹配所选内容': 'Update Style to Match Selection', '新建样式…': 'New Style…', '新建样式': 'New Style', '样式名称': 'Style Name',
  '样式「{name}」已更新': 'Style "{name}" updated', '先选中要应用字符样式的文字': 'Select the text to style first',
  '黄色': 'Yellow', '绿色': 'Green', '蓝色': 'Blue', '红色': 'Red', '无': 'None', '鲜绿色': 'Bright Green', '青绿色': 'Turquoise', '粉红色': 'Pink', '深蓝色': 'Dark Blue', '青色': 'Teal', '紫罗兰': 'Violet', '深红色': 'Dark Red', '深黄色': 'Dark Yellow', '深灰色': 'Gray 50%', '灰色-25%': 'Gray 25%', '黑色': 'Black',
  '苹方': 'PingFang SC', '宋体-简': 'Songti SC', '黑体-简': 'Heiti SC', '楷体-简': 'Kaiti SC', '华文宋体': 'STSong', '华文楷体': 'STKaiti', '华文仿宋': 'STFangsong', '华文细黑': 'STXihei', '宋体': 'SimSun', '黑体': 'SimHei', '微软雅黑': 'Microsoft YaHei', '楷体': 'KaiTi', '仿宋': 'FangSong', '等线': 'DengXian',
  '字符间距': 'Character Spacing', '紧缩 1 磅': 'Condensed 1 pt', '标准': 'Normal', '加宽 1 磅': 'Expanded 1 pt', '加宽 2 磅': 'Expanded 2 pt', '加宽 3 磅': 'Expanded 3 pt',
  '文字效果': 'Text Effects', '轮廓': 'Outline', '阴影': 'Shadow', '更改大小写': 'Change Case', '句首字母大写': 'Sentence case', '小写': 'lowercase', '大写': 'UPPERCASE', '每个单词首字母大写': 'Capitalize Each Word', '切换大小写': 'tOGGLE cASE', '半角': 'Half-width', '全角': 'Full-width',
  '先选中要设置的文字': 'Select the text first',

  // tabs
  '开始': 'Home', '插入': 'Insert', '布局': 'Layout', '审阅': 'Review', '视图': 'View', '表格': 'Table', '图片': 'Picture',
  '格式面板': 'Format panel', '格式标签': 'Format tabs', '大纲': 'Outline', '上一条': 'Previous', '下一条': 'Next', '第 {n} / {total} 页': v => `Page ${v.n} of ${v.total}`,
  '工具': 'Tools', '文字格式与颜色': 'Text formatting and color', '对齐与缩进': 'Alignment and indentation', '列表': 'Lists', '间距、边框与底纹': 'Spacing, borders and shading',
  '常用': 'Common', '形状与文本': 'Shapes and text', '分隔': 'Breaks', '符号与日期': 'Symbols and date', '封面与书签': 'Cover and bookmarks', '目录与批注': 'Contents and comments',
  '符号与批注': 'Symbols and comments', '图表': 'Chart', '插入图表': 'Insert Chart', '插入柱状图图片': 'Insert a bar chart image', '数据标签（逗号分隔）': 'Data labels (comma separated)', '数值（逗号分隔）': 'Values (comma separated)', '请为每个标签输入一个非负数值': 'Enter one nonnegative value for each label',
  '纸张与分栏': 'Paper and columns', '段落设置': 'Paragraph settings', '脚注与尾注': 'Footnotes and endnotes', '题注与交叉引用': 'Captions and cross references',
  '阅读方式': 'Reading mode', '沉浸': 'Immersive', '插入行列': 'Insert rows and columns', '单元格': 'Cells', '样式与边框': 'Styles and borders', '对齐与行高': 'Alignment and row height', '其它': 'Other',
  '大小与环绕': 'Size and wrapping', '填充与轮廓': 'Fill and outline', '形状与大小': 'Shape and size',

  // the format panel (the design's 定稿): groups, controls and their values
  '字形': 'Font Style', '加粗 倾斜': 'Bold Italic', '间距': 'Spacing', '纸张': 'Paper', '方向': 'Orientation', '分栏与分隔': 'Columns & Breaks', '边框与底纹': 'Borders & Shading',
  '换行和分页': 'Line & Page Breaks', '孤行控制': 'Widow/Orphan Control', '避免段落的第一行或最后一行单独出现在页面顶端或底端': 'Keep the first or last line of a paragraph from standing alone at the top or bottom of a page', '项目符号': 'Bullets', '多级列表': 'Multilevel List', '编号样式': 'Numbering Style', '• 项目符号': '• Bullets', // i18n-ok — the marker itself
  '段落边框': 'Paragraph Borders', '多条框线': 'Several Borders', '首行不缩进': 'No First-Line Indent', '首行缩进 2 字符': 'First Line 2 ch', '悬挂缩进 {n}': v => `Hanging ${v.n}`, '首行缩进 {n}': v => `First Line ${v.n}`,
  '{n} 行': v => `${v.n} ${+v.n === 1 ? 'line' : 'lines'}`, '{n} 倍': v => `${v.n}×`, '最小值 {v}': v => `At least ${v.v}`, '固定值 {v}': v => `Exactly ${v.v}`,
  '查找替换 Ctrl+F': 'Find & Replace Ctrl+F', '链接 Ctrl+K': 'Link Ctrl+K', '分页符 Ctrl+Enter': 'Page Break Ctrl+Enter', '插入表格：选择行列': 'Insert Table: choose rows and columns',
  '{size} · {w} × {h} 厘米': v => `${v.size} · ${v.w} × ${v.h} cm`, '{n} 厘米': v => `${v.n} cm`, '{side}页边距': v => `${v.side} margin`, '自定义页边距': 'Custom Margins', '适中': 'Moderate',
  '上下左右相同…': 'Same on All Sides…', '上下对称、左右对称…': 'Top & Bottom, Left & Right…', '上下左右（厘米）': 'All sides (cm)', '上下（厘米）': 'Top and bottom (cm)', '左右（厘米）': 'Left and right (cm)',
  '1.27 厘米': '1.27 cm', '2.54 厘米': '2.54 cm', '2.54 / 1.91 厘米': '2.54 / 1.91 cm', '2.54 / 5.08 厘米': '2.54 / 5.08 cm', '2.54 / 3.18 厘米': '2.54 / 3.18 cm',
  '整个表格': 'Whole Table', '单元格边框，或整个表格的线条': 'Cell borders, or the lines of the whole table', '表格属性，删除表格': 'Table properties, delete table',
  '替换为：回车替换找到的这一处': 'Replace with: press Return to replace the match found', '文档里还没有目录，先点「插入目录」': 'The document has no table of contents yet. Click Insert Table of Contents first.',
  '上': 'Top', '下': 'Bottom', '自动断字': 'Hyphenation', '水印：{w}': v => `Watermark: ${v.w}`,
  '插入目录': 'Insert Contents', '更新目录': 'Update Contents', '根据标题生成目录': 'Build the contents from the headings', '按现在的标题更新目录': 'Update the contents from the headings',
  '目录样式': 'Contents Style', '经典': 'Classic', '简洁': 'Simple', '无页码': 'No Page Numbers', '脚注和尾注': 'Footnotes & Endnotes', '脚注和尾注的编号样式': 'How footnotes and endnotes are numbered',
  '题注标签': 'Caption Label', '标签：{l}': v => `Label: ${v.l}`, '引用：{k}': v => `Refer to: ${v.k}`, '引用的类型': 'What to refer to', '没有可引用的项目': 'Nothing to refer to yet',
  '添加书签': 'Add Bookmark', '跳转到…': 'Go To…', '还没有书签': 'No bookmarks yet', '删除所在段落的书签': "Remove This Paragraph's Bookmark",
  '字符': 'Characters', '页': 'Pages', '文档里还没有批注。选中文字后点「新建」添加。': 'No comments yet. Select some text and click New.', '记录修改': 'Track Changes',
  '全部接受': 'Accept All', '全部拒绝': 'Reject All', '查找与替换': 'Find & Replace', '下一个': 'Next', '替换这一处': 'Replace this one',
  '网页': 'Web', '阅读': 'Reading', '翻页': 'Page Turning', '页面排列': 'Arrange Pages', '上下': 'Vertical', '左右': 'Side by Side', '页面上下翻': 'Pages scroll vertically', '页面并排左右翻': 'Pages side by side',
  '行和列': 'Rows & Columns', '合并右侧': 'Merge Right', '合并下方': 'Merge Down', '拆分': 'Split', '{side}对齐': v => `Align ${v.side}`, '转文本': 'To Text', '属性': 'Properties',
  '第一行使用标题行的格式': 'Format the first row as a header', '网格表 4（橙色标题行）': 'Grid Table 4 (orange header)', '大小与位置': 'Size & Position', '形状与排列': 'Shape & Arrangement',
  '调整': 'Adjust', '边框与效果': 'Border & Effects', '压缩': 'Compress', '重置': 'Reset', '增加': 'Increase', '减少': 'Decrease',
  '新页面': 'New Page', '缩小': 'Zoom Out', '放大': 'Zoom In', '交给 AI 助手改写': 'Rewrite with the assistant', '高亮（{c}）': v => `Highlight (${v.c})`, '上下左右': 'All Sides',

  // home ribbon
  '撤销 Ctrl+Z': 'Undo Ctrl+Z', '重做 Ctrl+Y': 'Redo Ctrl+Y',
  '格式刷': 'Format Painter', '复制格式后选择目标文字': 'Copy the format, then select text to apply it', '清除格式': 'Clear Formatting',
  '字体': 'Font', '字号': 'Size', '增大字号 Ctrl+Shift+>': 'Grow Font Ctrl+Shift+>', '减小字号 Ctrl+Shift+<': 'Shrink Font Ctrl+Shift+<',
  '加粗 Ctrl+B': 'Bold Ctrl+B', '斜体 Ctrl+I': 'Italic Ctrl+I', '下划线 Ctrl+U': 'Underline Ctrl+U', '删除线': 'Strikethrough',
  '上标': 'Superscript', '下标': 'Subscript', '文字颜色': 'Text Color', '高亮': 'Highlight', '段落样式': 'Paragraph Style',
  '左对齐': 'Align Left', '居中': 'Center', '右对齐': 'Align Right', '两端对齐': 'Justify', '分散对齐': 'Distribute',
  '固定值…': 'Exactly…', '行距（磅）': 'Line Spacing (pt)', '段落…': 'Paragraph…', '段落': 'Paragraph',
  '段前（磅）': 'Before (pt)', '段后（磅）': 'After (pt)', '行距（倍数或磅）': 'Line Spacing (multiple or pt)', '左缩进（厘米）': 'Left Indent (cm)', '右缩进（厘米）': 'Right Indent (cm)', '首行缩进（字符，负数为悬挂）': 'First Line Indent (characters; negative hangs)',
  '边框': 'Borders', '下框线': 'Bottom Border', '上框线': 'Top Border', '左框线': 'Left Border', '右框线': 'Right Border', '外侧框线': 'Outside Borders', '段落底纹': 'Paragraph Shading',
  '悬挂缩进 2 字符': 'Hanging Indent 2 Characters', '换行和分页': 'Line and Page Breaks', '与下段同页': 'Keep with Next', '段中不分页': 'Keep Lines Together', '段前分页': 'Page Break Before',
  '标尺': 'Ruler', '制表位 {pos}（点击切换类型或移除）': 'Tab stop at {pos} (click to change its kind or remove it)',
  '点击标尺添加制表位；点击制表位切换 左 / 居中 / 右 / 小数点，再点一次移除': 'Click the ruler to add a tab stop; click a stop to cycle left / center / right / decimal, once more to remove it',
  '• 列表': '• List', '1. 列表': '1. List', '减少缩进 Shift+Tab': 'Decrease Indent Shift+Tab', '增加缩进 Tab': 'Increase Indent Tab', '行距': 'Line Spacing',
  '编号': 'Numbering', '编号库': 'Numbering Library', '重新开始编号': 'Restart Numbering', '继续编号': 'Continue Numbering',
  '编号样式：1. a. i.、1. 1.1 1.1.1、一、（一）1.；Tab / Shift+Tab 升降级': 'Numbering styles: 1. a. i., 1. 1.1 1.1.1, 一、（一）1.; Tab / Shift+Tab change the level',
  '查找替换': 'Find & Replace',

  // insert ribbon
  '链接': 'Link', '分隔线': 'Horizontal Line', '分页符': 'Page Break',
  '标题@@sel': 'Heading', '改写': 'Rewrite', // the Focus Mode selection bar
  '页眉': 'Header', '编辑页眉（也可双击页面顶部）': 'Edit the header (or double-click the top of the page)',
  '页脚': 'Footer', '编辑页脚（也可双击页面底部）': 'Edit the footer (or double-click the bottom of the page)',
  '页码': 'Page Number', '在页脚居中插入「第 X 页 / 共 Y 页」': 'Insert "Page X of Y" centered in the footer',
  '日期': 'Date', '符号': 'Symbol', '目录': 'Table of Contents', '根据标题生成/更新目录': 'Build or update the table of contents from headings',
  '批注': 'Comment',

  // layout ribbon
  '纸张：{size}': 'Paper: {size}', '横向': 'Landscape', '纵向': 'Portrait',
  '页边距': 'Margins', '窄': 'Narrow', '普通': 'Normal', '宽': 'Wide',
  '分栏：{n}': 'Columns: {n}', '{n} 栏': v => v.n === 1 ? '1 Column' : v.n + ' Columns',
  '页面颜色': 'Page Color',
  '水印': 'Watermark', '机密': 'Confidential', '草稿': 'Draft', '严禁复制': 'Do Not Copy',
  '自定义…': 'Custom…', '自定义水印': 'Custom Watermark', '水印文字': 'Watermark Text',
  '段前': 'Before', '段后': 'After', '首行缩进': 'First Line Indent', '2 字符': '2 Characters',
  '应用于本节': 'Applies to this section', '分隔符': 'Breaks', '分节符': 'Section Breaks', '分节符（{kind}）': 'Section Break ({kind})', '偶数页': 'Even Page', '奇数页': 'Odd Page', '删除分节符': 'Remove Section Break',
  '分页符与分节符：每节可以有自己的纸张、方向、页边距和分栏': 'Page and section breaks: each section can have its own paper, orientation, margins and columns',
  '行号@@word': 'Line Numbers', '在页边显示行号（Word 中显示）': 'Line numbers in the margin (shown in Word)',
  '断字': 'Hyphenation', '英文单词在行尾自动断字（Word 中生效）': 'Hyphenate English words at line ends (applied in Word)',

  // references ribbon and the notes panel
  '引用@@tab': 'References', '插入脚注': 'Insert Footnote', '在光标处插入脚注，注释显示在页面底端': 'Insert a footnote at the caret; its text goes at the foot of the page',
  '插入尾注': 'Insert Endnote', '在光标处插入尾注，注释显示在文档末尾': 'Insert an endnote at the caret; its text goes at the end of the document',
  '脚注 {n}': 'Footnote {n}', '尾注 {n}': 'Endnote {n}', '脚注和尾注 · {{ noteCount }}': 'Notes · {{ noteCount }}', '输入注释文字…': 'Type the note…',
  '插入题注': 'Insert Caption', '图、表、公式的题注，编号自动更新': 'Captions for figures, tables and equations, numbered automatically',
  '交叉引用': 'Cross-reference', '没有可引用的标题、题注或书签': 'No headings, captions or bookmarks to refer to', '引用标题、题注或书签；在 Word 中按住 Ctrl 点击可跳转': 'Refer to a heading, caption or bookmark; Ctrl-click it in Word to go there',

  // insert: text boxes, shapes, equations, cover pages, drop caps, bookmarks
  '插入文本框，可拖到页面任意位置': 'Insert a text box you can drag anywhere on the page',
  '公式@@word': 'Equation', '用 LaTeX 输入公式，保存为 Word 公式': 'Type an equation in LaTeX; it is saved as a Word equation',
  '插入公式': 'Insert Equation', '编辑公式': 'Edit Equation', '输入 LaTeX，例如 x^{2}+\\frac{1}{2}': 'Type LaTeX, e.g. x^{2}+\\frac{1}{2}',
  '分式': 'Fraction', '根式': 'Radical', '积分': 'Integral', '极限': 'Limit', '矩阵': 'Matrix', '分段函数': 'Cases', '二次公式': 'Quadratic Formula',
  '封面': 'Cover Page', '简约': 'Simple', '商务': 'Business', '学术': 'Academic', '在文档开头插入一页封面': 'Insert a cover page at the start of the document',
  '文档标题': 'Document Title', '项目报告': 'Project Report', '作者': 'Author', '单位名称': 'Institution', '论文题目': 'Thesis Title', '作者：': 'Author: ', '日期：': 'Date: ',
  '首字下沉': 'Drop Cap', '下沉': 'Dropped', '悬挂': 'In Margin', '段落的第一个字放大到三行高': 'The paragraph\'s first letter, three lines high', '这一段没有文字可以下沉': 'This paragraph has no text to drop',
  '书签': 'Bookmark', '添加书签…': 'Add Bookmark…', '删除书签': 'Remove Bookmark', '定位': 'Go To', '给所在段落加上书签，交叉引用和链接可以跳到这里': 'Bookmark this paragraph so cross-references and links can go to it',
  '书签名（字母开头，可含数字和下划线）': 'Bookmark name (a letter first, then letters, digits or _)', '书签名须以字母开头，只含字母、数字和下划线': 'A bookmark name starts with a letter and holds only letters, digits and _',
  '已添加书签「{name}」': 'Bookmark "{name}" added', '（空）': '(empty)',

  // shape format tab
  '形状格式': 'Shape Format', '形状填充': 'Shape Fill', '形状轮廓': 'Shape Outline', '无轮廓': 'No Outline', '上下型': 'Top and Bottom',
  '大小…': 'Size…', '宽度（厘米）': 'Width (cm)', '高度（厘米）': 'Height (cm)', '删除形状': 'Delete Shape',

  // review ribbon
  '字数统计': 'Word Count', '新建批注': 'New Comment', '删除所有批注': 'Delete All Comments',
  '修订：开': 'Track Changes: On', '修订：关': 'Track Changes: Off', '开启后，插入与删除会被标记': 'When on, insertions and deletions are marked',
  '接受所有修订': 'Accept All Changes', '拒绝所有修订': 'Reject All Changes',

  // view ribbon
  '页面视图': 'Page View', '分页显示，带页眉页脚': 'Paged, with headers and footers',
  '网页视图': 'Web View', '连续阅读，不分页': 'Continuous, no page breaks', '连续': 'Continuous',
  '缩略图': 'Thumbnails', '页眉页脚': 'Headers & Footers', '点击切换纸张方向': 'Click to switch page orientation',
  '适应宽度': 'Fit Width', '阅读模式': 'Reading Mode', '沉浸阅读 Ctrl+.': 'Immersive Reading Ctrl+.',
  '页面移动': 'Page Movement', '并排': 'Side to Side', '页面上下翻，或并排左右翻': 'Turn pages top to bottom, or side by side left to right',
  '缩放：{pct}': 'Zoom: {pct}', '页宽': 'Page Width', '整页': 'Whole Page', '多页': 'Multiple Pages',

  // table ribbon and table-insert popover
  '上方插入行': 'Insert Row Above', '下方插入行': 'Insert Row Below', '左侧插入列': 'Insert Column Left', '右侧插入列': 'Insert Column Right',
  '删除行': 'Delete Row', '删除列': 'Delete Column', '删除表格': 'Delete Table',
  '合并右侧单元格': 'Merge Right', '合并下方单元格': 'Merge Down', '拆分单元格': 'Split Cells',
  '把合并的单元格拆回一格一格': 'Splits a merged cell back into separate cells',
  '底纹': 'Shading', '单元格底纹': 'Cell Shading', '表格样式': 'Table Style', '网格': 'Grid', '三线表': 'Three-Line Table', '无框线': 'No Borders',
  '简明表格（隔行底纹）': 'Plain Table (banded rows)', '网格表 4（彩色标题行）': 'Grid Table 4 (coloured header row)', '单元格边框': 'Cell Borders', '所有框线': 'All Borders', '随表格': 'As the Table',
  '对齐方式': 'Alignment', '靠上左对齐': 'Top Left', '靠上居中': 'Top Center', '靠上右对齐': 'Top Right', '中部左对齐': 'Middle Left', '中部居中': 'Middle Center', '中部右对齐': 'Middle Right', '靠下左对齐': 'Bottom Left', '靠下居中': 'Bottom Center', '靠下右对齐': 'Bottom Right',
  '行高': 'Row Height', '自动': 'Auto', '标题行重复': 'Repeat Header Row', '所在行在每页顶端重复': 'Repeat this row at the top of every page',
  '排序': 'Sort', '升序': 'Ascending', '降序': 'Descending', '按所在列排序，数字按大小': 'Sort by the caret\'s column; numbers by value', '转换为文本': 'Convert to Text', '表格属性…': 'Table Properties…', '表格属性': 'Table Properties',
  '表格宽度（如 100% 或 12cm）': 'Table width (e.g. 100% or 12cm)', '对齐方式（left / center / right）': 'Alignment (left / center / right)', '边框颜色（如 808080）': 'Border colour (e.g. 808080)',
  '文本转换成表格（按 Tab）': 'Convert Text to Table (at tabs)', '文本转换成表格（按逗号）': 'Convert Text to Table (at commas)', '文本转换成表格（按空格）': 'Convert Text to Table (at spaces)',
  '{r} × {c} 表格': '{r} × {c} Table', '插入表格': 'Insert Table', '自定义行列…': 'Custom Rows & Columns…', '行数': 'Rows', '列数': 'Columns',

  // image tab
  '大小': 'Size', '版心宽度的 {pct}': '{pct} of text width', '位置': 'Position',
  '环绕': 'Wrap', '嵌入型': 'In Line with Text', '四周型': 'Square', '浮于文字上方': 'In Front of Text', '衬于文字下方': 'Behind Text',
  '嵌入行内，或浮在页面上任意拖动': 'In the line of text, or floating anywhere on the page', '删除图片': 'Delete Picture',

  // word count dialog
  '字数': 'Words', '字符数（不计空格）': 'Characters (no spaces)', '段落数': 'Paragraphs', '页数': 'Pages',

  // link dialog
  '插入链接': 'Insert Link', '链接地址': 'Link', '显示文字': 'Display Text', '留空则使用地址': 'Leave blank to use the link',

  // comments panel
  '批注 · {{ commentCount }}': 'Comments · {{ commentCount }}', '删除': 'Delete', '输入批注…': 'Add a comment…',
  '我': 'Me', '先选中要批注的文字': 'Select the text to comment on first',
  '回复…（回车发送）': 'Reply… (Return to send)', '解决': 'Resolve', '重新打开': 'Reopen', '这条批注已解决，点击重新打开': 'This comment is resolved; click to reopen it', '标记为已解决': 'Mark as resolved',
  '显示标记': 'Show Markup', '显示以供审阅': 'Display for Review', '所有标记': 'All Markup', '无标记': 'No Markup', '原始版本': 'Original',
  '所有标记：修订和批注都显示；无标记：显示修改后的样子；原始版本：显示修改前的样子': 'All Markup shows changes and comments; No Markup shows the result; Original shows the text before the changes',

  // toasts
  '已接受所有修订': 'All changes accepted', '已拒绝所有修订': 'All changes rejected',
  '已替换 {n} 处': v => v.n === 1 ? 'Replaced 1 occurrence' : 'Replaced ' + v.n + ' occurrences',
  '目录已更新': 'Table of contents updated', '已在页脚插入页码': 'Page number inserted in the footer', '已保存': 'Saved',

  // header / footer editing bar
  '首页页眉': 'First Page Header', '首页页脚': 'First Page Footer',
  '双击编辑页眉': 'Double-click to edit the header', '双击编辑页脚': 'Double-click to edit the footer',
  '插入页码': 'Insert Page Number', '总页数': 'Total Pages', '插入总页数': 'Insert Total Pages',
  '左': 'Left', '中': 'Center', '右': 'Right',
  '首页不同': 'Different First Page', '首页使用单独的页眉页脚': 'Use separate header and footer for the first page',
  '退出页眉页脚 Esc': 'Exit Header & Footer Esc',

  // status bar
  '共 {n} 页': v => v.n === 1 ? '1 page' : v.n + ' pages', '{n} 字': v => v.n === 1 ? '1 word' : v.n + ' words',
  '共 {n} 处': v => v.n === 1 ? '1 match' : v.n + ' matches', '未找到': 'No matches',

  // the header/footer edit box and its 样式 menu
  '输入文字，点「页码」插入自动变化的页码': 'Type text, or click Page Number to insert a number that updates itself',
  '样式@@hf': 'Format',
  '快速选择页码格式，或直接输入自己想要的形式': 'Pick a page number format, or type your own',
  '在页脚插入页码': 'Insert a page number in the footer',
});
