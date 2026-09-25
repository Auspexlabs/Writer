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
  '黄色': 'Yellow', '绿色': 'Green', '蓝色': 'Blue', '红色': 'Red', '无': 'None',

  // tabs
  '开始': 'Home', '插入': 'Insert', '布局': 'Layout', '审阅': 'Review', '视图': 'View', '表格': 'Table', '图片': 'Picture',

  // home ribbon
  '撤销 Ctrl+Z': 'Undo Ctrl+Z', '重做 Ctrl+Y': 'Redo Ctrl+Y',
  '格式刷': 'Format Painter', '复制格式后选择目标文字': 'Copy the format, then select text to apply it', '清除格式': 'Clear Formatting',
  '字体': 'Font', '字号': 'Size', '增大字号 Ctrl+Shift+>': 'Grow Font Ctrl+Shift+>', '减小字号 Ctrl+Shift+<': 'Shrink Font Ctrl+Shift+<',
  '加粗 Ctrl+B': 'Bold Ctrl+B', '斜体 Ctrl+I': 'Italic Ctrl+I', '下划线 Ctrl+U': 'Underline Ctrl+U', '删除线': 'Strikethrough',
  '上标': 'Superscript', '下标': 'Subscript', '文字颜色': 'Text Color', '高亮': 'Highlight', '段落样式': 'Paragraph Style',
  '左对齐': 'Align Left', '居中': 'Center', '右对齐': 'Align Right', '两端对齐': 'Justify',
  '• 列表': '• List', '1. 列表': '1. List', '减少缩进': 'Decrease Indent', '增加缩进': 'Increase Indent', '行距': 'Line Spacing',
  '查找替换': 'Find & Replace',

  // insert ribbon
  '链接': 'Link', '分隔线': 'Horizontal Line', '分页符': 'Page Break',
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

  // review ribbon
  '字数统计': 'Word Count', '新建批注': 'New Comment', '删除所有批注': 'Delete All Comments',
  '修订：开': 'Track Changes: On', '修订：关': 'Track Changes: Off', '开启后，插入与删除会被标记': 'When on, insertions and deletions are marked',
  '接受所有修订': 'Accept All Changes', '拒绝所有修订': 'Reject All Changes',

  // view ribbon
  '页面视图': 'Page View', '分页显示，带页眉页脚': 'Paged, with headers and footers',
  '网页视图': 'Web View', '连续阅读，不分页': 'Continuous, no page breaks', '连续': 'Continuous',
  '缩略图': 'Thumbnails', '页眉页脚': 'Headers & Footers', '点击切换纸张方向': 'Click to switch page orientation',
  '适应宽度': 'Fit Width', '阅读模式': 'Reading Mode', '沉浸阅读 Ctrl+.': 'Immersive Reading Ctrl+.',

  // table ribbon and table-insert popover
  '上方插入行': 'Insert Row Above', '下方插入行': 'Insert Row Below', '左侧插入列': 'Insert Column Left', '右侧插入列': 'Insert Column Right',
  '删除行': 'Delete Row', '删除列': 'Delete Column', '删除表格': 'Delete Table',
  '合并右侧单元格': 'Merge Right', '合并下方单元格': 'Merge Down', '拆分单元格': 'Split Cells',
  '把合并的单元格拆回一格一格': 'Splits a merged cell back into separate cells',
  '底纹': 'Shading', '单元格底纹': 'Cell Shading', '表格样式': 'Table Style', '网格': 'Grid', '三线表': 'Three-Line Table', '无框线': 'No Borders',
  '{r} × {c} 表格': '{r} × {c} Table', '插入表格': 'Insert Table', '自定义行列…': 'Custom Rows & Columns…', '行数': 'Rows', '列数': 'Columns',

  // image tab
  '大小': 'Size', '版心宽度的 {pct}': '{pct} of text width', '位置': 'Position',
  '靠左': 'Left', '靠右': 'Right', '文字环绕': 'Wrap Text', '删除图片': 'Delete Picture',

  // word count dialog
  '字数': 'Words', '字符数（不计空格）': 'Characters (no spaces)', '段落数': 'Paragraphs', '页数': 'Pages',

  // link dialog
  '插入链接': 'Insert Link', '链接地址': 'Link', '显示文字': 'Display Text', '留空则使用地址': 'Leave blank to use the link',

  // comments panel
  '批注 · {{ commentCount }}': 'Comments · {{ commentCount }}', '删除': 'Delete', '输入批注…': 'Add a comment…',
  '我': 'Me', '先选中要批注的文字': 'Select the text to comment on first',

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
