// English UI text, keyed by the Chinese source string (rules in ui/i18n.js). Loaded only when the UI is in English.
Object.assign(window.I18N_EN = window.I18N_EN || {}, {
  // top bar, thumbnails, zoom bar
  '文件': 'File', '关闭': 'Close', 'AI 助手': 'Assistant',
  '放映': 'Play', '放映@@tab': 'Slide Show', '放映 F5': 'Play F5',
  '已隐藏': 'Hidden', '＋ 新幻灯片': '+ New Slide',
  '没有幻灯片 · 点击左侧「＋ 新幻灯片」': 'No slides · click "+ New Slide" on the left',
  '单击此处添加备注': 'Click here to add notes',
  '＋': '+', '适应窗口': 'Fit to Window', '适应窗口 Ctrl+0': 'Fit to Window Ctrl+0',
  '缩小 Ctrl+−': 'Zoom Out Ctrl+−', '放大 Ctrl+＋': 'Zoom In Ctrl+＋', // i18n-ok: the full-width plus is the key's
  '输入缩放比例后回车（20–400，100 = 适应窗口）': 'Enter a zoom percentage and press Return (20–400, 100 = fit to window)',

  // ribbon tabs
  '开始': 'Home', '插入': 'Insert', '设计': 'Design', '切换': 'Transitions', '动画': 'Animations',
  '图片': 'Picture', '形状格式': 'Shape Format',

  // home tab: font, paragraph
  '主题字体': 'Theme Font', '思源宋体': 'Source Han Serif', '思源黑体': 'Source Han Sans',
  '字体': 'Font', '字号': 'Size', '增大字号': 'Increase Font Size', '减小字号': 'Decrease Font Size',
  '文字颜色': 'Text Color', '左对齐': 'Align Left', '居中': 'Center', '右对齐': 'Align Right',
  '垂直对齐': 'Vertical Align', '顶端': 'Top', '底端': 'Bottom',
  '• 列表': '• List', '1. 列表': '1. List', '行距': 'Line Spacing',
  '减少缩进 ⇧Tab': 'Decrease Indent ⇧Tab', '增加缩进 Tab': 'Increase Indent Tab', '⇤': '⇤', '⇥': '⇥',
  '段落': 'Paragraph', '段落间距': 'Paragraph Spacing', '段前': 'Before', '段后': 'After', '{n} 磅': '{n} pt',
  '字距': 'Spacing', '字符间距': 'Character Spacing', '紧缩': 'Tight', '常规': 'Normal', '宽松': 'Loose', '很松': 'Very Loose', '极松': 'Extra Loose',
  '分栏': 'Columns', '一栏': 'One Column', '两栏': 'Two Columns', '三栏': 'Three Columns',
  '自动调整': 'Autofit', '不自动调整': 'Do Not Autofit', '溢出时缩排文字': 'Shrink Text on Overflow', '根据文字调整形状大小': 'Resize Shape to Fit Text',
  '文字方向': 'Text Direction', '横排': 'Horizontal', '竖排': 'Vertical', '竖排（东亚）': 'Stacked (East Asian)', '所有文字旋转 270°': 'Rotate All Text 270°',
  '艺术字': 'WordArt', '描边': 'Outline', '渐变填充': 'Gradient Fill', '描边 + 阴影': 'Outline + Shadow', '渐变 + 阴影': 'Gradient + Shadow', '空心描边': 'Hollow Outline',
  '撤销 Ctrl+Z': 'Undo Ctrl+Z', '重做 Ctrl+Y': 'Redo Ctrl+Y',

  // slide management: new/duplicate/delete, layouts (office-io.js LAYOUTS, used at the call site)
  '新建幻灯片': 'New Slide', '版式': 'Layout', '复制@@dup': 'Duplicate', '复制幻灯片': 'Duplicate Slide',
  '删除': 'Delete', '删除幻灯片': 'Delete Slide', '至少保留一页': 'Keep at least one slide',
  '从此页放映': 'Play from This Slide',
  '标题幻灯片': 'Title Slide', '标题和内容': 'Title and Content', '节标题': 'Section Header', '两栏内容': 'Two Content',
  '比较': 'Comparison', '仅标题': 'Title Only', '空白': 'Blank', '内容与标题': 'Content with Caption',
  '图片与标题': 'Picture with Caption', '引用': 'Quote',

  // ✦ 美化: the assistant designs a slide or the deck (the requests it is sent)
  '美化': 'Polish', '让助手美化本页或整份': 'Have the assistant polish this slide or the whole deck',
  '美化本页': 'Polish This Slide', '美化整份': 'Polish the Whole Deck',
  '美化第 {n} 页幻灯片：统一字号层级、对齐与留白，缩小溢出的文字，需要时换版式或配色；只改这一页，不改动文字内容。':
    'Polish slide {n}: unify the type scale, alignment and margins, shrink overflowing text, and change the layout or palette if needed; change only this slide and leave the wording as it is.',
  '美化整份幻灯片：先为全稿定一套配色（palette）和字体（fonts），再逐页统一字号层级、对齐与留白，缩小溢出的文字，需要时换版式；不改动文字内容。':
    'Polish the whole deck: first pick one palette and one pair of fonts for all of it, then go slide by slide to unify the type scale, alignment and margins, shrink overflowing text, and change layouts where needed; leave the wording as it is.',

  // 动画 and 切换 (office-io.js FX / TRANS at the call site)
  '进入': 'Entrance', '退出@@fx': 'Exit', '出现': 'Appear', '浮入': 'Float In', '擦除': 'Wipe', '放大/缩小': 'Grow/Shrink', '陀螺旋': 'Spin', '透明': 'Transparency',
  '消失': 'Disappear', '淡出': 'Fade Out', '飞出': 'Fly Out', '收缩': 'Shrink', '其他效果': 'Other effect', '分割': 'Split', '覆盖': 'Cover', '平滑': 'Morph',
  '单击时': 'On Click', '与上一动画同时': 'With Previous', '上一动画之后': 'After Previous', '动画开始方式': 'Start', '持续时间': 'Duration', '持续时间（秒）': 'Duration (seconds)',
  '延迟': 'Delay', '延迟（秒）': 'Delay (seconds)', '添加进入动画': 'Add an entrance effect', '添加强调动画': 'Add an emphasis effect', '添加退出动画': 'Add an exit effect',
  '动画窗格': 'Animation Pane', '其他对象': 'Other object', '先在动画窗格中选中一个动画': 'Select an effect in the Animation Pane first',
  '选中对象后，从「进入」「强调」「退出」添加动画': 'Select an object, then add an effect from Entrance, Emphasis or Exit',

  '拖动以调整备注窗格的高度': 'Drag to resize the notes pane',
  // sections and 幻灯片浏览
  '默认节': 'Default Section', '无标题节': 'Untitled Section', '新增节': 'Add Section', '删除节': 'Remove Section', '节名': 'Section name', '{n} 张': '{n} slides',
  '幻灯片浏览': 'Slide Sorter', '普通视图': 'Normal',
  // 放映: presenter view (office-io.js presenterHtml), tools and keys
  '演示者视图': 'Presenter View', '循环放映': 'Loop', '放映到最后一张后从头开始，按 Esc 结束': 'Start over after the last slide; Esc ends the show',
  '上一张': 'Previous', '下一张': 'Next', '黑屏': 'Black Screen', '重置计时': 'Reset Timer', '结束放映': 'End Show', '放映结束': 'End of show', '备注@@notes': 'Notes', '跳转到第 {n} 张': 'Go to slide {n}',

  // 表格工具 (office-io.js TABLE_STYLES at the call site)
  '表格工具': 'Table Tools', '上方插入行': 'Insert Above', '下方插入行': 'Insert Below', '左侧插入列': 'Insert Left', '右侧插入列': 'Insert Right',
  '删除行': 'Delete Row', '删除列': 'Delete Column', '合并单元格': 'Merge Cells', '拆分单元格': 'Split Cell',
  '单元格填充': 'Cell Fill', '单元格边框': 'Cell Borders', '表格样式': 'Table Styles', '标题行': 'Header Row', '镶边行': 'Banded Rows', '第一列': 'First Column',
  '中等样式 2': 'Medium Style 2', '浅色样式 1': 'Light Style 1', '浅色样式 2': 'Light Style 2', '深色样式 1': 'Dark Style 1', '网格': 'Table Grid', '无样式': 'No Style',

  // insert tab
  '文本框': 'Text Box', '形状': 'Shape', '表格': 'Table', '日期': 'Date', '页码': 'Page Number',
  '为所有幻灯片添加页码': 'Add page numbers to all slides',
  '{r} × {c} 表格': '{r} × {c} table', '插入表格': 'Insert Table',

  // design tab, theme names (office-io.js THEMES, used at the call site)
  '背景': 'Background', "本页背景颜色": "This slide's background color",
  '背景应用到全部': 'Apply Background to All', '重置背景': 'Reset Background',
  '幻灯片大小': 'Slide Size', '宽屏 16:9': 'Widescreen (16:9)', '标准 4:3': 'Standard (4:3)',
  '墨色': 'Ink', '素白': 'Paper', '海蓝': 'Sea', '陶土': 'Clay', '雾灰': 'Mist', '暖沙': 'Sand', '玫红': 'Rose', '夜蓝': 'Night',

  // transitions + animations (shared words first)
  '无': 'None', '缩放': 'Zoom', '预览': 'Preview', '应用到全部': 'Apply to All',
  '已应用到全部幻灯片': 'Applied to all slides',
  '淡入淡出': 'Fade', '推入': 'Push',
  '淡入': 'Fade In', '上浮': 'Rise Up', '飞入': 'Fly In',
  '先选中一个对象': 'Select an object first', '清除本页动画': 'Clear Animations on This Slide',

  // slide show tab
  '从头开始': 'From Beginning', '从当前幻灯片开始': 'From Current Slide',
  '取消隐藏': 'Unhide Slide', '隐藏幻灯片': 'Hide Slide',
  '隐藏备注': 'Hide Notes', '显示备注': 'Show Notes', '没有可放映的幻灯片': 'No slides to play',

  // object arrange/align, format tab (shape)
  '排列': 'Arrange', '置于顶层': 'Bring to Front', '上移一层': 'Bring Forward',
  '下移一层': 'Send Backward', '置于底层': 'Send to Back',
  '对齐': 'Align', '水平居中': 'Align Center', '顶端对齐': 'Align Top',
  '垂直居中': 'Align Middle', '底端对齐': 'Align Bottom',
  '填充': 'Fill', '填充颜色': 'Fill Color', '主题色': 'Theme Colors', '强调色': 'Accent',
  '卡片色': 'Card', '文字色': 'Text', '无填充': 'No Fill',
  '边框': 'Border', '边框颜色': 'Border Color', '边框粗细': 'Border Weight', '无边框': 'No Border',
  '透明度': 'Transparency', '更改形状': 'Change Shape',
  '旋转': 'Rotate', '向右旋转 90°': 'Rotate Right 90°', '向左旋转 90°': 'Rotate Left 90°', '重置旋转': 'Reset Rotation',

  // shape gallery (office-io.js SHAPE_GALLERY), outline, gradients, groups, exact size
  '基本形状': 'Basic Shapes', '直角三角形': 'Right Triangle', '平行四边形': 'Parallelogram', '梯形': 'Trapezoid', '五边形': 'Pentagon', '八边形': 'Octagon',
  '右箭头': 'Right Arrow', '左箭头': 'Left Arrow', '上箭头': 'Up Arrow', '下箭头': 'Down Arrow', '左右箭头': 'Left-Right Arrow', '燕尾形': 'Chevron',
  '星形与标注': 'Stars and Callouts', '四角星': '4-Point Star', '六角星': '6-Point Star', '矩形标注': 'Rectangular Callout', '圆角矩形标注': 'Rounded Rectangular Callout', '椭圆标注': 'Oval Callout',
  '直线': 'Line', '双箭头': 'Double Arrow', '肘形连接符': 'Elbow Connector', '线型': 'Line Type',
  '实线': 'Solid', '短划线': 'Dash', '长划线': 'Long Dash', '圆点': 'Dot', '划线–点': 'Dash-Dot', '方点划线': 'Square Dot',
  '三角箭头': 'Triangle Arrow', '开放箭头': 'Open Arrow', '燕尾箭头': 'Stealth Arrow', '圆形': 'Oval',
  '线条': 'Line', '线条颜色': 'Line Color', '线条粗细': 'Line Weight', '虚线': 'Dashes', '虚线类型': 'Dash Type',
  '起点箭头': 'Begin Arrow', '终点箭头': 'End Arrow',
  '渐变': 'Gradient', '无渐变': 'No Gradient', '浅 → 深': 'Light → Dark', '深 → 浅': 'Dark → Light', '左 → 右': 'Left → Right', '对角': 'Diagonal', '强调色 → 卡片色': 'Accent → Card',
  '组合': 'Group', '取消组合': 'Ungroup', '先选中两个以上的对象': 'Select two or more objects first', '锁定纵横比': 'Lock Aspect Ratio',
  '宽': 'W', '高': 'H',

  // status bar
  '幻灯片 {i} / {n}': 'Slide {i} / {n}',
  '{type}  X {x}  Y {y}  宽 {w}  高 {h}{rot}': '{type}  X {x}  Y {y}  W {w}  H {h}{rot}',
  '旋转 {deg}°': 'Rotation {deg}°', '主题「{name}」': 'Theme "{name}"',

  // placeholder prompts (office-io.js: drawn over an empty placeholder, never part of its text) and new-object content
  '单击此处添加标题': 'Click to add title', '单击此处添加副标题': 'Click to add subtitle',
  '单击此处添加文本': 'Click to add text', '单击此处添加图片': 'Click to add picture',
  '单击输入文本': 'Click to enter text', '标题 {n}': 'Header {n}',
  "暂不支持 .{ext} 文件": "Doesn't support .{ext} files yet",
  '引擎还没有就绪': "The engine isn't ready yet",

  // picture.js: shape/ratio/border-width/compress option lists
  '矩形': 'Rectangle', '圆角矩形': 'Rounded Rectangle', '椭圆': 'Oval', '三角形': 'Triangle',
  '菱形': 'Diamond', '六边形': 'Hexagon', '五角星': '5-Point Star', '胶囊': 'Pill', '箭头': 'Arrow', '线条': 'Line',
  '1:1 方形': '1:1 Square', '16:9 宽屏': '16:9 Widescreen', '3:4 竖版': '3:4 Portrait', '9:16 竖版': '9:16 Portrait',
  '0.75 磅': '0.75 pt', '1.5 磅': '1.5 pt', '3 磅': '3 pt', '4.5 磅': '4.5 pt', '6 磅': '6 pt',
  '打印（220 ppi）': 'Print (220 ppi)', '网页（150 ppi）': 'Web (150 ppi)', '电子邮件（96 ppi）': 'Email (96 ppi)',

  // picture.js: the 图片 tab's own tools
  '抠图中…': 'Removing Background…', '抠图': 'Remove Background',
  '移除背景，只留下主体（Apple Vision，macOS 14 以上）': 'Removes the background, keeping only the subject (Apple Vision, macOS 14 or later)',
  '裁剪': 'Crop', '拖动裁剪…': 'Drag to Crop…', '拖边角': 'Drag corners',
  '取消裁剪': 'Remove Crop', '裁剪到比例，或拖动边角': 'Crop to a ratio, or drag the corners',
  '裁剪为形状': 'Crop to a shape',
  '旋转 90°': 'Rotate 90°', '水平翻转': 'Flip Horizontal', '垂直翻转': 'Flip Vertical',
  '亮度': 'Brightness', '对比度': 'Contrast', '灰度': 'Grayscale',
  '粗细': 'Weight', '阴影': 'Shadow', '圆角': 'Rounded Corners',
  '压缩中…': 'Compressing…', '压缩图片': 'Compress Picture',
  '按显示尺寸重新编码，并删除裁掉的部分': 'Re-encodes at the size shown and removes the cropped-out parts',
  '重置中…': 'Resetting…', '重置图片': 'Reset Picture', '去掉所有调整，恢复原图': 'Removes every adjustment and restores the original',
  '替换中…': 'Replacing…', '替换图片': 'Replace Picture', '换一张图片，保留宽度与位置': 'Swaps in another picture, keeping the width and position',
  '取消': 'Cancel', '完成': 'Done',
  '已压缩：{before} → {after}，节省 {saved}': 'Compressed: {before} → {after}, saved {saved}',
  "这张图片已经不大于显示所需，没有再压缩": "This picture is already no bigger than it's displayed, so there's nothing to compress",
  '图片没有改成': "The picture wasn't updated",
  // the format panel (the Word 定稿 style)
  'AI 美化本页或整份': 'Have AI beautify this slide or the whole deck',
  '主题色填充': 'Fill with a theme colour',
  '从头开始 F5': 'From Beginning F5',
  '从当前幻灯片开始 Shift+F5': 'From Current Slide Shift+F5',
  '从当前页': 'From Current',
  '切换效果': 'Transition',
  '取消组合 Ctrl+Shift+G': 'Ungroup Ctrl+Shift+G',
  '复制 Ctrl+D': 'Duplicate Ctrl+D',
  '对齐幻灯片或所选对象': 'Align to the slide or the selected objects',
  '幻灯片': 'Slides',
  '开始放映': 'Start Slide Show',
  '放映时跳过这一页': 'Skip this slide in the show',
  '文字溢出时': 'When text overflows',
  '本页背景': 'Slide Background',
  '段前、段后': 'Space before and after',
  '添加动画': 'Add Animation',
  '演示者视图 Alt+F5': 'Presenter View Alt+F5',
  '箭头与线型': 'Arrows & Line Type',
  '组合 Ctrl+G': 'Group Ctrl+G',
  '终点': 'End',
  '置于顶层、上移、下移、置于底层': 'Bring to front, forward, backward, send to back',
  '计时': 'Timing',
  '起点': 'Start',
  '隐藏本页': 'Hide This Slide',
  '复制一份': 'Duplicate', // the slide's right-click menu
});
