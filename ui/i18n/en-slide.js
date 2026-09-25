// English UI text, keyed by the Chinese source string (rules in ui/i18n.js). Loaded only when the UI is in English.
Object.assign(window.I18N_EN = window.I18N_EN || {}, {
  // top bar, thumbnails, zoom bar
  '文件': 'File', '关闭': 'Close', 'AI 助手': 'Assistant',
  '放映': 'Play', '放映@@tab': 'Slide Show', '放映 F5': 'Play F5',
  '已隐藏': 'Hidden', '＋ 新幻灯片': '+ New Slide',
  '没有幻灯片 · 点击左侧「＋ 新幻灯片」': 'No slides · click "+ New Slide" on the left',
  '单击此处添加备注': 'Click here to add notes',
  '＋': '+', '适应窗口': 'Fit to Window', '适应窗口 Ctrl+0': 'Fit to Window Ctrl+0',
  '缩小 Ctrl+−': 'Zoom Out Ctrl+−', '放大 Ctrl+＋': 'Zoom In Ctrl+＋',
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
  '撤销 Ctrl+Z': 'Undo Ctrl+Z', '重做 Ctrl+Y': 'Redo Ctrl+Y',

  // slide management: new/duplicate/delete, layouts (office-io.js LAYOUTS, used at the call site)
  '新建幻灯片': 'New Slide', '版式': 'Layout', '复制@@dup': 'Duplicate', '复制幻灯片': 'Duplicate Slide',
  '删除': 'Delete', '删除幻灯片': 'Delete Slide', '至少保留一页': 'Keep at least one slide',
  '从此页放映': 'Play from This Slide',
  '标题幻灯片': 'Title Slide', '标题和内容': 'Title and Content', '两栏内容': 'Two Content',
  '数据卡片': 'Stat Cards', '节标题': 'Section Header', '仅标题': 'Title Only', '空白': 'Blank',

  // insert tab
  '文本框': 'Text Box', '形状': 'Shape', '表格': 'Table', '日期': 'Date', '页码': 'Page Number',
  '为所有幻灯片添加页码': 'Add page numbers to all slides',
  '{r} × {c} 表格': '{r} × {c} table', '插入表格': 'Insert Table',

  // design tab, theme names (office-io.js THEMES, used at the call site)
  '背景': 'Background', "本页背景颜色": "This slide's background color",
  '背景应用到全部': 'Apply Background to All', '重置背景': 'Reset Background',
  '幻灯片大小': 'Slide Size', '宽屏 16:9': 'Widescreen (16:9)', '标准 4:3': 'Standard (4:3)',
  '墨色': 'Ink', '素白': 'Paper', '海蓝': 'Sea', '陶土': 'Clay',

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

  // status bar
  '幻灯片 {i} / {n}': 'Slide {i} / {n}',
  '{type}  X {x}  Y {y}  宽 {w}  高 {h}{rot}': '{type}  X {x}  Y {y}  W {w}  H {h}{rot}',
  '旋转 {deg}°': 'Rotation {deg}°', '主题「{name}」': 'Theme "{name}"',

  // new-slide/new-object placeholder content (office-io.js, guarded — new slides only, existing ones are untouched)
  '单击添加标题': 'Click to add title', '单击添加副标题': 'Click to add subtitle',
  '单击添加文本': 'Click to add text', '单击输入文本': 'Click to enter text',
  '左栏要点': 'Left column bullet point', '右栏要点': 'Right column bullet point',
  '指标': 'Stat', '标题 {n}': 'Header {n}',
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
});
