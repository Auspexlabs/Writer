# 在网页里嵌入 Writer

把 Writer 放进你自己的网页：一个 Word、Excel 或 PowerPoint 编辑器，或者整个 Writer（标签页、新建页、模板都在）。加一个 `<div>`、一行 `<script>` 就能用，和嵌入 TradingView 图表的感觉一样。

- **访问者直接编辑。** 和桌面版是同一套编辑器：Word 分页、Excel 公式和图表、幻灯片、Markdown、思维导图，PDF 可以查看。
- **引擎在浏览器里。** `writer` 引擎编译成了 WebAssembly，文档只在访问者的浏览器内存里，不经过任何服务器。
- **网站可以调用引擎。** 网站的代码用和命令行一样的命令读、改文档，随时取回文件，并在用户改动时收到通知。
- **AI 可以调用引擎。** 组件里的 AI 助手用网站提供的模型，改法和桌面版助手一样（改动可保留或撤销）；网站自己的智能体也能拿到同一套工具，直接改文档。

[ui/embed/example.html](../ui/embed/example.html) 是一个完整的示例页面。

## 快速开始

**1. 构建。** 需要 .NET 10 SDK。

```bash
./build.sh embed      # 输出到 dist/embed/writer/，是一个纯静态文件夹
```

**2. 部署。** 把 `dist/embed/writer/` 整个上传到任意静态托管（Nginx、对象存储、CDN、GitHub Pages 等），比如放在 `https://你的域名/writer/`。想先在本机试一下：

```bash
cd dist/embed && python3 -m http.server 8000
# 打开 http://localhost:8000/writer/embed/example.html
```

**3. 嵌入。**

```html
<div id="doc" style="height:640px"></div>
<script src="https://你的域名/writer/embed/writer-embed.js"></script>
<script>
  const writer = Writer.embed('#doc', { mode: 'docx', file: '/files/方案.docx' });

  writer.ready.then(async () => {
    await writer.run('add 方案.docx /body --type paragraph --prop text="由网站写入的一段"');
    const blob = await writer.get('方案.docx');   // 当前的文件，可以上传回你的服务器
  });
</script>
```

组件是一个 iframe，默认填满目标元素，所以要给目标元素一个高度（或者在选项里给 `height`）。

## 两种形态

| `mode` | 形态 |
|---|---|
| `'docx'` `'xlsx'` `'pptx'` `'md'` `'mm'` `'pdf'` | **单个编辑器**：只有这一份文档，没有标签页和新建页，⌘N、⌘T、⌘W、⌘O 不起作用。没给文件时新建一份空白文档（PDF 除外）。 |
| `'app'`（默认） | **整个软件**：标签页、新建页和模板都在，可以同时打开几份文档，也可以新建。 |

顶部是一条 48 像素的工具条：侧边栏开关、撤销和重做、文件名（单个编辑器）或标签页（整个软件），右边是格式面板、沉浸书写和 AI 助手按钮。

## 选项

`Writer.embed(target, options)`：`target` 是元素或选择器。

| 选项 | 说明 |
|---|---|
| `mode` | 见上表，默认 `'app'` |
| `file` | 一份文档 |
| `files` | 多份文档（数组），和 `file` 可以同时给 |
| `open` | 先显示哪一份（文件名），默认第一份 |
| `theme` | `'light'`（默认）、`'dark'`、`'auto'`（跟随系统） |
| `lang` | `'zh'` 或 `'en'`；不给时跟随浏览器语言 |
| `width` `height` | iframe 的尺寸，数字按像素算，也可以是 CSS 值；默认 `'100%'` |
| `radius` | 圆角，默认 `12` |
| `title` | iframe 的标题（读屏软件会读），默认 `'Writer'` |
| `blankName` | 单个编辑器没给文件时，空白文档的名字，默认「未命名」 |
| `ai` | AI 助手用的模型，见下文「AI 助手」；不给就没有 AI 按钮 |
| `aiModel` | 在助手面板里显示的模型名 |
| `base` | Writer 文件夹的地址。默认是 `writer-embed.js` 所在文件夹的上一级，只有把脚本单独放到别处时才需要 |

文档（`file`、`files` 的每一项，以及 `put` 的内容）可以是：

| 写法 | 说明 |
|---|---|
| `'/files/方案.docx'` 或 `new URL(…)` | 由你的页面下载（同源时带 Cookie），文件名取 URL 的最后一段 |
| `{ name: '方案.docx', url: '/download?id=7' }` | URL 里没有文件名时，用 `name` 指定 |
| `File`、`Blob` | 比如 `<input type="file">` 选中的文件 |
| `ArrayBuffer`、`Uint8Array` | 文件的字节 |
| `{ name: '笔记.md', data: '# 标题' }` | `data` 可以是文字、字节或 Blob |

文件名的扩展名决定用哪个编辑器打开：`.docx` `.xlsx` `.pptx` `.md` `.mm` `.pdf`，以及 `.csv` 等引擎支持的格式。

## API

`Writer.embed` 返回的组件对象：

| 方法 | 说明 |
|---|---|
| `ready` | Promise：组件加载好、第一份文档显示出来时完成，值是 `{ file }`。文档下载失败或引擎没能启动时失败 |
| `run(command)` | 执行一条 writer 命令，返回 `{ code, output }` 或 `{ code, error }`。命令可以是一行文字，也可以是参数数组 |
| `put(name, data)` | 放入或替换一份文档，返回 `{ path, size, mtime }` |
| `get(name)` | 取回文档当前的样子（`Blob`）。用户的修改在一秒内自动存入，取到的是存入后的最新内容 |
| `list()` | 所有文档：`[{ path, name, format, size, modified }]` |
| `open(name)` | 显示某一份文档，成功返回 `true` |
| `current()` | 正在显示的文档名，在新建页时是 `''` |
| `tree(name)` | 文档的完整结构树：`{ type, path, props, children }` |
| `system(file)` `tools()` `callTool(name, input)` | 给网站自己的智能体用，见下文 |
| `setTheme(theme)` | 换成 `'light'`、`'dark'` 或 `'auto'` |
| `on(event, fn)` `off(event, fn)` | 监听事件，见下表 |
| `destroy()` | 移除组件，还在等待的调用会失败 |

调用失败时 Promise 失败，错误对象带 `code`（比如 `FILE_NOT_FOUND`）和 `hint`（怎么改）。所有调用都会先等 `ready`。

| 事件 | 数据 | 什么时候 |
|---|---|---|
| `ready` | `{ file }` | 组件就绪 |
| `open` | `{ file }` | 显示了另一份文档 |
| `change` | `{ files, from }` | 文档被改动。`from` 是改动的来源：`'editor'` 用户编辑，`'assistant'` 组件里的 AI 助手，`'api'` 网站的调用 |

### 命令

`run` 用的是 writer 引擎的命令，和命令行完全一样（只是不带程序名），全部命令、路径语法和属性见 [engine.md](engine.md)。几个例子：

```js
await writer.run('view 方案.docx outline');                          // 每个元素和它的路径
await writer.run(['set', '方案.docx', '/body/paragraph[2]', '--prop', 'text=新的第二段']);
await writer.run('set 预算.xlsx /sheet[1]/cell[B3] --prop value=1200');
await writer.run('export 方案.docx --to 方案.md');                   // 转换出的文件也在组件里，可以用 get 取回
```

大多数命令的 `output` 是 JSON 文字，用 `JSON.parse` 解析。编辑器正在显示的文档被命令改了，会立即刷新。`mcp`、`serve`、`watch`、`app` 这四个启动服务的命令在浏览器里不可用。

### 保存用户的修改

文档在访问者的浏览器内存里，刷新页面就没了。要保存，就在 `change` 时取回文件，传回你的服务器：

```js
let timer;
writer.on('change', ({ files }) => {
  clearTimeout(timer);
  timer = setTimeout(async () => {
    for (const name of files) await fetch('/api/save?name=' + encodeURIComponent(name), { method: 'PUT', body: await writer.get(name) });
  }, 1500);
});
```

## AI 助手

### 组件里的助手

给 `ai` 选项，工具条右边就会出现 AI 按钮。助手和桌面版是同一个：同样的系统提示词、同样的编辑工具（`writer`、`batch`、`plan`），改动在文档里标出来，可以保留或撤销。模型由你的网站提供，Writer 自带两个适配器：

```js
// OpenAI 兼容接口：DeepSeek、通义千问、Kimi、智谱、豆包、vLLM、Ollama 等
Writer.embed('#doc', { mode: 'docx', file: '/files/方案.docx', aiModel: 'deepseek-chat',
  ai: Writer.ai.openai({ url: '/api/llm/chat/completions', model: 'deepseek-chat' }) });

// Anthropic Messages API
Writer.embed('#doc', { mode: 'docx', aiModel: 'claude-sonnet-5',
  ai: Writer.ai.anthropic({ url: '/api/anthropic/v1/messages', model: 'claude-sonnet-5' }) });
```

适配器的选项：`url`（接口地址）、`model`、`headers`（额外的请求头）、`credentials`（默认 `'same-origin'`）；Anthropic 的还有 `maxTokens`（默认 8192）。

**不要把 API Key 写进网页。** 网页里的一切访问者都看得到。`url` 应该指向你自己后端的一个转发地址，由后端加上密钥再转给模型服务商，顺便做登录校验和用量限制。只有自己试用时，才可以直接填模型服务商的地址、在 `headers` 里放密钥（直接调用 `api.anthropic.com` 时，适配器会自动加上浏览器直连需要的请求头）。

也可以自己写 `ai` 函数，接任何模型：

```js
ai: async (request, { onDelta, signal }) => {
  // request: { system, messages, tools }
  //   messages 的格式和模型服务商无关：
  //     { role: 'user', content }
  //     { role: 'assistant', content, toolCalls: [{ id, name, input }] }
  //     { role: 'tool', toolCallId, name, content, isError }
  //   tools: [{ name, description, input_schema }]
  // 回复文字时可以边生成边调用 onDelta(一小段文字)；用户点停止时 signal 会中止
  return { text: '回复的文字', toolCalls: [{ id: 'call_1', name: 'writer', input: { command: 'view 方案.docx outline' } }] };
}
```

助手每轮最多调用模型 24 次，工具在组件里执行，结果自动接到下一次调用。

### 网站自己的智能体

如果网站已经有自己的 AI（比如客服、写作助手），可以让它直接改组件里的文档：

```js
const system = await writer.system('方案.docx');   // 桌面版助手的系统提示词：编辑规则、命令说明、这份文档的大纲
const tools = await writer.tools();                // [{ name, description, input_schema }]：writer、batch、plan
// 把 system 和 tools 交给你的模型；模型要调用工具时：
const result = await writer.callTool('writer', { command: 'set 方案.docx /body/paragraph[1] --prop text=新标题' });
// result: { display, code, output, wrote }，把 output 作为工具结果交还给模型
```

改动会触发 `change` 事件（`from` 是 `'api'`），编辑器里立即显示。

## 部署

- **静态文件。** `dist/embed/writer/` 里只有静态文件，不需要后端。引擎约 4.8 MB（brotli 压缩后），浏览器缓存后再打开很快。
- **MIME 类型。** `.wasm` 文件要以 `application/wasm` 返回，常见的服务器和 CDN 默认如此。每个文件旁边有 `.br` 和 `.gz` 预压缩版本，服务器支持时（比如 Nginx 的 `brotli_static`、`gzip_static`）可以直接用。
- **可以放在别的域名。** Writer 可以放在 CDN 或单独的域名上，和你的网站不同源也没关系：组件是 iframe，和网站之间只通过 `postMessage` 通信。
- **CSP。** 你的网站设置了 CSP 时，`frame-src` 要包含 Writer 所在的地址。查看 PDF、画 Mermaid 图和部分导出功能第一次用到时，会从 `cdn.jsdelivr.net` 和 `esm.sh` 加载对应的库；如果给 Writer 所在的地址也设置了 CSP，要允许这两个地址，否则这些功能不可用。

## 安全与隐私

- 文档只在访问者的浏览器里，Writer 没有任何服务器，也不收集数据。网站传进来的文档、用户的修改，只有网站自己的代码能取走。
- 组件只接受创建它的那个网站发来的消息，也只把结果发给它；网站这边的脚本只接受自己那个 iframe 的消息。
- 组件里的 AI 助手只能改组件里的文档，不能访问网页、网络或访问者的电脑。
- 部署好的 Writer 谁都能嵌入。想只允许自己的网站嵌入，给 `embed.dc.html` 加上 `Content-Security-Policy: frame-ancestors https://你的域名` 响应头。

## 限制

- 每个组件运行一份自己的引擎，一个页面放很多个组件会占用较多内存。
- 需要支持 WebAssembly 和 ES 模块的浏览器（近几年的 Chrome、Edge、Safari、Firefox）。
- 图片的「抠图」和「压缩」要用桌面版附带的本机程序，网页里不显示这两个按钮。

## 工作原理

```
你的网页                                    iframe：writer/embed.dc.html
┌──────────────────────────┐  postMessage  ┌────────────────────────────────────────────┐
│ writer-embed.js          │ ◀───────────▶ │ embed/host.js     协议、初始文档、AI 转发    │
│  Writer.embed(…)         │               │ index.dc.html     和桌面版相同的编辑器        │
│  options.ai → 你的模型    │               │ embed/server.js   拦截编辑器的 fetch 请求    │
└──────────────────────────┘               │ _framework/       writer 引擎（WebAssembly） │
                                           └────────────────────────────────────────────┘
```

编辑器原本通过本地 HTTP 接口（`writer serve`）和引擎通信。在网页里，`embed/server.js` 接管这些请求，交给编译成 WebAssembly 的同一个引擎（`src/Writer.Browser`）处理，所以编辑器本身不需要为网页做任何修改。AI 助手的循环（`embed/ai.js`）和桌面版（`src/Writer.Cli/Chat.cs`）一致，系统提示词和工具来自引擎里的同一份代码（`src/Writer.Cli/Assistant.cs`）。
