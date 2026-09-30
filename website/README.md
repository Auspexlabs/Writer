# Writer 官网

`dist/` 是整个站点：一个 `index.html`，加上 `assets/` 里的样式、脚本、标志、文件图标和截图。没有框架，没有构建步骤。

站点放在中国大陆的服务器上，所以页面不发任何第三方请求：不用 Google Fonts、CDN 或第三方统计脚本，字体用系统自带的（苹方、微软雅黑等）。下载计数只请求本站的接口。改页面时保持这一点：

```bash
grep -rE "https?://" website/dist   # 只应看到 SVG 里的 xmlns 命名空间
```

## 本地预览

```bash
python3 -m http.server 8765 --directory website/dist
```

然后打开 http://127.0.0.1:8765/ 。深色模式跟随系统的外观设置。

## 下载约定

服务器那边由部署脚本实现，页面只依赖下面两个相对地址：

- 下载按钮链接到 `download/mac` 和 `download/windows`，服务器把它们重定向到当前的 `Writer-<version>-mac.dmg` 和 `Writer-<version>-windows-x64-setup.exe`。
- `download/latest.json` 形如 `{"version":"0.1.0","size":20971520,"date":"2026-09-23"}`，`size` 是 dmg 的字节数。`assets/site.js` 读取它，在下载按钮下显示「版本 0.1.0，2026-09-23 更新」；读不到时这一行保持隐藏，按钮照常可用。
- 应用自己的更新器先读 `updates/mac/latest.json` 和 `updates/windows/latest.json`，读不到（连接错误或非 2xx）才退回 GitHub。所以 `updates/` 下缺文件时必须是普通 404，不能配成返回 200 页面或 204，否则更新检查会失效。

仓库里不放 `download/` 目录。本地想看版本行，就临时建一个 `dist/download/latest.json`，看完删掉。

## 下载统计

打开 `https://thewriter.cn/stats/` 即可查看，无需登录，官网页脚也有入口。静态页面显示累计、今日、Mac、Windows 和最近 30 天的每日次数。服务每 30 秒生成一次静态 JSON，页面每 30 秒读取；日期以北京时间计算。

统计口径：官网实际 Mac/Windows 下载按钮的一次点击记一次，重复点击分别计数。支持普通点击、键盘激活和鼠标中键；顶部跳转下载区域、页面浏览、安装包直链、应用更新和下载分片请求不计数。右键菜单另存为、禁用 JavaScript、网络阻断统计请求时无法记录。数字不是独立人数、安装数或成功下载数，启用前的历史无法补记。

`assets/download-count.js` 用 `sendBeacon` 发送平台；无法排队时仅回退一次 `fetch keepalive`，不阻止下载、不重试。服务端仅保存日期、平台和聚合计数，没有访客 ID、Cookie、IP、浏览器信息或单次点击记录。轻量接口不识别独立访客，也不保证抵御伪造点击。

`stats/server.py` 只监听 `127.0.0.1:8787`，由 Caddy 代理。`/api/download-clicks` 接受本站 Origin 的小型 POST。统计服务将汇总原子写入 `/srv/writer/public-stats/data.json`，Caddy 在 `/stats/data.json` 直接提供静态文件，并让浏览器重新验证缓存；旧地址 `/api/download-stats` 也读取同一文件。只有专用 `writer-stats-public` 组可以写入该目录，Caddy 读取公开的汇总文件；数据库仍在 systemd 私有目录中，没有公开。即使计数服务暂时停止，已有静态快照仍可查看，页面显示其更新时间。

`deploy.sh` 会调用 `deploy-stats.sh` 安装并启动 `writer-download-stats.service`。数据库保存在服务器 `/var/lib/writer-download-stats/stats.sqlite3`，与站点文件分开，重新部署和重启保留计数。静态快照在服务启动时立即重新生成，且不会被站点的 rsync 删除。备份运行中的数据库时使用 Python sqlite3 的 `Connection.backup()`，不要单独复制 WAL 模式下的主文件。早期密码文件已不再使用。

验证：`python3 -B -m unittest discover -s website/stats -p 'test_*.py'` 和 `node --test website/stats/*.test.mjs`。本机接口测试可运行 `python3 -B website/stats/server.py --db /tmp/writer-stats-test.sqlite3 --port 8787`。测试点击应使用临时数据库，不往正式统计写入样例。

## 截图

`assets/shots/` 里是真实的应用截图，浅色、深色各一套（WebP，2 倍像素），页面按系统外观自动选择；README 用的 `docs/images/` 是其中浅色的 hero 和四张格式图。界面改版后重拍，以免官网展示旧界面：

```bash
node website/shots/shots.mjs                 # 全部，浅色和深色
node website/shots/shots.mjs hero ai --dark  # 只重拍其中几张、一种外观
```

需要 macOS、Google Chrome 和一个引擎：仓库已用 `dotnet` 构建，或 `WRITER=<writer 可执行文件>`，或本机装好的 `/Applications/Writer.app`（没有 dotnet 时自动使用它的引擎，界面仍取仓库的 `ui/`）。屏幕上不会打开任何窗口。脚本依次：用 `writer` 命令在临时目录里生成示例文档（`make-samples.sh`：咖啡节方案、预算表、演示文稿、思维导图、会议纪要，内容都是虚构的）；启动 `writer serve --no-token --ui ui`；用无头 Chrome 打开 `mac.dc.html?native=1`，通过应用自己的菜单和按钮摆好每个画面（缩略图、AI 面板、适应画布、放映），截下窗口。深色一套同时打开「设置 › 外观 › 深色页面」，页面和 AI 的改动标记才是同一套深色。

两张 AI 对话截图里，模型的回复来自本地的模拟接口 `mock-model.mjs`（面板里显示的模型名只是配置的名字），对文档的修改由引擎真实执行，界面里的改动标记和「保留 / 撤销」也都是应用自己画的。可选环境变量：`SHOTS_TMP`（临时目录）、`KEEP=1`（保留临时目录，里面有示例文档和窗口 PNG 原图）、`CHROME`（Chrome 路径）。

## 其他

- 标志：`assets/writer-logo-light.svg`、`writer-logo-dark.svg` 是应用图标（与 `ui/assets/` 相同），`writer-logo.svg` 是 `brand/` 里的 Writer 字标；文件图标取自 `ui/assets/icons/b/`。
- `lab/` 是内部设计页（图标方案对比），不在 `dist/` 里，不会发布。

## 发布到服务器

```bash
website/deploy.sh                                   # 只更新网站
website/deploy.sh --dmg desktop/dist/Writer-<版本>-mac.dmg --exe desktop/dist/Writer-<版本>-windows-x64-setup.exe --updates desktop/dist/updates
```

每次发版都要和 GitHub Releases 同时跑一遍上面第二条：`--dmg`、`--exe` 把安装包放进 `download/` 并更新 `download/mac`、`download/windows` 的重定向，`--updates` 把 `desktop/dist/updates/mac/` 和 `windows/` 镜像到 `updates/`（manifest 最后传）。网站落后于 GitHub 时，用户会被告知「已是最新版」。

```bash
```

服务器地址和 SSH 私钥写在 `website/.deploy.local`（不提交到仓库）：

```bash
WRITER_HOST=user@host
WRITER_KEY="$HOME/.ssh/<私钥文件>"
```
