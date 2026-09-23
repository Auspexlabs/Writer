# Writer 官网

`dist/` 是整个站点：一个 `index.html`，加上 `assets/` 里的样式、脚本、标志、文件图标和截图。没有框架，没有构建步骤。

站点放在中国大陆的服务器上，所以页面不发任何外部请求：不用 Google Fonts、CDN 或统计脚本，字体用系统自带的（苹方、微软雅黑等）。改页面时保持这一点：

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

- 下载按钮链接到 `download/mac`，服务器把它重定向到当前的 `Writer-<version>-mac.dmg`。
- `download/latest.json` 形如 `{"version":"0.1.0","size":20971520,"date":"2026-09-23"}`，`size` 是字节数。`assets/site.js` 读取它，在两个下载按钮旁显示「版本 0.1.0 · 20 MB · 2026-09-23」；读不到时这一行保持隐藏，按钮照常可用。

仓库里不放 `download/` 目录。本地想看版本行，就临时建一个 `dist/download/latest.json`，看完删掉。

## 截图

`assets/shots/` 里是真实的应用截图，浅色、深色各一套（WebP，最宽 1600 像素），页面按系统外观自动选择。截图用无头 Chrome 打开 `writer app --no-browser` 拍摄；示例文档（咖啡节方案、预算表、演示文稿、思维导图、会议纪要）是用 `writer` 命令生成的。两张 AI 对话截图里，模型的回复来自一个本地的模拟接口，对文档的修改由引擎真实执行，界面里的改动标记和「保留 / 撤销」也都是应用自己画的。

界面改版后需要重拍，以免官网展示旧界面。

## 其他

- 标志：`assets/writer-logo-light.svg`、`writer-logo-dark.svg` 是应用图标（与 `ui/assets/` 相同），`writer-logo.svg` 是 `brand/` 里的 Writer 字标；文件图标取自 `ui/assets/icons/b/`。
- `lab/` 是内部设计页（图标方案对比），不在 `dist/` 里，不会发布。

## 发布到服务器

```bash
website/deploy.sh                                   # 只更新网站
website/deploy.sh --dmg desktop/dist/Writer-<版本>-mac.dmg [--updates desktop/dist/updates]
```

服务器地址和 SSH 私钥写在 `website/.deploy.local`（不提交到仓库）：

```bash
WRITER_HOST=user@host
WRITER_KEY="$HOME/.ssh/<私钥文件>"
```
