# Writer 桌面版、MCP 与插件架构

状态：产品方向已确认，macOS 桌面竖切已实现并在本机验证；Windows 安装包和插件宿主仍需完成。网页编辑器是同一产品的一个入口，正式交付目标包括 macOS 与 Windows 桌面应用。

## 一套引擎，三种使用方式

```text
macOS / Windows 桌面应用 ─┐
浏览器版                  ├─ Writer 引擎 ─ DOCX / XLSX / PPTX / MD / PDF / MM
外部 AI 客户端 ─ MCP ─────┘       │
                                  └─ 插件宿主 ─ 外部 MCP 服务 / 扩展模块
```

文档模型、保存语义、路径、权限边界和历史记录属于引擎；桌面窗口与浏览器只是不同的用户界面入口。桌面应用应随包携带引擎，不要求用户安装 .NET 或手动启动 CLI。

## 当前代码事实

| 能力 | 现状 | 关键位置 |
|---|---|---|
| 跨平台引擎 | .NET 自包含单文件发布脚本包含 `osx-arm64`、`osx-x64`、`win-x64` | `build.sh` |
| 桌面窗口和安装包 | Tauri 宿主已能启动随包引擎、加载现有编辑器、保存 DOCX 并在退出时清理进程；macOS `.app` 本机通过，Windows 尚未实测 | `desktop/`、`src/Writer.Cli/Commands.cs` |
| Writer 作为 MCP 服务 | 已有 stdio 服务，但只暴露一个接受命令字符串的 `writer` 工具 | `src/Writer.Cli/Mcp.cs` |
| 桌面应用连接外部 MCP | 尚无 MCP 客户端、连接配置、工具授权与会话管理 | `src/Writer.Cli/Chat.cs` |
| 引擎格式扩展 | `IFormatAdapter` 存在，但六种适配器由静态列表构造，不能安装扩展 | `src/Writer.Formats/Adapters.cs` |

## 目标边界

1. **桌面宿主。** 复用现有 `ui/` 作为初版编辑视图，加入真正的应用窗口、文件打开/拖放、菜单、最近文件、窗口生命周期和系统集成。推荐 Tauri 2 作为窗口与安装包宿主，将已能自包含发布的 .NET `writer` 作为随包 sidecar；macOS 和 Windows 分别在对应平台构建并实测。Tauri 支持现有前端与外部二进制随包发布：<https://tauri.app/>、<https://v2.tauri.app/develop/sidecar/>。
2. **Writer MCP 服务端。** 保留 `writer mcp` 作为其他 AI 客户端的入口，并逐步把注册表生成的能力以结构化工具与资源暴露；CLI、桌面 UI 和 MCP 必须调用同一引擎操作，不能各自实现一套保存逻辑。
3. **Writer MCP 客户端。** 在引擎进程管理用户配置的本地 stdio 和远程 Streamable HTTP 连接，发现工具/资源/提示，统一呈现在桌面 UI；AI 助手调用外部工具前经过用户可见的授权。现有 C# MCP SDK提供客户端和两类传输：<https://csharp.sdk.modelcontextprotocol.io/v2/concepts/transports/transports.html>。
4. **插件。** 首先把外部 MCP 服务作为可安装、可启停的进程隔离插件，清单声明名称、版本、启动方式、能力与所需访问范围。引擎内的格式适配器保留类型化契约；第三方格式插件再设计独立进程协议，避免依赖运行时装载任意程序集和单文件裁剪行为。
5. **边界。** 插件不能直接获得整个用户目录或未授权文档；桌面宿主负责选择文件和授予范围，Writer 引擎负责执行与记录，MCP 连接负责外部能力。插件故障不得导致编辑器或文档保存进程崩溃。

## 实现顺序

1. 先做可运行的桌面竖切：macOS 窗口启动随包引擎，打开现有编辑器，正常打开/编辑/保存一个文档（已验证）；同一套配置在 Windows 构建并完成同等 smoke（待 Windows 环境）。
2. 把桌面进程的工作区、启动令牌、退出清理和文件选择接入现有本地 HTTP 服务，再生成 `.app` 与 Windows 安装包。
3. 引入 MCP 客户端连接管理与最小插件清单：添加、启停、列出工具、调用一个授权工具。随后接入助手与用户界面。
4. 扩展 `writer mcp` 的结构化发现能力，并定义第三方格式插件协议。每一步用同一份跨入口文档操作测试验证。

验收时要分别证明：桌面窗口与安装包可运行、文档没有丢失、外部 MCP 能访问 Writer、Writer 能连接外部 MCP、插件安装和撤销权限可用。`dotnet build` 和浏览器访问不能代替这些验收。
