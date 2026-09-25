// Writer for Mac: one window per folder, each served by its own bundled engine (`writer app` sidecar) and showing
// ui/mac.dc.html. The native pieces the design imitates live here: the menu bar, the settings / about / gestures
// windows with vibrancy, opening files from Finder, the Dock and the open panel in place, recent files, and settings
// shared by every window.

use std::collections::HashMap;
use std::fs;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, AtomicUsize, Ordering};
use std::sync::Mutex;

use serde_json::{json, Map, Value};
use tauri::ipc::{InvokeBody, Request};
use tauri::menu::{
    CheckMenuItem, CheckMenuItemBuilder, Menu, MenuItem, MenuItemBuilder, PredefinedMenuItem, Submenu,
    SubmenuBuilder,
};
#[cfg(not(feature = "appstore"))]
use tauri::utils::config::WindowEffectsConfig;
#[cfg(not(feature = "appstore"))]
use tauri::window::{Effect, EffectState};
use tauri::{
    AppHandle, DragDropEvent, Manager, RunEvent, Theme, TitleBarStyle, WebviewUrl, WebviewWindow,
    WebviewWindowBuilder, WindowEvent, Wry,
};
use tauri_plugin_dialog::DialogExt;
use tauri_plugin_shell::process::{CommandChild, CommandEvent};
use tauri_plugin_shell::ShellExt;
use tauri_plugin_window_state::{StateFlags, WindowExt};

mod bookmarks;
mod storage;
mod update;
mod win;

const DOC_EXTS: [&str; 7] = ["docx", "xlsx", "pptx", "md", "markdown", "mm", "pdf"];
/// localStorage keys every window shares (settings, the collapsed toolbar, the sidebar/assistant panel); kept in
/// <app data>/web-storage.json.
const SHARED_KEYS: [&str; 2] = ["writer-settings", "writer-mac"];
const RECENT_MAX: usize = 10;
/// The design draws its own traffic lights (12px, 8px apart; ui/mac.dc.html and the Mac panels). macOS 27's are 14px
/// and 23px apart, so no inset can match them: the native buttons stay hidden on every window, also after resizing
/// and full screen, and the pages' lights call the window plugin (close, minimize, set_fullscreen, toggle_maximize).
fn hide_lights(window: &WebviewWindow) {
    #[cfg(target_os = "macos")]
    let _ = window.with_webview(|webview| {
        use objc2_app_kit::{NSWindow, NSWindowButton};
        // SAFETY: with_webview runs on the main thread while the window is alive, and ns_window() is its NSWindow.
        let window = unsafe { &*(webview.ns_window() as *const NSWindow) };
        for kind in [NSWindowButton::CloseButton, NSWindowButton::MiniaturizeButton, NSWindowButton::ZoomButton] {
            if let Some(button) = window.standardWindowButton(kind) {
                button.setHidden(true);
            }
        }
    });
}

/// A document window and the engine serving its folder.
struct Folder {
    dir: PathBuf,
    origin: Option<String>,
    engine: Option<CommandChild>,
    /// The page has reported its state once, so window.__writerOpen exists and the listing is loaded.
    ready: bool,
    queue: Vec<PathBuf>,
    state: Value,
}

#[derive(Default)]
struct Shell {
    windows: HashMap<String, Folder>,
    count: u32,
    /// The document window the menu bar acts on: the focused one, or the last one that was.
    front: Option<String>,
    launched: bool,
    /// Files macOS asked to open before the app finished launching.
    launch_files: Vec<PathBuf>,
    /// A settings / about / gestures window asked for while no engine was running yet.
    panel: Option<(String, Option<String>)>,
    store: Map<String, Value>,
    recent: Vec<PathBuf>,
}

struct Writer(Mutex<Shell>);

/// Menu items whose check mark or title follows the front window.
struct Menus {
    sidebar: CheckMenuItem<Wry>,
    toolbar: MenuItem<Wry>,
    dark: MenuItem<Wry>,
    recent: Submenu<Wry>,
}

fn shell(app: &AppHandle) -> std::sync::MutexGuard<'_, Shell> {
    app.state::<Writer>().inner().0.lock().unwrap_or_else(|e| e.into_inner())
}

fn data_file(app: &AppHandle, name: &str) -> PathBuf {
    app.path().app_data_dir().unwrap_or_else(|_| std::env::temp_dir().join("Writer")).join(name)
}

fn read_json(path: &Path) -> Value {
    fs::read(path).ok().and_then(|b| serde_json::from_slice(&b).ok()).unwrap_or(Value::Null)
}

fn write_json(path: &Path, value: &Value) {
    if let Some(dir) = path.parent() {
        let _ = fs::create_dir_all(dir);
    }
    let _ = fs::write(path, serde_json::to_vec_pretty(value).unwrap_or_default());
}

fn is_doc(path: &Path) -> bool {
    path.extension()
        .and_then(|e| e.to_str())
        .is_some_and(|e| DOC_EXTS.contains(&e.to_ascii_lowercase().as_str()))
}

/// The window a plain launch opens: the drafts folder, where new documents live until their first 存储. There is no
/// default documents folder.
fn drafts_folder(app: &AppHandle) -> PathBuf {
    let dir = data_file(app, "Drafts");
    let _ = fs::create_dir_all(&dir);
    dir.canonicalize().map(win::plain_path).unwrap_or(dir)
}

fn show_error(window: &WebviewWindow, message: &str) {
    if let Ok(value) = serde_json::to_string(message) {
        let _ = window.eval(format!("document.getElementById('status').textContent = {value};"));
    }
}

/// Every window gets the shared localStorage keys and the native glue (native.js) before its page runs, plus the
/// system's preferred language for ui/i18n.js (WKWebView's navigator.language reports the app's own localization);
/// __WRITER_NATIVE__.lang is the language this launch chose, for the splash page; uiBuild is the build of the interface
/// package this run serves, for 关于 Writer (update.rs).
fn init_script(app: &AppHandle, kind: &str) -> String {
    let store = shell(app).store.clone();
    let lang = sys_locale::get_locale().unwrap_or_default();
    let native = json!({
        "kind": kind, "store": store, "opaque": cfg!(feature = "appstore"), "updater": !cfg!(feature = "appstore"),
        "uiBuild": update::ui_package(app).map(|(_, build)| build), "lang": if ENGLISH.load(Ordering::Relaxed) { "en" } else { "zh" }
    });
    format!("window.__WRITER_NATIVE__ = {native};\nwindow.__WRITER_SYS_LANG = {};\n{}", json!(lang), include_str!("native.js"))
}

// ---------------------------------------------------------------- language

/// The menus, dialogs and window titles here are English when the UI is: decided at launch by the rule ui/i18n.js
/// uses for the pages (设置 › 语言 简体中文 or English, else the system's language, zh… being Chinese), so a change
/// applies after a restart (restart_app).
static ENGLISH: AtomicBool = AtomicBool::new(false);

fn english(lang: &str, system: &str) -> bool {
    match lang {
        "English" => true,
        "简体中文" => false,
        _ => !system.to_ascii_lowercase().starts_with("zh"),
    }
}

/// AppKit's own menu items and panels (服务, 开始听写, the Open and Save panels) follow AppleLanguages in Writer's
/// defaults, which macOS reads as the app starts: kept in step with 设置 › 语言 (at launch and on every change), so
/// they match the pages from the next launch, as the pages do. 跟随系统 leaves the system's order. Info.plist lists both.
#[cfg(target_os = "macos")]
fn app_languages(lang: &str) {
    use objc2_foundation::{NSArray, NSString, NSUserDefaults};
    let (defaults, key) = (NSUserDefaults::standardUserDefaults(), NSString::from_str("AppleLanguages"));
    match lang {
        "English" | "简体中文" => {
            let list = NSArray::from_retained_slice(&[NSString::from_str(if lang == "English" { "en" } else { "zh-Hans" })]);
            // SAFETY: AppleLanguages holds an array of language codes
            unsafe { defaults.setObject_forKey(Some(&***list), &key) };
        }
        _ => defaults.removeObjectForKey(&key),
    }
}

/// 设置 › 语言 from the stored 'writer-settings' ('' when never set: 跟随系统).
fn lang_setting(store: &Map<String, Value>) -> String {
    let settings = store.get("writer-settings").and_then(Value::as_str).and_then(|s| serde_json::from_str::<Value>(s).ok());
    settings.as_ref().and_then(|v| v.get("lang")).and_then(Value::as_str).unwrap_or_default().to_string()
}

/// Rust-side text in the UI language, keyed by the Chinese as in ui/i18n/en-*.js.
fn t(zh: &'static str) -> &'static str {
    if !ENGLISH.load(Ordering::Relaxed) {
        return zh;
    }
    match zh {
        "关于 Writer" => "About Writer",
        "设置…" => "Settings…",
        "设置" => "Settings",
        "服务" => "Services",
        "隐藏 Writer" => "Hide Writer",
        "隐藏其他" => "Hide Others",
        "全部显示" => "Show All",
        "退出 Writer" => "Quit Writer",
        "文件" => "File",
        "新建…" => "New…",
        "新建标签页" => "New Tab",
        "打开…" => "Open…",
        "打开最近使用" => "Open Recent",
        "清除菜单" => "Clear Menu",
        "关闭标签页" => "Close Tab",
        "关闭窗口" => "Close Window",
        "存储" => "Save",
        "另存为…" => "Save As…",
        "导出为…" => "Export As…",
        "打印…" => "Print…",
        "编辑" => "Edit",
        "撤销" => "Undo",
        "重做" => "Redo",
        "剪切" => "Cut",
        "拷贝" => "Copy",
        "粘贴" => "Paste",
        "全选" => "Select All",
        "查找…" => "Find…",
        "替换…" => "Replace…",
        "插入" => "Insert",
        "表格" => "Table",
        "图片…" => "Picture…",
        "链接" => "Link",
        "批注" => "Comment",
        "分页符" => "Page Break",
        "格式" => "Format",
        "字体" => "Font",
        "粗体" => "Bold",
        "斜体" => "Italic",
        "下划线" => "Underline",
        "段落样式" => "Paragraph Styles",
        "清除格式" => "Clear Formatting",
        "显示" => "View",
        "显示边栏" => "Show Sidebar",
        "收起工具栏" => "Collapse Toolbar",
        "展开工具栏" => "Expand Toolbar",
        "✦ AI 助手" => "✦ Assistant",
        "沉浸书写" => "Focus Mode",
        "深色模式" => "Dark Mode",
        "浅色模式" => "Light Mode",
        "进入全屏幕" => "Enter Full Screen",
        "窗口" => "Window",
        "最小化" => "Minimize",
        "缩放" => "Zoom",
        "显示上一个标签页" => "Show Previous Tab",
        "显示下一个标签页" => "Show Next Tab",
        "将所有窗口移到前面" => "Bring All to Front",
        "帮助" => "Help",
        "Writer 帮助" => "Writer Help",
        "新功能" => "What’s New",
        "触控板手势" => "Trackpad Gestures",
        "键盘快捷键" => "Keyboard Shortcuts",
        "反馈问题…" => "Report an Issue…",
        "打开" => "Open",
        "Writer 文稿" => "Writer Documents",
        "缺少 Writer 界面文件，请重新安装。" => "Writer’s interface files are missing. Please reinstall Writer.",
        "文档引擎未能启动：" => "The document engine couldn’t start: ",
        "无法打开编辑器：" => "Couldn’t open the editor: ",
        "文档引擎启动失败，请检查本地日志。" => "The document engine failed to start. Check the local logs.",
        "文档引擎未能启动。请重新启动 Writer。" => "The document engine couldn’t start. Please restart Writer.",
        // the updater (src/update.rs)
        "检查更新…" => "Check for Updates…",
        "新版本已下载" => "Update Downloaded",
        "重新打开 Writer 后生效。" => "It takes effect the next time you open Writer.",
        "立即重启" => "Restart Now",
        "已是最新版本" => "You’re Up to Date",
        "Writer {v} 是目前的最新版本。" => "Writer {v} is the latest version.",
        "无法检查更新" => "Couldn’t Check for Updates",
        "更新失败" => "Update Failed",
        "请稍后再试，或前往 https://github.com/Auspexlabs/writer/releases 下载最新版本。" => "Try again later, or download the latest version from https://github.com/Auspexlabs/writer/releases.",
        "前往下载页" => "Open Download Page",
        "好" => "OK",
        _ => zh,
    }
}

fn theme_of(store: &Map<String, Value>) -> Option<Theme> {
    let settings: Value = store.get("writer-settings").and_then(Value::as_str).and_then(|s| serde_json::from_str(s).ok())?;
    match settings.get("theme").and_then(Value::as_str) {
        Some("dark") => Some(Theme::Dark),
        Some("light") | None => Some(Theme::Light),
        _ => None, // auto: follow macOS
    }
}

// ---------------------------------------------------------------- document windows

/// Opens (or focuses) the window for the file's folder and opens the file's tab there, editing it in place.
fn open_file(app: &AppHandle, file: PathBuf) {
    if !is_doc(&file) || !file.is_file() {
        return;
    }
    let file = file.canonicalize().map(win::plain_path).unwrap_or(file);
    let Some(dir) = file.parent().map(Path::to_path_buf) else { return };
    note_recent(app, &file);
    let existing = shell(app).windows.iter().find(|(_, f)| f.dir == dir).map(|(label, _)| label.clone());
    if let Some(window) = existing.and_then(|label| app.get_webview_window(&label)) {
        let _ = window.unminimize();
        let _ = window.set_focus();
        grant(app, window.label(), "list", &file);
        send_open(app, &window, file);
    } else if let Some(label) = open_folder(app, dir, 0) {
        grant(app, &label, "list", &file);
        if let Some(folder) = shell(app).windows.get_mut(&label) {
            folder.queue.push(file);
        }
    }
}

/// Asks the page to open the file, or keeps it until the page has loaded the folder's listing.
fn send_open(app: &AppHandle, window: &WebviewWindow, file: PathBuf) {
    let ready = {
        let mut s = shell(app);
        let Some(folder) = s.windows.get_mut(window.label()) else { return };
        if !folder.ready {
            folder.queue.push(file.clone());
        }
        folder.ready
    };
    if ready {
        let path = json!(file.to_string_lossy());
        let _ = window.eval(format!("window.__writerOpen && window.__writerOpen({path})"));
    }
}

/// A new document window on `dir`, listing `depth` folder levels; returns its label.
fn open_folder(app: &AppHandle, dir: PathBuf, depth: u32) -> Option<String> {
    let (label, near) = {
        let mut s = shell(app);
        s.count += 1;
        (format!("main-{}", s.count), s.front.clone())
    };
    let mut config = app.config().app.windows.first()?.clone();
    config.label = label.clone();
    let mut builder = WebviewWindowBuilder::from_config(app, &config)
        .ok()?
        .title(dir.file_name().map(|n| n.to_string_lossy().into_owned()).unwrap_or_else(|| "Writer".into()))
        .initialization_script(init_script(app, "main"));
    let cascaded = near.and_then(|l| app.get_webview_window(&l));
    if let Some(prev) = &cascaded {
        if let (Ok(pos), Ok(scale)) = (prev.outer_position(), prev.scale_factor()) {
            let pos = pos.to_logical::<f64>(scale);
            builder = builder.position(pos.x + 26.0, pos.y + 26.0);
        }
    }
    let window = builder.build().ok()?;
    // Opened next to another window: keep the 26px cascade above. Otherwise (the first window this run), come back
    // where the document windows were last time (tauri-plugin-window-state tracks them as they move and resize).
    if cascaded.is_none() {
        let _ = window.restore_state(StateFlags::SIZE | StateFlags::POSITION);
    }
    hide_lights(&window);
    shell(app).windows.insert(
        label.clone(),
        Folder { dir: dir.clone(), origin: None, engine: None, ready: false, queue: vec![], state: Value::Null },
    );
    start_engine(app, window, dir, depth);
    Some(label)
}

/// The interface every window and panel of this run shows: a downloaded interface package for this version (chosen once,
/// at launch: update::ui_package), else the bundled ui/.
fn ui_dir(app: &AppHandle) -> Option<PathBuf> {
    if let Some((package, _)) = update::ui_package(app) {
        return Some(package.clone());
    }
    let bundled = app.path().resource_dir().ok()?.join("ui");
    if bundled.join("mac.dc.html").is_file() {
        Some(bundled)
    } else if cfg!(debug_assertions) {
        Some(PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../ui"))
    } else {
        None
    }
}

/// The drafts folder keeps its port between launches, so its origin, cache and cookie stay warm. The engine itself
/// finds out whether the port is still free (a socket opened here would make macOS ask for local-network access).
fn remembered_port(app: &AppHandle, dir: &Path) -> Option<u16> {
    if *dir != drafts_folder(app) {
        return None;
    }
    read_json(&data_file(app, "engine.json")).get("port")?.as_u64().and_then(|p| u16::try_from(p).ok())
}

fn start_engine(app: &AppHandle, window: WebviewWindow, dir: PathBuf, depth: u32) {
    let port = remembered_port(app, &dir);
    spawn_engine(app, window, dir, depth, port);
}

fn spawn_engine(app: &AppHandle, window: WebviewWindow, dir: PathBuf, depth: u32, port: Option<u16>) {
    let Some(ui) = ui_dir(app) else {
        show_error(&window, t("缺少 Writer 界面文件，请重新安装。"));
        return;
    };
    let depth_arg = depth.to_string();
    let port_arg = port.unwrap_or(0).to_string();
    let spawned = app.shell().sidecar("writer").and_then(|cmd| {
        cmd.args(["app", "--no-browser", "--list-depth", &depth_arg, "--port", &port_arg, "--dir"])
            .arg(&dir)
            .arg("--ui")
            .arg(&ui)
            .arg("--drafts")
            .arg(drafts_folder(app))
            .arg("--grants-stdin")
            .spawn()
    });
    let (mut events, child) = match spawned {
        Ok(started) => started,
        Err(error) => {
            show_error(&window, &format!("{}{error}", t("文档引擎未能启动：")));
            return;
        }
    };
    if let Some(folder) = shell(app).windows.get_mut(window.label()) {
        folder.engine = Some(child);
    }
    let app = app.clone();
    let remember = dir == drafts_folder(&app);
    tauri::async_runtime::spawn(async move {
        let mut buffer = String::new();
        let mut ready = false;
        while let Some(event) = events.recv().await {
            match event {
                CommandEvent::Stderr(bytes) if !ready => {
                    buffer.push_str(&String::from_utf8_lossy(&bytes));
                    if let Some(url) = page_url(buffer.trim()) {
                        ready = true;
                        let origin = url.origin().ascii_serialization();
                        if remember {
                            write_json(&data_file(&app, "engine.json"), &json!({ "port": url.port() }));
                        }
                        if let Some(folder) = shell(&app).windows.get_mut(window.label()) {
                            folder.origin = Some(origin);
                        }
                        if let Err(error) = window.navigate(url) {
                            show_error(&window, &format!("{}{error}", t("无法打开编辑器：")));
                        }
                        if update::ui_package(&app).is_some() {
                            // a downloaded interface that does not report in (shell_state) within 20 s is marked bad,
                            // and Writer restarts on the bundled one
                            let (app, label) = (app.clone(), window.label().to_string());
                            std::thread::spawn(move || {
                                std::thread::sleep(std::time::Duration::from_secs(20));
                                if shell(&app).windows.get(&label).is_some_and(|f| f.state.is_null()) {
                                    update::reject_package(&app);
                                    app.request_restart();
                                }
                            });
                        }
                        let panel = shell(&app).panel.take();
                        if let Some((which, tab)) = panel {
                            open_panel(&app, &which, tab);
                        }
                        continue;
                    }
                    if buffer.len() > 8192 {
                        show_error(&window, t("文档引擎启动失败，请检查本地日志。"));
                        buffer.clear();
                    }
                }
                CommandEvent::Terminated(_) if !ready => {
                    if port.is_some() {
                        spawn_engine(&app, window, dir, depth, None); // the remembered port is taken: any free one
                    } else {
                        show_error(&window, t("文档引擎未能启动。请重新启动 Writer。"));
                    }
                    break;
                }
                _ => {}
            }
        }
    });
}

/// Tells this window's engine about a file: "allow" lets it read and write one the user chose in the Save dialog; "list"
/// names one opened in the window, which its listing then shows even where the folder cannot be read (the App Store
/// sandbox grants the file, not its folder, and the engines share the app's sandbox).
fn grant(app: &AppHandle, label: &str, verb: &str, path: &Path) -> bool {
    let Some(line) = storage::grant_line(verb, path) else { return false };
    let mut s = shell(app);
    let Some(engine) = s.windows.get_mut(label).and_then(|f| f.engine.as_mut()) else { return false };
    engine.write(line.as_bytes()).is_ok()
}

/// The engine prints {"url": "http://127.0.0.1:<port>/app/index.dc.html?auth=<code>"}; the window page is mac.dc.html
/// (win.dc.html on Windows, whose frameless window it draws the title bar for).
fn page_url(line: &str) -> Option<tauri::Url> {
    let value: Value = serde_json::from_str(line).ok()?;
    let mut url = tauri::Url::parse(value.get("url")?.as_str()?).ok()?;
    if url.scheme() != "http" || url.host_str() != Some("127.0.0.1") || url.path() != "/app/index.dc.html" {
        return None;
    }
    let auth = url.query_pairs().find(|(k, _)| k == "auth").map(|(_, v)| v.into_owned());
    url.set_path(if cfg!(target_os = "windows") { "/app/win.dc.html" } else { "/app/mac.dc.html" });
    url.set_query(None);
    url.query_pairs_mut().append_pair("native", "1");
    // WRITER_PERF=1|auto: frame-time HUD / scripted measurements (ui/perf.js) inside the release webview.
    if let Some(perf) = std::env::var_os("WRITER_PERF") {
        url.query_pairs_mut().append_pair("perf", &perf.to_string_lossy());
    }
    if let Some(auth) = auth {
        url.query_pairs_mut().append_pair("auth", &auth);
    }
    Some(url)
}

// ---------------------------------------------------------------- settings / about / gestures windows

fn open_panel(app: &AppHandle, which: &str, tab: Option<String>) {
    let (page, title, width, height) = match which {
        "settings" => ("MacSettings", t("设置"), 660.0, 520.0),
        "about" => ("MacAbout", t("关于 Writer"), 300.0, 380.0),
        "gestures" => ("MacGestures", t("触控板手势"), 560.0, 560.0),
        _ => return,
    };
    let tab = tab.filter(|t| !t.is_empty() && t.chars().all(|c| c.is_ascii_alphanumeric()));
    if let Some(window) = app.get_webview_window(which) {
        if let Some(tab) = &tab {
            let _ = window.eval(format!("window.__writerTab && window.__writerTab({})", json!(tab)));
        }
        let _ = window.unminimize();
        let _ = window.set_focus();
        return;
    }
    let origin = {
        let s = shell(app);
        let front = s.front.as_ref().and_then(|l| s.windows.get(l)).and_then(|f| f.origin.clone());
        front.or_else(|| s.windows.values().find_map(|f| f.origin.clone()))
    };
    let Some(origin) = origin else {
        // No engine is serving pages yet: open one and come back when it is up.
        let idle = shell(app).windows.is_empty();
        shell(app).panel = Some((which.to_string(), tab));
        if idle {
            open_folder(app, drafts_folder(app), 3);
        }
        return;
    };
    let mut url = format!("{origin}/app/{page}.dc.html?native=1");
    if let Some(tab) = &tab {
        url.push_str(&format!("&tab={tab}"));
    }
    let Ok(url) = url.parse() else { return };
    let built = WebviewWindowBuilder::new(app, which, WebviewUrl::External(url))
        .title(title)
        .inner_size(width, height)
        .resizable(false)
        .minimizable(false)
        .maximizable(false)
        .title_bar_style(TitleBarStyle::Overlay)
        .hidden_title(true)
        .initialization_script(init_script(app, which))
        .center();
    // A transparent window over macOS vibrancy needs Tauri's private-API feature, which the Mac App Store forbids;
    // the store build keeps the panels opaque and their pages draw the design's glass themselves.
    #[cfg(not(feature = "appstore"))]
    let built = built.transparent(true).effects(WindowEffectsConfig {
        effects: vec![Effect::Popover],
        state: Some(EffectState::Active),
        radius: None,
        color: None,
    });
    let built = built.build();
    if let Ok(window) = built {
        hide_lights(&window);
    }
}

fn is_panel(label: &str) -> bool {
    matches!(label, "settings" | "about" | "gestures")
}

// ---------------------------------------------------------------- commands (capabilities/writer-mac-shell.json)

/// The page's state: window title, the menu's check marks and titles, the active document for 打开最近使用.
#[tauri::command]
fn shell_state(app: AppHandle, window: WebviewWindow, request: Request<'_>) {
    let InvokeBody::Json(state) = request.body() else { return };
    #[cfg(feature = "test-hooks")]
    eprintln!("shell_state {} {}", window.label(), state);
    // On load the shell opens the folder's newest document by itself; files asked for wait until that has happened,
    // or the shell's own open could finish last and take the front tab (seen with a slow .xlsx).
    let settled = state.get("path").is_some_and(|p| !p.is_null()) || state.get("docs").and_then(Value::as_u64) == Some(0);
    let (dir, front, first) = {
        let mut s = shell(&app);
        let front = s.front.is_none() || s.front.as_deref() == Some(window.label());
        let Some(folder) = s.windows.get_mut(window.label()) else { return };
        let first = folder.state.is_null();
        folder.state = state.clone();
        (folder.dir.clone(), front, first)
    };
    let title = state.get("title").and_then(Value::as_str).filter(|t| !t.is_empty()).unwrap_or("Writer");
    let _ = window.set_title(title);
    if let Some(path) = state.get("path").and_then(Value::as_str).filter(|p| !p.is_empty()) {
        note_recent(&app, &dir.join(path));
    }
    if front {
        apply_menu(&app, state);
    }
    if settled {
        flush_opens(&app, &window);
    } else if first {
        // the shell's first document may fail to open: do not wait for it forever
        let (app, window) = (app.clone(), window.clone());
        std::thread::spawn(move || {
            std::thread::sleep(std::time::Duration::from_secs(3));
            flush_opens(&app, &window);
        });
    }
}

/// The page is ready for window.__writerOpen: open what was asked for while it loaded.
fn flush_opens(app: &AppHandle, window: &WebviewWindow) {
    let queue = {
        let mut s = shell(app);
        let Some(folder) = s.windows.get_mut(window.label()) else { return };
        folder.ready = true;
        std::mem::take(&mut folder.queue)
    };
    for file in queue {
        send_open(app, window, file);
    }
}

#[tauri::command]
async fn open_window(app: AppHandle, which: String, tab: Option<String>) {
    open_panel(&app, &which, tab);
}

/// Double-click on the title strip, as the system setting says: zoom (or fill), minimize, or nothing.
#[tauri::command]
fn toggle_zoom(window: WebviewWindow) {
    let action = std::process::Command::new("defaults")
        .args(["read", "-g", "AppleActionOnDoubleClick"])
        .output()
        .map(|o| String::from_utf8_lossy(&o.stdout).trim().to_string())
        .unwrap_or_default();
    match action.as_str() {
        "Minimize" => {
            let _ = window.minimize();
        }
        "None" => {}
        _ if window.is_maximized().unwrap_or(false) => {
            let _ = window.unmaximize();
        }
        _ => {
            let _ = window.maximize();
        }
    }
}

/// A window changed a shared localStorage key: keep it for the next launch and hand it to every other window.
#[tauri::command]
fn settings_set(app: AppHandle, window: WebviewWindow, key: String, json: String) {
    if !SHARED_KEYS.contains(&key.as_str()) || json.len() > 65536 {
        return;
    }
    let (theme, lang) = {
        let mut s = shell(&app);
        if s.store.get(&key).and_then(Value::as_str) == Some(json.as_str()) {
            return;
        }
        s.store.insert(key.clone(), Value::String(json.clone()));
        write_json(&data_file(&app, "web-storage.json"), &Value::Object(s.store.clone()));
        (theme_of(&s.store), lang_setting(&s.store))
    };
    if key == "writer-settings" {
        app.set_theme(theme);
        #[cfg(target_os = "macos")]
        app_languages(&lang);
    }
    #[cfg(not(target_os = "macos"))]
    let _ = lang;
    let script = format!("window.__writerStore && window.__writerStore({}, {})", json!(key), json!(json));
    for (label, other) in app.webview_windows() {
        if label != window.label() {
            let _ = other.eval(&script);
        }
    }
}

/// 设置 › 语言 › 立即重启: the menu bar and every page read the language once, at launch. The engines stop on the way
/// out (RunEvent::Exit), as on quit.
#[tauri::command]
fn restart_app(app: AppHandle) {
    app.request_restart();
}

#[tauri::command]
fn aux_close(window: WebviewWindow) {
    if is_panel(window.label()) {
        let _ = window.close();
    }
}

// ---------------------------------------------------------------- recent files

fn note_recent(app: &AppHandle, file: &Path) {
    // 最近使用 lists saved documents only: not drafts (the user has not put them anywhere yet), nor what a page reports
    // that is not a document at all
    if !is_doc(file) || file.starts_with(drafts_folder(app)) {
        return;
    }
    let changed = {
        let mut s = shell(app);
        if s.recent.first().map(PathBuf::as_path) == Some(file) {
            false
        } else {
            s.recent.retain(|p| p != file);
            s.recent.insert(0, file.to_path_buf());
            s.recent.truncate(RECENT_MAX);
            write_json(&data_file(app, "recent.json"), &json!(s.recent));
            save_bookmarks(app, &s.recent, |p| p == file);
            true
        }
    };
    if changed {
        refresh_recent(app);
    }
}

/// A bookmark for every recent file (see bookmarks.rs): a new one where `fresh` says so, the saved one for the rest.
fn save_bookmarks(app: &AppHandle, recent: &[PathBuf], fresh: impl Fn(&Path) -> bool) {
    let path = data_file(app, "bookmarks.json");
    let old = read_json(&path);
    let marks: Map<String, Value> = recent
        .iter()
        .filter_map(|p| {
            let key = p.to_string_lossy().into_owned();
            let saved = || old.get(&key).cloned();
            let mark = if fresh(p) { bookmarks::make(p).map(Value::from).or_else(saved) } else { saved() };
            Some((key, mark?))
        })
        .collect();
    write_json(&path, &Value::Object(marks));
}

fn refresh_recent(app: &AppHandle) {
    let Some(menus) = app.try_state::<Menus>() else { return };
    let recent = shell(app).recent.clone();
    let sub = &menus.recent;
    for item in sub.items().unwrap_or_default() {
        let _ = sub.remove(&item);
    }
    let name = |p: &PathBuf| p.file_name().map(|n| n.to_string_lossy().into_owned()).unwrap_or_default();
    for (i, path) in recent.iter().enumerate() {
        let mut label = name(path);
        if recent.iter().filter(|p| name(p) == label).count() > 1 {
            if let Some(folder) = path.parent().and_then(Path::file_name) {
                label = format!("{label} — {}", folder.to_string_lossy());
            }
        }
        if let Ok(item) = MenuItemBuilder::with_id(format!("recent:{i}"), label).build(app) {
            let _ = sub.append(&item);
        }
    }
    if !recent.is_empty() {
        if let Ok(sep) = PredefinedMenuItem::separator(app) {
            let _ = sub.append(&sep);
        }
    }
    if let Ok(clear) = MenuItemBuilder::with_id("recentClear", t("清除菜单")).enabled(!recent.is_empty()).build(app) {
        let _ = sub.append(&clear);
    }
}

// ---------------------------------------------------------------- menu bar (design: WriterMac menuDefs)

fn build_menu(app: &AppHandle) -> tauri::Result<(Menu<Wry>, Menus)> {
    let item = |id: &str, text: &str, key: &str| {
        let b = MenuItemBuilder::with_id(id, text);
        if key.is_empty() { b.build(app) } else { b.accelerator(key).build(app) }
    };
    let sep = || PredefinedMenuItem::separator(app);

    let writer = SubmenuBuilder::new(app, "Writer")
        .item(&item("about", t("关于 Writer"), "")?)
        .item(&sep()?)
        .item(&item("settings", t("设置…"), "Cmd+,")?)
        .item(&sep()?)
        .item(&PredefinedMenuItem::services(app, Some(t("服务")))?)
        .item(&sep()?)
        .item(&PredefinedMenuItem::hide(app, Some(t("隐藏 Writer")))?)
        .item(&PredefinedMenuItem::hide_others(app, Some(t("隐藏其他")))?)
        .item(&PredefinedMenuItem::show_all(app, Some(t("全部显示")))?)
        .item(&sep()?)
        .item(&PredefinedMenuItem::quit(app, Some(t("退出 Writer")))?)
        .build()?;
    if !cfg!(feature = "appstore") {
        writer.insert(&item("checkUpdate", t(update::MENU), "")?, 1)?;
    }

    let recent = SubmenuBuilder::new(app, t("打开最近使用")).build()?;
    let file = SubmenuBuilder::new(app, t("文件"))
        .item(&item("new", t("新建…"), "Cmd+N")?)
        .item(&item("newTab", t("新建标签页"), "Cmd+T")?)
        .item(&item("open", t("打开…"), "Cmd+O")?)
        .item(&recent)
        .item(&sep()?)
        .item(&item("closeTab", t("关闭标签页"), "Cmd+W")?)
        .item(&item("closeWindow", t("关闭窗口"), "Shift+Cmd+W")?)
        .item(&item("save", t("存储"), "Cmd+S")?)
        .item(&item("saveAs", t("另存为…"), "Shift+Cmd+S")?)
        .item(&item("exportAs", t("导出为…"), "")?)
        .item(&sep()?)
        .item(&item("print", t("打印…"), "Cmd+P")?)
        .build()?;

    let edit = SubmenuBuilder::new(app, t("编辑"))
        .item(&PredefinedMenuItem::undo(app, Some(t("撤销")))?)
        .item(&PredefinedMenuItem::redo(app, Some(t("重做")))?)
        .item(&sep()?)
        .item(&PredefinedMenuItem::cut(app, Some(t("剪切")))?)
        .item(&PredefinedMenuItem::copy(app, Some(t("拷贝")))?)
        .item(&PredefinedMenuItem::paste(app, Some(t("粘贴")))?)
        .item(&PredefinedMenuItem::select_all(app, Some(t("全选")))?)
        .item(&sep()?)
        .item(&item("find", t("查找…"), "Cmd+F")?)
        .item(&item("replace", t("替换…"), "Alt+Cmd+F")?)
        .build()?;

    let insert = SubmenuBuilder::new(app, t("插入"))
        .item(&item("insertTable", t("表格"), "")?)
        .item(&item("insertImage", t("图片…"), "Shift+Cmd+I")?)
        .item(&item("insertLink", t("链接"), "Cmd+K")?)
        .item(&item("insertComment", t("批注"), "Alt+Cmd+A")?)
        .item(&sep()?)
        .item(&item("pageBreak", t("分页符"), "Cmd+Enter")?)
        .build()?;

    let format = SubmenuBuilder::new(app, t("格式"))
        .item(&SubmenuBuilder::new(app, t("字体")).enabled(false).build()?)
        .item(&item("bold", t("粗体"), "Cmd+B")?)
        .item(&item("italic", t("斜体"), "Cmd+I")?)
        .item(&item("underline", t("下划线"), "Cmd+U")?)
        .item(&sep()?)
        .item(&SubmenuBuilder::new(app, t("段落样式")).enabled(false).build()?)
        .item(&item("clearFormat", t("清除格式"), "Alt+Cmd+\\")?)
        .build()?;

    let sidebar = CheckMenuItemBuilder::with_id("toggleSidebar", t("显示边栏")).accelerator("Ctrl+Cmd+S").checked(true).build(app)?;
    let toolbar = item("toggleToolbar", t("收起工具栏"), "Alt+Cmd+T")?;
    let dark = item("toggleDark", t("深色模式"), "Shift+Cmd+L")?;
    let view = SubmenuBuilder::new(app, t("显示"))
        .item(&sidebar)
        .item(&toolbar)
        .item(&item("toggleAI", t("✦ AI 助手"), "Cmd+J")?)
        .item(&item("immersive", t("沉浸书写"), "Cmd+.")?)
        .item(&sep()?)
        .item(&dark)
        .item(&sep()?)
        .item(&PredefinedMenuItem::fullscreen(app, Some(t("进入全屏幕")))?)
        .build()?;

    let window = SubmenuBuilder::new(app, t("窗口"))
        .item(&PredefinedMenuItem::minimize(app, Some(t("最小化")))?)
        .item(&PredefinedMenuItem::maximize(app, Some(t("缩放")))?)
        .item(&sep()?)
        .item(&item("prevTab", t("显示上一个标签页"), "Ctrl+Shift+Tab")?)
        .item(&item("nextTab", t("显示下一个标签页"), "Ctrl+Tab")?)
        .item(&sep()?)
        .item(&PredefinedMenuItem::bring_all_to_front(app, Some(t("将所有窗口移到前面")))?)
        .build()?;

    let help = SubmenuBuilder::new(app, t("帮助"))
        .item(&item("help", t("Writer 帮助"), "Shift+Cmd+/")?)
        .item(&item("whatsNew", t("新功能"), "")?)
        .item(&sep()?)
        .item(&item("gestures", t("触控板手势"), "")?)
        .item(&item("shortcuts", t("键盘快捷键"), "")?)
        .item(&sep()?)
        .item(&MenuItemBuilder::with_id("feedback", t("反馈问题…")).enabled(false).build(app)?)
        .build()?;

    let menu = Menu::with_items(app, &[&writer, &file, &edit, &insert, &format, &view, &window, &help])?;
    #[cfg(target_os = "macos")]
    {
        window.set_as_windows_menu_for_nsapp()?;
        help.set_as_help_menu_for_nsapp()?;
    }
    Ok((menu, Menus { sidebar, toolbar, dark, recent }))
}

fn apply_menu(app: &AppHandle, state: &Value) {
    let Some(menus) = app.try_state::<Menus>() else { return };
    let on = |k: &str| state.get(k).and_then(Value::as_bool).unwrap_or(false);
    let _ = menus.sidebar.set_checked(on("showThumbs"));
    let _ = menus.toolbar.set_text(t(if on("collapsed") { "展开工具栏" } else { "收起工具栏" }));
    let _ = menus.dark.set_text(t(if on("dark") { "浅色模式" } else { "深色模式" }));
}

fn on_menu(app: &AppHandle, id: &str) {
    let focused = app.webview_windows().into_values().find(|w| w.is_focused().unwrap_or(false));
    match id {
        "about" | "whatsNew" => open_panel(app, "about", None),
        "settings" => open_panel(app, "settings", None),
        "gestures" => open_panel(app, "gestures", None),
        "shortcuts" => open_panel(app, "settings", Some("keys".into())),
        "checkUpdate" => update::check(app, true),
        "open" => {
            let handle = app.clone();
            app.dialog().file().set_title(t("打开")).add_filter(t("Writer 文稿"), &DOC_EXTS).pick_files(move |files| {
                for file in files.unwrap_or_default() {
                    if let Ok(path) = file.into_path() {
                        open_file(&handle, path);
                    }
                }
            });
        }
        "recentClear" => {
            shell(app).recent.clear();
            write_json(&data_file(app, "recent.json"), &json!([]));
            write_json(&data_file(app, "bookmarks.json"), &json!({}));
            refresh_recent(app);
        }
        _ if id.starts_with("recent:") => {
            let path = id[7..].parse::<usize>().ok().and_then(|i| shell(app).recent.get(i).cloned());
            if let Some(path) = path {
                if path.is_file() {
                    open_file(app, path);
                } else {
                    shell(app).recent.retain(|p| *p != path);
                    refresh_recent(app);
                }
            }
        }
        "closeWindow" => {
            if let Some(w) = focused {
                let _ = w.close();
            }
        }
        "closeTab" if focused.as_ref().is_some_and(|w| is_panel(w.label())) => {
            let _ = focused.map(|w| w.close());
        }
        _ => {
            // Everything else acts on the page of the front document window (window.__writerMenu in ui/mac.dc.html).
            let label = focused.map(|w| w.label().to_string()).filter(|l| !is_panel(l)).or_else(|| shell(app).front.clone());
            match label.and_then(|l| app.get_webview_window(&l)) {
                Some(w) => {
                    let _ = w.eval(format!("window.__writerMenu && window.__writerMenu({})", json!(id)));
                }
                None if matches!(id, "new" | "newTab") => {
                    open_folder(app, drafts_folder(app), 3);
                }
                None => {}
            }
        }
    }
}

// ---------------------------------------------------------------- app

/// Engines of closed windows that are still running (stop_engine); a Windows update waits for them (update.rs).
static STOPPING: AtomicUsize = AtomicUsize::new(0);

fn stop_engine(folder: Folder) {
    if let Some(engine) = folder.engine {
        STOPPING.fetch_add(1, Ordering::SeqCst);
        // a moment for the page's last autosave to land before the engine goes away
        std::thread::spawn(move || {
            std::thread::sleep(std::time::Duration::from_millis(1500));
            let _ = engine.kill();
            STOPPING.fetch_sub(1, Ordering::SeqCst);
        });
    }
}

fn launch(app: &AppHandle) -> Result<(), Box<dyn std::error::Error>> {
    {
        let mut s = shell(app);
        if let Value::Object(store) = read_json(&data_file(app, "web-storage.json")) {
            s.store = store;
        }
        if let Value::Array(list) = read_json(&data_file(app, "recent.json")) {
            s.recent = list.iter().filter_map(Value::as_str).map(PathBuf::from).collect();
        }
        // the store sandbox forgets opened files when the app quits: the recent ones come back through their bookmarks,
        // which also follow a file the user moved
        let marks = read_json(&data_file(app, "bookmarks.json"));
        let mut moved = false;
        for p in s.recent.iter_mut() {
            if let Some(now) = marks.get(&*p.to_string_lossy()).and_then(Value::as_str).and_then(bookmarks::open) {
                moved |= now != *p;
                *p = now;
            }
        }
        if moved {
            write_json(&data_file(app, "recent.json"), &json!(s.recent));
            save_bookmarks(app, &s.recent, |_| true);
        }
    }
    let theme = theme_of(&shell(app).store);
    app.set_theme(theme);
    let lang = lang_setting(&shell(app).store);
    ENGLISH.store(english(&lang, &sys_locale::get_locale().unwrap_or_else(|| "zh".into())), Ordering::Relaxed);
    #[cfg(target_os = "macos")]
    app_languages(&lang);
    // Windows has no menu bar: an app menu would become a bar inside every window; win.dc.html has the W menu instead.
    if cfg!(target_os = "macos") {
        let (menu, menus) = build_menu(app)?;
        app.set_menu(menu)?;
        app.manage(menus);
        refresh_recent(app);
    }
    update::start(app)?;

    let mut files = std::mem::take(&mut shell(app).launch_files);
    files.extend(std::env::args_os().skip(1).map(PathBuf::from).filter(|p| is_doc(p) && p.is_file()));
    shell(app).launched = true;
    if files.is_empty() {
        open_folder(app, drafts_folder(app), 3);
    }
    for file in files {
        open_file(app, file);
    }
    #[cfg(feature = "test-hooks")]
    test_plan(app);
    Ok(())
}

/// UI checks without Accessibility access (`tauri build --features test-hooks`): WRITER_TEST_PLAN is a JSON file of steps
/// [{"at": seconds, "menu": id} | {"at": s, "eval": [window label, js]} | {"at": s, "exit": true}], or the steps themselves
/// (a sandboxed build cannot read a plan file outside its container).
#[cfg(feature = "test-hooks")]
fn test_plan(app: &AppHandle) {
    let Ok(path) = std::env::var("WRITER_TEST_PLAN") else { return };
    let plan = if path.trim_start().starts_with('[') { serde_json::from_str(&path).unwrap_or(Value::Null) } else { read_json(Path::new(&path)) };
    let Value::Array(steps) = plan else { return };
    for step in steps {
        let app = app.clone();
        std::thread::spawn(move || {
            std::thread::sleep(std::time::Duration::from_secs_f64(step["at"].as_f64().unwrap_or(1.0)));
            let handle = app.clone();
            let _ = app.run_on_main_thread(move || {
                if let Some(id) = step["menu"].as_str() {
                    on_menu(&handle, id);
                } else if let (Some(label), Some(js)) = (step["eval"][0].as_str(), step["eval"][1].as_str()) {
                    let label: String = if label == "front" { shell(&handle).front.clone().unwrap_or_default() } else { label.into() };
                    eprintln!("test eval in {label}");
                    if let Some(w) = handle.get_webview_window(&label) {
                        let _ = w.eval(js);
                    }
                } else if let Some(label) = step["url"].as_str() {
                    let url = handle.get_webview_window(label).and_then(|w| w.url().ok());
                    eprintln!("test url of {label}: {url:?}");
                } else if let Some(on) = step["top"].as_bool() {
                    // float every window so it paints and can be captured without taking focus from the user
                    for w in handle.webview_windows().values() {
                        let _ = w.set_always_on_top(on);
                    }
                } else if step["menus"].as_bool() == Some(true) {
                    use tauri::menu::MenuItemKind;
                    for top in handle.menu().and_then(|m| m.items().ok()).unwrap_or_default() {
                        let MenuItemKind::Submenu(sub) = top else { continue };
                        let items: Vec<String> = sub.items().unwrap_or_default().iter().map(|i| match i {
                            MenuItemKind::MenuItem(x) => x.text().unwrap_or_default(),
                            MenuItemKind::Check(x) => format!("{}{}", if x.is_checked().unwrap_or(false) { "✓" } else { "" }, x.text().unwrap_or_default()),
                            MenuItemKind::Predefined(x) => format!("[{}]", x.text().unwrap_or_default()),
                            MenuItemKind::Submenu(x) => format!("{} ›", x.text().unwrap_or_default()),
                            MenuItemKind::Icon(x) => x.text().unwrap_or_default(),
                        }).collect();
                        eprintln!("test menu {}: {}", sub.text().unwrap_or_default(), items.join(" | "));
                    }
                } else if step["exit"].as_bool() == Some(true) {
                    handle.exit(0);
                }
            });
        });
    }
}

fn main() {
    let app = tauri::Builder::default()
        .plugin(tauri_plugin_shell::init())
        .plugin(tauri_plugin_dialog::init())
        .plugin(
            tauri_plugin_window_state::Builder::default()
                .with_denylist(&["settings", "about", "gestures"])
                .with_state_flags(StateFlags::SIZE | StateFlags::POSITION)
                // Document windows (main-1, main-2, …) share one saved frame, tracked as they move and written on quit.
                // It is restored in open_folder, only for a window not cascaded next to another, so no restore here.
                .map_label(|label| if label.starts_with("main-") { "main" } else { label })
                .skip_initial_state("main")
                .build(),
        )
        .manage(Writer(Mutex::new(Shell::default())))
        .invoke_handler(tauri::generate_handler![
            shell_state, open_window, toggle_zoom, settings_set, aux_close, restart_app, win::open_files, win::snap_window,
            storage::save_dialog, storage::recent_files, storage::open_recent, update::check_update
        ])
        .on_menu_event(|app, event| on_menu(app, event.id().as_ref()))
        .on_window_event(|window, event| {
            let app = window.app_handle();
            let label = window.label().to_string();
            if matches!(event, WindowEvent::Resized(_) | WindowEvent::Focused(true)) {
                if let Some(w) = app.get_webview_window(&label) {
                    hide_lights(&w); // AppKit can bring them back after full screen
                }
            }
            match event {
                WindowEvent::Focused(true) if label.starts_with("main-") => {
                    let state = {
                        let mut s = shell(app);
                        s.front = Some(label.clone());
                        s.windows.get(&label).map(|f| f.state.clone())
                    };
                    if let Some(state) = state.filter(|v| !v.is_null()) {
                        apply_menu(app, &state);
                    }
                }
                WindowEvent::DragDrop(DragDropEvent::Drop { paths, .. }) => {
                    for path in paths.clone() {
                        open_file(app, path);
                    }
                }
                WindowEvent::Destroyed => {
                    let folder = {
                        let mut s = shell(app);
                        let folder = s.windows.remove(&label);
                        if s.front.as_deref() == Some(label.as_str()) {
                            s.front = s.windows.keys().next().cloned();
                        }
                        folder
                    };
                    if let Some(folder) = folder {
                        stop_engine(folder);
                    }
                }
                _ => {}
            }
        })
        .setup(|app| launch(app.handle()))
        .build(tauri::generate_context!())
        .expect("failed to build Writer desktop app");

    let mut restarting = false;
    app.run(move |app, event| match event {
        RunEvent::Opened { urls } => {
            let files: Vec<PathBuf> = urls.iter().filter_map(|u| u.to_file_path().ok()).collect();
            if shell(app).launched {
                files.into_iter().for_each(|f| open_file(app, f));
            } else {
                shell(app).launch_files.extend(files);
            }
        }
        RunEvent::Reopen { has_visible_windows: false, .. } => {
            let front = shell(app).front.clone().and_then(|l| app.get_webview_window(&l));
            match front {
                Some(w) => {
                    let _ = w.unminimize();
                    let _ = w.set_focus();
                }
                None => {
                    open_folder(app, drafts_folder(app), 3);
                }
            }
        }
        // Like other Mac document apps, Writer stays in the Dock when its last window closes; ⌘Q quits. Windows quits.
        RunEvent::ExitRequested { code: None, api, .. } if cfg!(target_os = "macos") => api.prevent_exit(),
        // 立即重启 (an update, a language change): a version installed on the way out opens again
        RunEvent::ExitRequested { code: Some(tauri::RESTART_EXIT_CODE), .. } => restarting = true,
        RunEvent::Exit => {
            let engines: Vec<Folder> = shell(app).windows.drain().map(|(_, f)| f).collect();
            for folder in engines {
                if let Some(engine) = folder.engine {
                    let _ = engine.kill();
                }
            }
            update::install_on_exit(restarting); // a downloaded version
        }
        _ => {}
    });
}

#[cfg(test)]
mod tests {
    use super::english;

    #[test]
    fn the_language_follows_the_setting_then_the_system() {
        assert!(english("English", "zh-Hans-CN"));
        assert!(!english("简体中文", "en-US"));
        assert!(!english("跟随系统", "zh-Hans-CN"));
        assert!(!english("", "ZH-TW"));
        assert!(english("跟随系统", "en-GB"));
        assert!(english("", "fr-FR"));
    }
}
