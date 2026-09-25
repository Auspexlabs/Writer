// In-app updates, in the builds that update themselves (the App Store updates the store build): a check 10 s after
// launch and once a day while Writer runs, unless 设置 › 通用 › 自动更新 is off, and 检查更新… (the Writer menu on
// macOS, the W menu on Windows). An automatic check shows nothing; 检查更新… ends in one dialog: 已是最新版本, or
// 新版本已下载 with 立即重启 / 好.
//
// New versions (tauri-plugin-updater): latest.json on the website, then on the newest GitHub release
// (plugins.updater.endpoints in tauri.macos.conf.json / tauri.windows.conf.json). A newer version downloads in the
// background, is checked against plugins.updater.pubkey and installs when Writer quits (install_on_exit, after the
// engines stopped), so the next launch runs it. macOS replaces the app bundle only then: new windows start the engine
// and read the interface from the bundle, so replacing it mid-session would mix an old shell with a new engine and
// interface. Windows runs the NSIS installer quietly (plugins.updater.windows.installMode).
//
// Interface packages, when there is no new version: ui.json (plugins.updater.uiEndpoints, tried like the endpoints
// above) may offer a newer interface for this exact version, {"base", "build", "url", "signature", …}: a tar.gz of ui/,
// signed with `tauri signer sign --app-version <base>-ui.<build>`. It is unpacked into <app data>/ui/<base>-<build>/ and
// recorded in <app data>/ui/state.json; from the next launch every window and panel is served from it (ui_package). If
// a window on it does not report in, it is marked bad and Writer restarts on the bundled ui/ (main.rs spawn_engine). A
// feed that answers 404 (no interface release yet) means there is none.
// `scripts/release-mac.sh web` and `scripts/release-win.sh` make the feeds and the signed files, `scripts/release-ui.sh`
// the interface packages.

use std::fs;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, Ordering::SeqCst};
use std::sync::{Mutex, OnceLock};
use std::time::Duration;

use base64::Engine;
use serde_json::{Map, Value};
use tauri::AppHandle;
use tauri_plugin_dialog::{DialogExt, MessageDialogButtons};
use tauri_plugin_updater::UpdaterExt;

// Everything the updater shows, in one place; the English is in t() in main.rs. {v} is a version.
pub const MENU: &str = "检查更新…";
const DOWNLOADED: &str = "新版本已下载";
const DOWNLOADED_BODY: &str = "重新打开 Writer 后生效。";
const RESTART: &str = "立即重启";
const OK: &str = "好";
const LATEST: &str = "已是最新版本";
const LATEST_BODY: &str = "Writer {v} 是目前的最新版本。";
const CHECK_FAILED: &str = "无法检查更新";
const FAILED: &str = "更新失败";
const FAILED_BODY: &str = "请稍后再试，或前往 https://github.com/Auspexlabs/writer/releases 下载最新版本。";
const OPEN: &str = "前往下载页";
const RELEASES: &str = "https://github.com/Auspexlabs/writer/releases";

type Error = Box<dyn std::error::Error>;

/// A check, a download or its dialog is under way.
static BUSY: AtomicBool = AtomicBool::new(false);
/// A new version (PENDING) or interface package waits for the next launch.
static READY: AtomicBool = AtomicBool::new(false);
/// Installs the downloaded version; true: open Writer again afterwards.
type Install = Box<dyn FnOnce(bool) + Send>;
/// The downloaded version's install, run when Writer quits. ponytail: the download is held in memory until then (tens of
/// MB); a temporary file if that ever matters.
static PENDING: Mutex<Option<Install>> = Mutex::new(None);
/// The interface package this run serves (ui_package).
static PACKAGE: OnceLock<Option<(PathBuf, u64)>> = OnceLock::new();
/// state.json is written by the update thread and when a window does not report in.
static STATE: Mutex<()> = Mutex::new(());

pub fn start(app: &AppHandle) -> tauri::Result<()> {
    ui_package(app); // this run's interface, chosen before a check can unpack a newer one
    if cfg!(feature = "appstore") {
        return Ok(());
    }
    app.plugin(tauri_plugin_updater::Builder::new().build())?;
    let app = app.clone();
    std::thread::spawn(move || {
        let mut wait = Duration::from_secs(10);
        loop {
            std::thread::sleep(wait);
            wait = Duration::from_secs(24 * 3600);
            if automatic(&app) {
                check(&app, false);
            }
        }
    });
    Ok(())
}

/// 设置 › 通用 › 自动更新, on unless turned off.
fn automatic(app: &AppHandle) -> bool {
    let settings = crate::shell(app).store.get("writer-settings").and_then(Value::as_str).and_then(|s| serde_json::from_str::<Value>(s).ok());
    settings.and_then(|s| s.get("autoUpdate")?.as_bool()).unwrap_or(true)
}

/// 检查更新… in the W menu (ui/win.dc.html, capabilities/writer-win-shell.json).
#[tauri::command]
pub fn check_update(app: AppHandle) {
    check(&app, true);
}

/// Looks for a new version, then for a newer interface package, and downloads what it finds. Only a `manual` check (the
/// menu item) says how it went: the website or GitHub may not be reachable, and an automatic check stays silent.
pub fn check(app: &AppHandle, manual: bool) {
    if BUSY.swap(true, SeqCst) {
        return;
    }
    let app = app.clone();
    std::thread::spawn(move || {
        run(&app, manual);
        BUSY.store(false, SeqCst);
    });
}

fn run(app: &AppHandle, manual: bool) {
    if !READY.load(SeqCst) {
        if let Err(title) = fetch(app, manual) {
            if manual {
                failed(app, crate::t(title));
            }
            return;
        }
    }
    if !manual {
        return;
    }
    if READY.load(SeqCst) {
        let buttons = MessageDialogButtons::OkCancelCustom(crate::t(RESTART).into(), crate::t(OK).into());
        if ask(app, crate::t(DOWNLOADED), crate::t(DOWNLOADED_BODY), buttons) {
            app.request_restart(); // RunEvent::Exit installs a downloaded version first
        }
    } else {
        let mut version = app.package_info().version.to_string();
        if let Some((_, build)) = ui_package(app) {
            version = format!("{version} ({build})");
        }
        let body = crate::t(LATEST_BODY).replace("{v}", &version);
        ask(app, crate::t(LATEST), &body, MessageDialogButtons::OkCustom(crate::t(OK).into()));
    }
}

/// Downloads a new version for PENDING or, when there is none, a newer interface package; READY once either is in
/// place. Err: the title of what failed.
fn fetch(app: &AppHandle, manual: bool) -> Result<(), &'static str> {
    let found = tauri::async_runtime::block_on(async { app.updater()?.check().await });
    match found {
        Ok(Some(update)) => {
            if !manual && !replaceable() {
                return Ok(());
            }
            // checks the signature against plugins.updater.pubkey; nothing is installed yet
            let bytes = tauri::async_runtime::block_on(update.download(|_, _| {}, || {})).map_err(|_| FAILED)?;
            *PENDING.lock().unwrap_or_else(|e| e.into_inner()) = Some(Box::new(move |relaunch| {
                let _ = update.restart_after_install(relaunch).install(bytes);
            }));
            READY.store(true, SeqCst);
        }
        Ok(None) => match stage_ui(app) {
            Ok(staged) => READY.store(staged, SeqCst),
            // a missing or broken package is not worth a dialog: this version is the latest
            Err(_error) => {
                #[cfg(feature = "test-hooks")]
                eprintln!("interface package: {_error}");
            }
        },
        Err(_) => return Err(CHECK_FAILED),
    }
    Ok(())
}

/// macOS: whether this user may replace the app bundle, so that installing at quit asks for no administrator password.
/// Where it cannot (a Writer someone else put in /Applications, one run from a disk image), only 检查更新… installs.
#[cfg(target_os = "macos")]
fn replaceable() -> bool {
    use std::os::unix::ffi::OsStrExt;
    let Some(bundle) = std::env::current_exe().ok().and_then(|exe| tauri_plugin_updater::extract_path_from_executable(&exe).ok()) else {
        return false;
    };
    // the bundle and the folder it is in: the updater moves it out and the new one in
    [bundle.as_path(), bundle.parent().unwrap_or(&bundle)].iter().all(|dir| {
        // SAFETY: access() only reads the NUL-terminated path it is given
        std::ffi::CString::new(dir.as_os_str().as_bytes()).is_ok_and(|path| unsafe { libc::access(path.as_ptr(), libc::W_OK) } == 0)
    })
}

#[cfg(not(target_os = "macos"))]
fn replaceable() -> bool {
    true
}

/// On the way out (RunEvent::Exit, after the engines stopped): a downloaded version installs now. macOS replaces the app
/// bundle; Windows starts the installer, which runs quietly, and this process exits. `relaunch` (a restart: 立即重启, or
/// after a language change) opens Writer again: Tauri's restart on macOS, the installer's /R on Windows. If the install
/// fails, Writer restarts or quits as it was going to.
pub fn install_on_exit(relaunch: bool) {
    let Some(install) = PENDING.lock().unwrap_or_else(|e| e.into_inner()).take() else { return };
    // the engines of windows closed just before stop 1.5 s after them (stop_engine); Windows' installer replaces their binary
    for _ in 0..40 {
        if crate::STOPPING.load(SeqCst) == 0 {
            break;
        }
        std::thread::sleep(Duration::from_millis(50));
    }
    install(relaunch);
}

fn failed(app: &AppHandle, title: &str) {
    let buttons = MessageDialogButtons::OkCancelCustom(crate::t(OPEN).into(), crate::t(OK).into());
    if ask(app, title, crate::t(FAILED_BODY), buttons) && !cfg!(feature = "test-hooks") {
        // Windows' explorer opens a URL in the default browser, as `open` does on macOS
        let _ = std::process::Command::new(if cfg!(windows) { "explorer" } else { "open" }).arg(RELEASES).status();
    }
}

/// A native dialog, from a thread other than the main one; true when its first button was chosen.
fn ask(app: &AppHandle, title: &str, body: &str, buttons: MessageDialogButtons) -> bool {
    // update tests: WRITER_TEST_UPDATE answers every dialog with its first button
    #[cfg(feature = "test-hooks")]
    if std::env::var_os("WRITER_TEST_UPDATE").is_some() {
        eprintln!("test dialog: {title} | {body}");
        return true;
    }
    app.dialog().message(body).title(title).buttons(buttons).blocking_show()
}

// ---------------------------------------------------------------- interface packages

/// <app data>/ui: a folder per package, <base>-<build>, and state.json: {"<base>-<build>": "ready" | "bad"}.
fn packages(app: &AppHandle) -> PathBuf {
    crate::data_file(app, "ui")
}

fn read_state(root: &Path) -> Map<String, Value> {
    match crate::read_json(&root.join("state.json")) {
        Value::Object(state) => state,
        _ => Map::new(),
    }
}

/// The build of a package named `<version>-<build>`, if it is one for `version`.
fn build_of(name: &str, version: &str) -> Option<u64> {
    name.strip_prefix(version)?.strip_prefix('-')?.parse().ok()
}

/// The package a launch serves: the newest ready one for exactly this version.
fn pick(state: &Map<String, Value>, version: &str) -> Option<u64> {
    state.iter().filter(|(_, status)| *status == "ready").filter_map(|(name, _)| build_of(name, version)).max()
}

/// The newest build of this version fetched so far, bad ones included, so that those are not fetched again.
fn newest(state: &Map<String, Value>, version: &str) -> u64 {
    state.keys().filter_map(|name| build_of(name, version)).max().unwrap_or(0)
}

/// The interface package this run serves, chosen once, at launch: its folder and build, or None for the bundled ui/.
/// Everything else in <app data>/ui goes: other versions' packages (a new version brought its own ui/), older builds,
/// unfinished downloads; this version's bad builds stay in state.json.
pub fn ui_package(app: &AppHandle) -> Option<&'static (PathBuf, u64)> {
    PACKAGE
        .get_or_init(|| {
            let root = packages(app);
            if cfg!(feature = "appstore") || !root.is_dir() {
                return None;
            }
            let version = app.package_info().version.to_string();
            let _lock = STATE.lock().unwrap_or_else(|e| e.into_inner());
            let mut state = read_state(&root);
            state.retain(|name, status| *status == "bad" || root.join(name).join("mac.dc.html").is_file());
            let build = pick(&state, &version);
            let keep = build.map(|b| format!("{version}-{b}"));
            state.retain(|name, status| Some(name) == keep.as_ref() || (*status == "bad" && build_of(name, &version).is_some()));
            crate::write_json(&root.join("state.json"), &Value::Object(state));
            for entry in fs::read_dir(&root).into_iter().flatten().flatten() {
                if entry.path().is_dir() && entry.file_name().to_str() != keep.as_deref() {
                    let _ = fs::remove_dir_all(entry.path());
                }
            }
            Some((root.join(keep?), build?))
        })
        .as_ref()
}

/// A window on this run's interface package did not report in: the package is never used again.
pub fn reject_package(app: &AppHandle) {
    let Some(name) = ui_package(app).and_then(|(dir, _)| dir.file_name()?.to_str()) else { return };
    let root = packages(app);
    let _lock = STATE.lock().unwrap_or_else(|e| e.into_inner());
    let mut state = read_state(&root);
    state.insert(name.into(), "bad".into());
    crate::write_json(&root.join("state.json"), &Value::Object(state));
}

/// A newer interface package for exactly this version, downloaded, checked and unpacked for the next launch; true when
/// there was one.
fn stage_ui(app: &AppHandle) -> Result<bool, Error> {
    let config = app.config().plugins.0.get("updater").cloned().unwrap_or_default();
    let insecure = config["dangerousInsecureTransportProtocol"].as_bool() == Some(true);
    let endpoints = config["uiEndpoints"].as_array().cloned().unwrap_or_default();
    let version = app.package_info().version.to_string();
    let root = packages(app);
    tauri::async_runtime::block_on(async {
        let mut found = None;
        for url in endpoints.iter().filter_map(Value::as_str).filter(|url| insecure || url.starts_with("https://")) {
            // as the updater does: the next endpoint only after a connection error or a non-2xx status
            if let Ok(response) = get(url).await {
                found = Some(response.bytes().await?);
                break;
            }
        }
        let manifest: Value = serde_json::from_slice(&found.ok_or("no ui.json")?)?;
        let build = manifest["build"].as_u64().unwrap_or(0);
        if manifest["base"].as_str() != Some(&version) || build <= newest(&read_state(&root), &version) {
            return Ok(false);
        }
        let archive = get(manifest["url"].as_str().ok_or("ui.json has no url")?).await?.bytes().await?;
        let signature = manifest["signature"].as_str().ok_or("ui.json has no signature")?;
        verify(&archive, signature, config["pubkey"].as_str().unwrap_or_default(), &format!("{version}-ui.{build}"))?;
        // unpacks the very bytes just verified
        let name = format!("{version}-{build}");
        unpack(&archive, &root, &name)?;
        let _lock = STATE.lock().unwrap_or_else(|e| e.into_inner());
        let mut state = read_state(&root);
        state.insert(name, "ready".into());
        crate::write_json(&root.join("state.json"), &Value::Object(state));
        Ok::<_, Error>(true)
    })
}

async fn get(url: &str) -> reqwest::Result<reqwest::Response> {
    // rustls' crypto provider, which this client needs, is in place: the updater's own check has just run
    let client = reqwest::Client::builder().user_agent(concat!("Writer/", env!("CARGO_PKG_VERSION"))).build()?;
    client.get(url).send().await?.error_for_status()
}

/// The archive's minisign signature against plugins.updater.pubkey, and the version its trusted comment names (the
/// signature covers the comment), as tauri-plugin-updater checks a new version under requireSignedVersion.
fn verify(archive: &[u8], signature: &str, pubkey: &str, version: &str) -> Result<(), Error> {
    let text = |b64: &str| -> Result<String, Error> { Ok(String::from_utf8(base64::engine::general_purpose::STANDARD.decode(b64)?)?) };
    let signature = minisign_verify::Signature::decode(&text(signature)?)?;
    minisign_verify::PublicKey::decode(&text(pubkey)?)?.verify(archive, &signature, true)?;
    let signed = signature.trusted_comment().split('\t').find_map(|field| field.strip_prefix("version:"));
    if signed != Some(version) {
        return Err(format!("the package is signed for {signed:?}, not {version}").into());
    }
    Ok(())
}

/// Unpacks the archive's ui/ into <root>/<name>, through a folder that is renamed into place once complete.
fn unpack(archive: &[u8], root: &Path, name: &str) -> Result<(), Error> {
    let part = root.join(format!(".part-{name}"));
    let _ = fs::remove_dir_all(&part);
    fs::create_dir_all(&part)?;
    let unpacked = (|| -> Result<(), Error> {
        tar::Archive::new(flate2::read::GzDecoder::new(archive)).unpack(&part)?;
        if !part.join("ui/mac.dc.html").is_file() {
            return Err("the package has no ui/mac.dc.html".into());
        }
        let _ = fs::remove_dir_all(root.join(name));
        Ok(fs::rename(part.join("ui"), root.join(name))?)
    })();
    let _ = fs::remove_dir_all(&part);
    unpacked
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    #[test]
    fn a_launch_serves_the_newest_good_package_of_its_own_version() {
        let state = |v: Value| v.as_object().unwrap().clone();
        let s = state(json!({ "0.1.2-9": "ready", "0.1.2-10": "ready", "0.1.1-12": "ready", "0.1.20-11": "ready" }));
        assert_eq!(pick(&s, "0.1.2"), Some(10), "builds count as numbers, other versions do not count");
        assert_eq!(pick(&s, "0.1.3"), None, "no package for this version: the bundled ui/");
        let s = state(json!({ "0.1.2-9": "ready", "0.1.2-10": "bad" }));
        assert_eq!(pick(&s, "0.1.2"), Some(9), "a bad build is never used");
        assert_eq!(newest(&s, "0.1.2"), 10, "nor fetched again");
        assert_eq!(pick(&state(json!({ "0.1.2-x": "ready", "0.1.2": "ready" })), "0.1.2"), None);
    }

    // made with a throwaway key: `tauri signer sign --app-version <v>` of b"writer ui package"
    const KEY: &str = "dW50cnVzdGVkIGNvbW1lbnQ6IG1pbmlzaWduIHB1YmxpYyBrZXk6IDIwMjg2ODNFNjE1RDlCOUIKUldTYm0xMWhQbWdvSVAyeVgrTDNXeGlpK2VyZzhFRTkzSTFmSm1KZ1dQL29Vb3BhSGFmQlBZaFcK";
    const SIGNED_UI_3: &str = "dW50cnVzdGVkIGNvbW1lbnQ6IHNpZ25hdHVyZSBmcm9tIHRhdXJpIHNlY3JldCBrZXkKUlVTYm0xMWhQbWdvSUIwUnFGemYxVlZJazFlN3F6ODJDV21lYVR3ZzN2Ukw3ODc2Rk1reEJZRDRyRldkZUoxNWg0R2lmaThRR01vV1VVMzU5T1doemVVcjg0RmVOSDRBeHdRPQp0cnVzdGVkIGNvbW1lbnQ6IHRpbWVzdGFtcDoxNzkwMzAzNjY0CWZpbGU6ZGF0YS5iaW4JdmVyc2lvbjowLjEuMi11aS4zCllLaDM2djRGbHRFOEtkL0VIbk9nRVpyT295SDlFTmFJejY5VFNIcVBRam0ydDlQUjVUODJxT2dScjN6RGJCOUtUdWlneGZXMS93YytPV1dFdGY3SkNBPT0K";
    const SIGNED_UI_4: &str = "dW50cnVzdGVkIGNvbW1lbnQ6IHNpZ25hdHVyZSBmcm9tIHRhdXJpIHNlY3JldCBrZXkKUlVTYm0xMWhQbWdvSVAxZXpGOUJ0TCtEMnl1bnA4ZVVzSEd0OEJZdy9md1pORXdaZ0Y0U2MxazkyL2hsenZtNFdURUYvR1Z2YUpTQlFONFhjRlNWUVNhdmQ2ZGRFakc2eXd3PQp0cnVzdGVkIGNvbW1lbnQ6IHRpbWVzdGFtcDoxNzkwMzAzNjczCWZpbGU6ZGF0YS5iaW4JdmVyc2lvbjowLjEuMi11aS40CmpDdEg1Mkg2eW1wNXJoM3ljVSsxV2NQcG95dUJCb2tqTk10aTBCeDF0NDhHRG9nZEtLSkNaNHpjZysxdVV4OGZuM0Z3cWVhYjVGSm5BMHdrRTNrN0JRPT0K";
    const SIGNED_NO_VERSION: &str = "dW50cnVzdGVkIGNvbW1lbnQ6IHNpZ25hdHVyZSBmcm9tIHRhdXJpIHNlY3JldCBrZXkKUlVTYm0xMWhQbWdvSU03YVNwaXA4cGw5K2V1WGREREdhSkhuWGhtcEZUQWk2NWJjcU5NSjl2OGk4cHBjMURPSFZYOG45MHJqV1BWamVtR3YycUxqS0JtYkllSnI1K3BxRmdjPQp0cnVzdGVkIGNvbW1lbnQ6IHRpbWVzdGFtcDoxNzkwMzAzNjczCWZpbGU6ZGF0YS5iaW4KRGt3NllIN3N0K3IzOUpkbjlacTc0MjZUVUpUS3ViVXRpK2FTeUZ5OXVKNkc3RTg2NktpYUZ5eGhwM08zL0twam1aR1lxUWZkSzhPMnZmKzBpbWFiQ2c9PQo=";

    #[test]
    fn a_package_must_be_signed_by_the_key_for_its_own_build() {
        let data = b"writer ui package";
        assert!(verify(data, SIGNED_UI_3, KEY, "0.1.2-ui.3").is_ok());
        assert!(verify(b"writer ui packagE", SIGNED_UI_3, KEY, "0.1.2-ui.3").is_err(), "other bytes");
        assert!(verify(data, SIGNED_UI_4, KEY, "0.1.2-ui.3").is_err(), "signed for another build");
        assert!(verify(data, SIGNED_NO_VERSION, KEY, "0.1.2-ui.3").is_err(), "signed for no version");
        let config: Value = serde_json::from_str(include_str!("../tauri.conf.json")).unwrap();
        let writer_key = config["plugins"]["updater"]["pubkey"].as_str().unwrap();
        assert!(verify(data, SIGNED_UI_3, writer_key, "0.1.2-ui.3").is_err(), "another key");
    }

    #[test]
    fn a_downloaded_version_installs_once_when_writer_quits() {
        let (sent, installed) = std::sync::mpsc::channel();
        *PENDING.lock().unwrap() = Some(Box::new(move |relaunch| sent.send(relaunch).unwrap()));
        install_on_exit(true);
        assert_eq!(installed.try_recv(), Ok(true), "installed, and Writer opens again after a restart");
        install_on_exit(false);
        assert!(installed.try_recv().is_err(), "nothing left to install");
    }
}
