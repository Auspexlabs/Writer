// 默认打开方式: does a double-click on a .docx, .xlsx, .pptx, .md or .mm file open Writer? The first document window of
// a version's first launch asks (ui/mac.dc.html, ui/win.dc.html: 把 Writer 设为默认打开方式？ with 设为默认 / 以后再说;
// the answer is kept per version in <app data>/default-app.json), and 设置 › 通用 › 默认打开方式 shows the state. macOS
// asks LaunchServices per extension (NSWorkspace, the way Finder resolves a double-click) and sets it the same way.
// Windows keeps the user's choice behind a hash no app may write, so 设为默认 opens 设置 › 默认应用 at Writer
// (ms-settings:defaultapps?registeredAppUser=Writer), where the installer registers Writer (windows/hooks.nsh).

use std::sync::atomic::{AtomicBool, Ordering::SeqCst};

use serde_json::{json, Value};
use tauri::{AppHandle, WebviewWindow};

/// The types the sheet and the setting are about (.markdown follows .md; Writer only views PDFs).
pub const EXTS: [&str; 5] = ["docx", "xlsx", "pptx", "md", "mm"];

/// One sheet per run: the first document window to ask gets it.
static ASKED: AtomicBool = AtomicBool::new(false);

/// `action`: "state" → {"isDefault": true | false | null, "types": [".docx", …]}; "ask" → {"ask": bool, "types"}, true for
/// the first document window of a run while Writer is not the default and this version has not asked; "set" → makes
/// Writer the default (macOS) or opens 设置 › 默认应用 (Windows), then the state; "later" → 以后再说 for this version.
#[tauri::command]
pub async fn default_app(app: AppHandle, window: WebviewWindow, action: String) -> Result<Value, String> {
    let (version, file) = (app.package_info().version.to_string(), crate::data_file(&app, "default-app.json"));
    let (id, types) = (app.config().identifier.clone(), EXTS.map(|e| format!(".{e}")));
    match action.as_str() {
        "ask" => {
            let ask = window.label().starts_with("main-")
                && should_ask(crate::read_json(&file)["asked"].as_str(), &version, is_default(&id))
                && !ASKED.swap(true, SeqCst);
            Ok(json!({ "ask": ask, "types": types }))
        }
        "set" | "later" => {
            crate::write_json(&file, &json!({ "asked": version }));
            if action == "set" {
                make_default()?;
            }
            Ok(json!({ "isDefault": is_default(&id), "types": types }))
        }
        _ => Ok(json!({ "isDefault": is_default(&id), "types": types })),
    }
}

/// The sheet shows while Writer is not the default, once per version: 设为默认 and 以后再说 both count as `asked`.
fn should_ask(asked: Option<&str>, version: &str, is_default: Option<bool>) -> bool {
    is_default == Some(false) && asked != Some(version)
}

/// Whether every type opens with Writer; None where that cannot be told.
#[cfg(target_os = "macos")]
fn is_default(identifier: &str) -> Option<bool> {
    Some(EXTS.iter().all(|ext| handler(ext).as_deref() == Some(identifier)))
}

/// The bundle identifier of the app a double-click on a .`ext` file opens.
#[cfg(target_os = "macos")]
fn handler(ext: &str) -> Option<String> {
    use objc2_app_kit::NSWorkspace;
    use objc2_foundation::{NSBundle, NSString};
    use objc2_uniform_type_identifiers::UTType;
    let kind = UTType::typeWithFilenameExtension(&NSString::from_str(ext))?;
    let url = NSWorkspace::sharedWorkspace().URLForApplicationToOpenContentType(&kind)?;
    Some(NSBundle::bundleWithURL(&url)?.bundleIdentifier()?.to_string())
}

/// This app bundle becomes the default for every type, one LaunchServices call each; each answer is waited for, so that
/// 设为默认 can say which type failed.
#[cfg(target_os = "macos")]
fn make_default() -> Result<(), String> {
    use block2::RcBlock;
    use objc2_app_kit::NSWorkspace;
    use objc2_foundation::{NSBundle, NSError, NSString};
    use objc2_uniform_type_identifiers::UTType;
    let bundle = NSBundle::mainBundle().bundleURL();
    for ext in EXTS {
        let kind = UTType::typeWithFilenameExtension(&NSString::from_str(ext)).ok_or(format!(".{ext}: unknown type"))?;
        let (sent, done) = std::sync::mpsc::channel();
        let finished = RcBlock::new(move |error: *mut NSError| {
            // SAFETY: a non-null error is a live NSError while the handler runs
            let _ = sent.send(unsafe { error.as_ref() }.map(|e| e.localizedDescription().to_string()));
        });
        NSWorkspace::sharedWorkspace().setDefaultApplicationAtURL_toOpenContentType_completionHandler(&bundle, &kind, Some(&*finished));
        if let Some(error) = done.recv_timeout(std::time::Duration::from_secs(10)).map_err(|_| format!(".{ext}: no answer from macOS"))? {
            return Err(format!(".{ext}: {error}"));
        }
    }
    Ok(())
}

#[cfg(windows)]
fn is_default(_identifier: &str) -> Option<bool> {
    let exe = std::env::current_exe().ok()?.to_string_lossy().into_owned();
    Some(EXTS.iter().all(|ext| handler(ext).is_some_and(|program| program.eq_ignore_ascii_case(&exe))))
}

/// The program a double-click on a .`ext` file runs, as the shell resolves it (the user's choice first).
#[cfg(windows)]
fn handler(ext: &str) -> Option<String> {
    use windows_sys::Win32::UI::Shell::{AssocQueryStringW, ASSOCF_NONE, ASSOCSTR_EXECUTABLE};
    let assoc: Vec<u16> = format!(".{ext}\0").encode_utf16().collect();
    let mut out = [0u16; 1024];
    let mut len = out.len() as u32;
    // SAFETY: assoc is NUL-terminated, and out/len describe the buffer the shell writes the NUL-terminated path into
    let ok = unsafe { AssocQueryStringW(ASSOCF_NONE, ASSOCSTR_EXECUTABLE, assoc.as_ptr(), std::ptr::null(), out.as_mut_ptr(), &mut len) } == 0;
    (ok && len > 1).then(|| String::from_utf16_lossy(&out[..len as usize - 1]))
}

/// Windows lets no app make itself the default: 设置 › 默认应用, opened at Writer (explorer hands the URI to the shell,
/// as update.rs does with the download page).
#[cfg(windows)]
fn make_default() -> Result<(), String> {
    std::process::Command::new("explorer").arg("ms-settings:defaultapps?registeredAppUser=Writer").spawn().map(drop).map_err(|e| e.to_string())
}

#[cfg(not(any(target_os = "macos", windows)))]
fn is_default(_identifier: &str) -> Option<bool> {
    None
}

#[cfg(not(any(target_os = "macos", windows)))]
fn make_default() -> Result<(), String> {
    Err("not on this system".into())
}

#[cfg(test)]
mod tests {
    use super::should_ask;

    #[test]
    fn the_sheet_shows_once_per_version_while_writer_is_not_the_default() {
        assert!(should_ask(None, "0.1.2", Some(false)), "the first launch");
        assert!(!should_ask(Some("0.1.2"), "0.1.2", Some(false)), "以后再说 in this version");
        assert!(should_ask(Some("0.1.1"), "0.1.2", Some(false)), "a new version asks once more");
        assert!(!should_ask(None, "0.1.2", Some(true)), "already the default");
        assert!(!should_ask(None, "0.1.2", None), "unknown: no nagging");
    }
}
