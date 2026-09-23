// 最近使用 across launches in the Mac App Store sandbox, which forgets a file the user opened once the app quits: every
// recent file keeps a security-scoped bookmark (<app data>/bookmarks.json), resolved when the app launches. The engines
// run inside the app's sandbox (com.apple.security.inherit) and share its access, so access started here reaches them
// too, even an engine that was already running. Outside the sandbox a bookmark also follows a file the user moved.

use std::path::{Path, PathBuf};

#[cfg(target_os = "macos")]
use objc2_foundation::{NSData, NSString, NSURLBookmarkCreationOptions, NSURLBookmarkResolutionOptions, NSURL};

/// A bookmark (hex) for a file this process may access right now: just opened, or just chosen in the Save dialog.
#[cfg(target_os = "macos")]
pub fn make(path: &Path) -> Option<String> {
    let url = NSURL::fileURLWithPath(&NSString::from_str(path.to_str()?));
    let data = url
        .bookmarkDataWithOptions_includingResourceValuesForKeys_relativeToURL_error(NSURLBookmarkCreationOptions::WithSecurityScope, None, None)
        .ok()?;
    Some(data.to_vec().iter().map(|b| format!("{b:02x}")).collect())
}

/// Resolves a bookmark and keeps its file accessible for the rest of the run; returns where the file is now.
#[cfg(target_os = "macos")]
pub fn open(hex: &str) -> Option<PathBuf> {
    let bytes = (0..hex.len()).step_by(2).map(|i| u8::from_str_radix(hex.get(i..i + 2)?, 16).ok()).collect::<Option<Vec<u8>>>()?;
    let options = NSURLBookmarkResolutionOptions::WithSecurityScope | NSURLBookmarkResolutionOptions::WithoutUI;
    // SAFETY: a null is_stale pointer is allowed; a stale bookmark is made again the next time its file is opened
    let url = unsafe {
        NSURL::URLByResolvingBookmarkData_options_relativeToURL_bookmarkDataIsStale_error(&NSData::with_bytes(&bytes), options, None, std::ptr::null_mut())
    }
    .ok()?;
    // ponytail: never stopped, at most RECENT_MAX files for the app's life; stop on close if that ever grows
    // SAFETY: plain message send on a file URL
    unsafe { url.startAccessingSecurityScopedResource() };
    Some(PathBuf::from(url.path()?.to_string()))
}

#[cfg(not(target_os = "macos"))]
pub fn make(_: &Path) -> Option<String> {
    None
}

#[cfg(not(target_os = "macos"))]
pub fn open(_: &str) -> Option<PathBuf> {
    None
}

#[cfg(all(test, target_os = "macos"))]
mod tests {
    use super::*;

    #[test]
    fn a_bookmark_finds_its_file_after_a_move() {
        let dir = std::env::temp_dir().join(format!("writer-bookmark-{}", std::process::id()));
        std::fs::create_dir_all(&dir).unwrap();
        let (a, b) = (dir.join("a.md"), dir.join("b.md"));
        std::fs::write(&a, "x").unwrap();
        let mark = make(&a.canonicalize().unwrap()).expect("bookmark");
        std::fs::rename(&a, &b).unwrap();
        assert_eq!(open(&mark), Some(b.canonicalize().unwrap()));
        assert_eq!(open("zz"), None);
        std::fs::remove_dir_all(&dir).unwrap();
    }
}
