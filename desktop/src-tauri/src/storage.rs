// No default folder: new documents are drafts in <app data>/Drafts until the first 存储, where the system Save dialog
// asks where they go. The chosen file is granted to that window's engine (writer app --grants-stdin); recent files open
// in their folder's window. Shared by macOS and Windows.

use std::path::{Path, PathBuf};

use tauri::{AppHandle, Manager, WebviewWindow};
use tauri_plugin_dialog::DialogExt;

/// A line of an engine's grants channel ("allow" or "list", see grant() in main.rs), or None for a path that is not
/// absolute or could forge a second line.
pub fn grant_line(verb: &str, path: &Path) -> Option<String> {
    let text = path.to_str()?;
    (path.is_absolute() && !text.contains(['\n', '\r'])).then(|| format!("{verb} {text}\n"))
}

/// Paths as the page sees them: '/' separators on every platform.
pub fn page_path(path: &Path) -> String {
    path.to_string_lossy().replace('\\', "/")
}

/// The recent file the page asked for, if it is one: a page may reopen recent files, not arbitrary folders.
pub fn in_recent(recent: &[PathBuf], wanted: &str) -> Option<PathBuf> {
    let wanted = wanted.replace('\\', "/");
    let same = |p: &PathBuf| if cfg!(target_os = "linux") { page_path(p) == wanted } else { page_path(p).eq_ignore_ascii_case(&wanted) };
    recent.iter().find(|p| same(p)).cloned()
}

/// The Save dialog's suggested file name: the title without path characters, plus an extension the engine writes.
pub fn save_name(title: &str, ext: &str) -> Option<String> {
    crate::SAVE_EXTS.contains(&ext).then_some(())?;
    let clean: String = title.chars().map(|c| if r#"\/:*?"<>|"#.contains(c) || c.is_control() { ' ' } else { c }).collect();
    let clean = clean.trim();
    Some(format!("{}.{ext}", if clean.is_empty() { "未命名" } else { clean }))
}

/// 存储 / 另存为: the system Save dialog (the suggested name, in dir when the page names one — a compatibility-mode
/// draft is offered beside the .doc/.xls/.ppt it came from — else Documents). The chosen file is granted to the calling
/// window's engine before its path goes back to the page.
#[tauri::command]
pub async fn save_dialog(app: AppHandle, window: WebviewWindow, name: String, ext: String, dir: Option<String>) -> Option<String> {
    let file_name = save_name(&name, &ext)?;
    let mut dialog = app.dialog().file().set_title(crate::t("存储")).set_file_name(&file_name).add_filter(&ext, &[ext.as_str()]).set_parent(&window);
    if let Some(dir) = dir.map(PathBuf::from).filter(|d| d.is_absolute() && d.is_dir()) {
        dialog = dialog.set_directory(dir);
    } else if let Ok(dir) = app.path().document_dir() {
        dialog = dialog.set_directory(dir);
    }
    let mut path = dialog.blocking_save_file()?.into_path().ok()?;
    if path.extension().is_none() {
        path.set_extension(&ext);
    }
    crate::grant(&app, window.label(), "allow", &path).then(|| page_path(&path))
}

/// 最近使用 on the start page: the recent files that still exist.
#[tauri::command]
pub fn recent_files(app: AppHandle) -> Vec<String> {
    crate::shell(&app).recent.iter().filter(|p| p.is_file()).map(|p| page_path(p)).collect()
}

/// Opens a recent file in its folder's window; anything not in the recent list is refused.
#[tauri::command]
pub fn open_recent(app: AppHandle, path: String) -> bool {
    let hit = in_recent(&crate::shell(&app).recent, &path);
    match hit {
        Some(p) if p.is_file() => { crate::open_file(&app, p); true }
        _ => false,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn grant_lines_are_one_absolute_path_per_line() {
        assert_eq!(grant_line("allow", Path::new("/Users/me/报告.docx")).as_deref(), Some("allow /Users/me/报告.docx\n"));
        assert_eq!(grant_line("list", Path::new("/Users/me/a.md")).as_deref(), Some("list /Users/me/a.md\n"));
        assert_eq!(grant_line("allow", Path::new("/Users/me/a\nallow /etc/x.md")), None); // a newline could forge a second grant
        assert_eq!(grant_line("allow", Path::new("relative.docx")), None);
    }

    #[test]
    fn recent_files_are_matched_by_their_page_form() {
        let recent = vec![PathBuf::from("/Users/me/Docs/a.docx"), PathBuf::from("/Users/me/b.md")];
        assert_eq!(in_recent(&recent, "/Users/me/b.md"), Some(PathBuf::from("/Users/me/b.md")));
        assert_eq!(in_recent(&recent, "/Users/me/c.md"), None);
        assert_eq!(page_path(Path::new(r"C:\Users\me\a.docx")), "C:/Users/me/a.docx");
    }

    #[test]
    fn save_names_are_plain_file_names_of_document_types() {
        assert_eq!(save_name("季度/报告", "docx").as_deref(), Some("季度 报告.docx"));
        assert_eq!(save_name("", "md").as_deref(), Some("未命名.md"));
        assert_eq!(save_name("x", "exe"), None);
        assert_eq!(save_name("x", "doc"), None); // read-only compatibility formats are opened, never saved as
        assert!(crate::DOC_EXTS.contains(&"doc") && crate::DOC_EXTS.contains(&"wps") && crate::DOC_EXTS.contains(&"csv"));
    }
}
