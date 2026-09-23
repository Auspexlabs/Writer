// Writer for Windows: the window is frameless (tauri.windows.conf.json) and ui/win.dc.html draws the design's title bar.
// The page drives its window with the window plugin (minimize, toggle_maximize, close, start_dragging); the two commands
// here are what it cannot do on its own. Compiled on every platform, granted only by capabilities/writer-win-shell.json.

use std::path::PathBuf;

use tauri::{AppHandle, PhysicalPosition, PhysicalSize, WebviewWindow};

/// canonicalize() on Windows returns verbatim paths (\\?\C:\…, \\?\UNC\server\…); the engine and the page get the
/// plain form. Other paths pass through unchanged.
pub fn plain_path(path: PathBuf) -> PathBuf {
    match path.to_str() {
        Some(s) if s.starts_with(r"\\?\UNC\") => PathBuf::from(format!(r"\\{}", &s[8..])),
        Some(s) if s.starts_with(r"\\?\") => PathBuf::from(&s[4..]),
        _ => path,
    }
}

/// 打开… (Ctrl+O and the W menu): the system open dialog; the chosen files open in place, like 文件 › 打开… on macOS.
#[tauri::command]
pub fn open_files(app: AppHandle) {
    crate::on_menu(&app, "open");
}

/// A zone of the design's snap layouts: x, y, width and height as fractions of the work area of the window's monitor.
#[tauri::command]
pub fn snap_window(window: WebviewWindow, x: f64, y: f64, w: f64, h: f64) {
    let Ok(Some(monitor)) = window.current_monitor() else { return };
    let area = monitor.work_area();
    let (px, py, pw, ph) = zone((area.position.x, area.position.y, area.size.width, area.size.height), [x, y, w, h]);
    if window.is_maximized().unwrap_or(false) {
        let _ = window.unmaximize();
    }
    let _ = window.set_position(PhysicalPosition::new(px, py));
    let _ = window.set_size(PhysicalSize::new(pw, ph));
}

/// The zone's pixels. Edges are rounded rather than widths, so neighbouring zones (thirds of 1920) meet without a gap.
fn zone((ax, ay, aw, ah): (i32, i32, u32, u32), f: [f64; 4]) -> (i32, i32, u32, u32) {
    let [fx, fy, fw, fh] = f.map(|v| if v.is_finite() { v.clamp(0.0, 1.0) } else { 0.0 });
    let edge = |origin: i32, len: u32, frac: f64| origin + (frac.min(1.0) * f64::from(len)).round() as i32;
    let (x0, x1) = (edge(ax, aw, fx), edge(ax, aw, fx + fw));
    let (y0, y1) = (edge(ay, ah, fy), edge(ay, ah, fy + fh));
    (x0, y0, (x1 - x0).max(1) as u32, (y1 - y0).max(1) as u32)
}

#[cfg(test)]
mod tests {
    use super::{plain_path, zone};
    use std::path::PathBuf;

    #[test]
    fn verbatim_paths_become_plain() {
        let p = |s: &str| plain_path(PathBuf::from(s)).to_string_lossy().into_owned();
        assert_eq!(p(r"\\?\C:\Users\me\Documents\Writer"), r"C:\Users\me\Documents\Writer");
        assert_eq!(p(r"\\?\UNC\server\share\a.docx"), r"\\server\share\a.docx");
        assert_eq!(p("/Users/me/Documents/Writer"), "/Users/me/Documents/Writer");
    }

    #[test]
    fn zones_tile_the_work_area() {
        let area = (0, 0, 1920, 1040); // 1080p above a 40px taskbar
        let thirds: Vec<_> = [0.0, 1.0 / 3.0, 2.0 / 3.0].iter().map(|&x| zone(area, [x, 0.0, 1.0 / 3.0, 1.0])).collect();
        assert_eq!(thirds, [(0, 0, 640, 1040), (640, 0, 640, 1040), (1280, 0, 640, 1040)]);
        assert_eq!(zone(area, [0.5, 0.5, 0.5, 0.5]), (960, 520, 960, 520));
        // a second monitor to the left, and fractions from the page that overshoot
        assert_eq!(zone((-2560, 0, 2560, 1400), [2.0 / 3.0, 0.0, 1.0 / 3.0, 1.0]), (-853, 0, 853, 1400));
        assert_eq!(zone(area, [0.5, 0.0, 0.8, f64::NAN]), (960, 0, 960, 1));
    }
}
