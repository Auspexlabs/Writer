fn main() {
    // App commands get generated allow-* permissions; capabilities/writer-mac-shell.json (and writer-win-shell.json on
    // Windows) grants them to the engine's pages.
    let commands = &[
        "shell_state", "open_window", "tear_off", "toggle_zoom", "settings_set", "aux_close", "open_files", "snap_window",
        "save_dialog", "recent_files", "open_recent", "restart_app", "check_update", "default_app", "open_with",
    ];
    tauri_build::try_build(
        tauri_build::Attributes::new().app_manifest(tauri_build::AppManifest::new().commands(commands)),
    )
    .expect("failed to run tauri-build");
}
