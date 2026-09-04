use std::path::{Path, PathBuf};
use std::sync::{Arc, RwLock};

/// Shared, mutable list of exe paths the user has marked as tunneled.
/// Proc-watcher reloads this from disk on its tick; the divert hot path
/// reads it on every PID-miss to decide whether to admit synchronously.
pub type TunneledPaths = Arc<RwLock<Vec<PathBuf>>>;

pub fn new() -> TunneledPaths {
    Arc::new(RwLock::new(Vec::new()))
}

pub fn replace(paths: &TunneledPaths, new_paths: Vec<PathBuf>) {
    *paths.write().unwrap() = new_paths;
}

pub fn matches(paths: &TunneledPaths, candidate: &Path) -> bool {
    let guard = paths.read().unwrap();
    matches_any(candidate, &guard)
}

pub fn matches_any(candidate: &Path, configured: &[PathBuf]) -> bool {
    let candidate = normalize_windows_path(candidate);
    configured
        .iter()
        .any(|path| normalize_windows_path(path) == candidate)
}

fn normalize_windows_path(path: &Path) -> String {
    let mut value = path.to_string_lossy().replace('/', "\\");
    if let Some(without_prefix) = value.strip_prefix(r"\\?\UNC\") {
        value = format!(r"\\{}", without_prefix);
    } else if let Some(without_prefix) = value.strip_prefix(r"\\?\") {
        value = without_prefix.to_owned();
    }
    value.trim_end_matches('\\').to_lowercase()
}

pub fn load_from_disk() -> Vec<PathBuf> {
    let Ok(appdata) = std::env::var("LOCALAPPDATA") else {
        return Vec::new();
    };
    let config_path = PathBuf::from(&appdata)
        .join("HaloMCCToolbox")
        .join("Vpn")
        .join("router.json");

    let Ok(bytes) = std::fs::read(&config_path) else {
        return Vec::new();
    };
    let Ok(json) = serde_json::from_slice::<serde_json::Value>(&bytes) else {
        return Vec::new();
    };
    let Some(apps) = json.get("tunneledApps").and_then(|v| v.as_array()) else {
        return Vec::new();
    };

    apps.iter()
        .filter_map(|app| {
            app.get("exePath")
                .and_then(|v| v.as_str())
                .filter(|s| !s.is_empty())
                .map(PathBuf::from)
        })
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn matches_windows_device_prefix_and_separator_variants() {
        let configured = vec![PathBuf::from(
            r"C:\Games\Halo The Master Chief Collection\MCC\Binaries\Win64\MCC-Win64-Shipping.exe",
        )];
        assert!(matches_any(
            Path::new(
                r"\\?\c:/games/halo the master chief collection/MCC/Binaries/Win64/MCC-Win64-Shipping.exe",
            ),
            &configured,
        ));
    }
}
