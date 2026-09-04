use std::collections::HashMap;
use std::collections::HashSet;
use std::sync::Mutex;

use windows::Win32::Foundation::CloseHandle;
use windows::Win32::System::Threading::{
    OpenProcess, QueryFullProcessImageNameW, PROCESS_NAME_WIN32, PROCESS_QUERY_LIMITED_INFORMATION,
};
use windows::core::PWSTR;

#[derive(Debug, Clone)]
pub struct ProcessInfo {
    pub exe_path: String,
    pub exe_name: String,
}

pub struct Resolver {
    cache: Mutex<HashMap<u32, Option<ProcessInfo>>>,
}

impl Resolver {
    pub fn new() -> Self {
        Self {
            cache: Mutex::new(HashMap::new()),
        }
    }

    pub fn resolve(&self, pid: u32) -> Option<ProcessInfo> {
        if pid == 0 || pid == 4 {
            return None;
        }
        {
            let cache = self.cache.lock().unwrap();
            if let Some(entry) = cache.get(&pid) {
                return entry.clone();
            }
        }
        let info = lookup(pid);
        // A process can be observable by the network stack slightly before
        // QueryFullProcessImageName succeeds. Do not make that transient miss
        // sticky; retry on the next packet or watcher pass.
        if info.is_some() {
            self.cache.lock().unwrap().insert(pid, info.clone());
        }
        info
    }

    pub fn retain_alive(&self, alive: &HashSet<u32>) {
        self.cache.lock().unwrap().retain(|pid, _| alive.contains(pid));
    }
}

fn lookup(pid: u32) -> Option<ProcessInfo> {
    unsafe {
        let handle = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, false, pid).ok()?;
        let mut buf = [0u16; 1024];
        let mut len = buf.len() as u32;
        let res = QueryFullProcessImageNameW(
            handle,
            PROCESS_NAME_WIN32,
            PWSTR(buf.as_mut_ptr()),
            &mut len,
        );
        let _ = CloseHandle(handle);
        res.ok()?;
        let path = String::from_utf16_lossy(&buf[..len as usize]);
        let exe_name = path.rsplit('\\').next().unwrap_or(&path).to_string();
        Some(ProcessInfo {
            exe_path: path,
            exe_name,
        })
    }
}
