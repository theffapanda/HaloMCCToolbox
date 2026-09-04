use std::collections::HashSet;
use std::path::PathBuf;
use std::sync::Arc;
use std::time::Duration;

use anyhow::Result;
use sysinfo::{ProcessRefreshKind, ProcessesToUpdate, System};

use crate::policy::{self, PolicyState};
use crate::flows::{self, FlowTable};
use crate::process::Resolver;
use crate::status_server::StatusBus;
use crate::tunneled::{self, TunneledPaths};

// We still poll, but the *primary* admission path is now in divert.rs (it
// admits PIDs the instant their first packet shows up). This watcher exists
// mainly to (a) catch already-running processes when the user adds a path
// to the tunnel list, and (b) garbage-collect dead PIDs.
const POLL_INTERVAL: Duration = Duration::from_millis(500);
const CONFIG_RELOAD_INTERVAL: Duration = Duration::from_secs(2);

/// Admit matching processes before WinDivert packet capture starts. This
/// prevents an already-running game from sending on its normal interface
/// during the watcher's first polling interval.
pub fn seed(policy_state: &PolicyState, bus: &StatusBus, tunneled_paths: &TunneledPaths) {
    let sys = System::new_all();
    let paths: Vec<PathBuf> = tunneled_paths.read().unwrap().clone();
    for (pid, process) in sys.processes() {
        let Some(exe) = process.exe() else { continue };
        if !tunneled::matches_any(exe, &paths) {
            continue;
        }
        let pid_u32 = pid.as_u32();
        policy::add(policy_state, pid_u32);
        bus.pid_paths
            .insert(pid_u32, exe.to_string_lossy().to_string());
        tracing::info!("auto-tunnel (startup): pid={} matches {:?}", pid_u32, exe);
    }
}

pub async fn run(
    policy_state: PolicyState,
    bus: Arc<StatusBus>,
    resolver: Arc<Resolver>,
    flows: FlowTable,
    tunneled_paths: TunneledPaths,
) -> Result<()> {
    tracing::info!("proc watcher running (polls every {:?})", POLL_INTERVAL);

    let mut sys = System::new_all();
    let mut last_config_load = std::time::Instant::now() - CONFIG_RELOAD_INTERVAL;

    loop {
        if last_config_load.elapsed() >= CONFIG_RELOAD_INTERVAL {
            let from_disk = tunneled::load_from_disk();
            tunneled::replace(&tunneled_paths, from_disk);
            last_config_load = std::time::Instant::now();
        }

        sys.refresh_processes_specifics(
            ProcessesToUpdate::All,
            true,
            ProcessRefreshKind::new().with_exe(sysinfo::UpdateKind::Always),
        );

        // Snapshot the path list so we don't hold the read lock during sysinfo
        // iteration (which can be a few ms on a busy box).
        let paths: Vec<PathBuf> = tunneled_paths.read().unwrap().clone();
        let mut matching_pids = HashSet::new();
        for (pid, process) in sys.processes() {
            let pid_u32 = pid.as_u32();
            let matched_path = process
                .exe()
                .filter(|exe| tunneled::matches_any(exe, &paths))
                .map(PathBuf::from)
                .or_else(|| {
                    // sysinfo can transiently omit executable paths on Windows,
                    // especially around an elevated child launch. Querying the
                    // process directly keeps a valid packet-time admission from
                    // being erased on the next 500 ms policy refresh.
                    resolver
                        .resolve(pid_u32)
                        .filter(|info| {
                            tunneled::matches_any(
                                std::path::Path::new(&info.exe_path),
                                &paths,
                            )
                        })
                        .map(|info| PathBuf::from(info.exe_path))
                });
            let Some(exe) = matched_path else { continue };
            let path_str = exe.to_string_lossy().to_string();
            if !policy::contains(&policy_state, pid_u32) {
                tracing::info!("auto-tunnel (poll): pid={} matches {:?}", pid_u32, exe);
            }
            bus.pid_paths.insert(pid_u32, path_str);
            matching_pids.insert(pid_u32);
        }

        let still_alive: HashSet<u32> = sys
            .processes()
            .keys()
            .map(|p| p.as_u32())
            .collect();
        // Policy is derived from the current process snapshot, rather than
        // accumulating PIDs forever. This prevents a later process from
        // inheriting MCC routing when Windows reuses a PID.
        policy::replace(&policy_state, matching_pids);
        resolver.retain_alive(&still_alive);
        flows::retain_pids(&flows, &still_alive);
        bus.pid_paths.retain(|p, _| still_alive.contains(p));
        bus.pid_bytes_out.retain(|p, _| still_alive.contains(p));
        bus.pid_bytes_in.retain(|p, _| still_alive.contains(p));

        tokio::time::sleep(POLL_INTERVAL).await;
    }
}
