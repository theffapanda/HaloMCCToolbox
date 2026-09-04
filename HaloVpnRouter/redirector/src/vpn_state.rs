use std::sync::atomic::Ordering;
use std::sync::{Arc, RwLock};
use std::time::Duration;

use anyhow::Result;

use crate::adapter::{self, VpnTarget};
use crate::status_server::StatusBus;

pub type VpnState = Arc<RwLock<Option<VpnTarget>>>;

const POLL_INTERVAL: Duration = Duration::from_secs(2);

pub fn new() -> VpnState {
    Arc::new(RwLock::new(None))
}

pub fn current(state: &VpnState) -> Option<VpnTarget> {
    *state.read().unwrap()
}

pub async fn watcher(state: VpnState, bus: Arc<StatusBus>) -> Result<()> {
    let mut last: Option<VpnTarget> = None;
    loop {
        let adapters = tokio::task::spawn_blocking(adapter::enumerate).await??;
        let now = adapter::find_vpn_target(&adapters);
        let same = match (last, now) {
            (Some(a), Some(b)) => a.ipv4 == b.ipv4 && a.if_index == b.if_index,
            (None, None) => true,
            _ => false,
        };
        if !same {
            match (last, now) {
                (None, Some(t)) => {
                    tracing::info!("VPN adapter UP at {} (ifindex {})", t.ipv4, t.if_index);
                    bus.vpn_up_since_ms.store(unix_millis(), Ordering::Relaxed);
                }
                (Some(old), None) => {
                    tracing::warn!("VPN adapter DOWN (was {})", old.ipv4);
                    bus.vpn_up_since_ms.store(0, Ordering::Relaxed);
                }
                (Some(old), Some(new)) => {
                    tracing::info!(
                        "VPN adapter changed: {} (ifindex {}) -> {} (ifindex {})",
                        old.ipv4, old.if_index, new.ipv4, new.if_index
                    );
                    bus.vpn_up_since_ms.store(unix_millis(), Ordering::Relaxed);
                }
                _ => {}
            }
            *state.write().unwrap() = now;
            last = now;
        }
        tokio::time::sleep(POLL_INTERVAL).await;
    }
}

fn unix_millis() -> u64 {
    std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|d| d.as_millis() as u64)
        .unwrap_or(0)
}
