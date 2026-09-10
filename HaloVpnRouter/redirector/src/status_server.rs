use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::Arc;

use dashmap::DashMap;

pub struct StatusBus {
    pub bytes_out: Arc<AtomicU64>,
    pub bytes_in: Arc<AtomicU64>,
    pub vpn_up_since_ms: Arc<AtomicU64>,
    pub pid_bytes_out: Arc<DashMap<u32, AtomicU64>>,
    pub pid_bytes_in: Arc<DashMap<u32, AtomicU64>>,
    pub pid_paths: Arc<DashMap<u32, String>>,
}

impl StatusBus {
    pub fn new() -> Self {
        Self {
            bytes_out: Arc::new(AtomicU64::new(0)),
            bytes_in: Arc::new(AtomicU64::new(0)),
            vpn_up_since_ms: Arc::new(AtomicU64::new(0)),
            pid_bytes_out: Arc::new(DashMap::new()),
            pid_bytes_in: Arc::new(DashMap::new()),
            pid_paths: Arc::new(DashMap::new()),
        }
    }

    pub fn add_pid_out(&self, pid: u32, n: u64) {
        self.pid_bytes_out
            .entry(pid)
            .or_insert_with(|| AtomicU64::new(0))
            .fetch_add(n, Ordering::Relaxed);
    }

    pub fn add_pid_in(&self, pid: u32, n: u64) {
        self.pid_bytes_in
            .entry(pid)
            .or_insert_with(|| AtomicU64::new(0))
            .fetch_add(n, Ordering::Relaxed);
    }
}
