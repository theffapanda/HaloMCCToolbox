use std::collections::HashMap;
use std::net::{IpAddr, Ipv4Addr, Ipv6Addr};
use std::sync::{Arc, RwLock};

pub const PROTO_TCP: u8 = 6;
pub const PROTO_UDP: u8 = 17;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct LocalEndpoint {
    pub addr: IpAddr,
    pub port: u16,
    pub proto: u8,
}

#[derive(Debug, Clone, Copy)]
pub struct FlowEntry {
    pub pid: u32,
}

pub type FlowTable = Arc<RwLock<HashMap<LocalEndpoint, FlowEntry>>>;

pub fn new() -> FlowTable {
    Arc::new(RwLock::new(HashMap::new()))
}

pub fn insert(table: &FlowTable, key: LocalEndpoint, entry: FlowEntry) {
    table.write().unwrap().insert(key, entry);
}

pub fn lookup(table: &FlowTable, key: &LocalEndpoint) -> Option<FlowEntry> {
    table.read().unwrap().get(key).copied()
}

pub fn remove(table: &FlowTable, key: &LocalEndpoint) {
    table.write().unwrap().remove(key);
}

pub fn retain_pids(table: &FlowTable, alive: &std::collections::HashSet<u32>) {
    table.write().unwrap().retain(|_, entry| alive.contains(&entry.pid));
}

/// Lookup that falls back to a wildcard (0.0.0.0 / ::) local address. UDP
/// sockets that use sendto without an explicit bind end up registered against
/// the wildcard address even though their outbound packets carry the real LAN
/// source IP. Tries the exact (addr, port, proto) first, then the wildcard
/// variant.
pub fn lookup_with_wildcard(
    table: &FlowTable,
    addr: IpAddr,
    port: u16,
    proto: u8,
) -> Option<FlowEntry> {
    let exact = LocalEndpoint { addr, port, proto };
    if let Some(e) = lookup(table, &exact) {
        return Some(e);
    }
    let wildcard_addr = match addr {
        IpAddr::V4(_) => IpAddr::V4(Ipv4Addr::UNSPECIFIED),
        IpAddr::V6(_) => IpAddr::V6(Ipv6Addr::UNSPECIFIED),
    };
    lookup(table, &LocalEndpoint { addr: wildcard_addr, port, proto })
}
