use std::sync::Arc;

use anyhow::{Context, Result};
use windivert::layer::SocketLayer;
use windivert::prelude::*;

use crate::flows::{self, FlowEntry, FlowTable, LocalEndpoint};
use crate::policy::{self, PolicyState};
use crate::process::Resolver;

pub fn run(resolver: Arc<Resolver>, policy_state: PolicyState, flows: FlowTable) -> Result<()> {
    let filter = "event = CONNECT or event = ACCEPT or event = BIND or event = CLOSE";
    let flags = WinDivertFlags::new().set_sniff().set_recv_only();
    let handle = WinDivert::socket(filter, 0, flags).context(
        "WinDivert::socket failed — start an elevated PowerShell, or check that WinDivert.dll + WinDivert64.sys are next to redirector.exe",
    )?;

    tracing::info!("SOCKET observer started, filter: {}", filter);

    loop {
        match handle.recv() {
            Ok(packet) => handle_event(&packet.address, &resolver, &policy_state, &flows),
            Err(WinDivertError::Recv(WinDivertRecvError::NoData)) => break,
            Err(e) => tracing::warn!("recv error: {}", e),
        }
    }
    Ok(())
}

fn handle_event(
    addr: &WinDivertAddress<SocketLayer>,
    resolver: &Resolver,
    policy_state: &PolicyState,
    flows: &FlowTable,
) {
    let pid = addr.process_id();
    let event = addr.event();
    let proto = addr.protocol();
    let local_addr = addr.local_address();
    let local_port = addr.local_port();

    match event {
        WinDivertEvent::SocketBind | WinDivertEvent::SocketConnect => {
            let key = LocalEndpoint { addr: local_addr, port: local_port, proto };
            flows::insert(flows, key, FlowEntry { pid });
        }
        WinDivertEvent::SocketClose => {
            let key = LocalEndpoint { addr: local_addr, port: local_port, proto };
            flows::remove(flows, &key);
        }
        _ => {}
    }

    log_event(addr, resolver, policy_state, pid, event, proto);
}

fn log_event(
    addr: &WinDivertAddress<SocketLayer>,
    resolver: &Resolver,
    policy_state: &PolicyState,
    pid: u32,
    event: WinDivertEvent,
    proto: u8,
) {
    let proto_s = proto_str(proto);
    let local = format!("{}:{}", addr.local_address(), addr.local_port());
    let remote = format!("{}:{}", addr.remote_address(), addr.remote_port());

    let event_str = match event {
        WinDivertEvent::SocketBind => "BIND",
        WinDivertEvent::SocketConnect => "CONNECT",
        WinDivertEvent::SocketAccept => "ACCEPT",
        WinDivertEvent::SocketListen => "LISTEN",
        WinDivertEvent::SocketClose => "CLOSE",
        _ => "OTHER",
    };

    let exe = resolver
        .resolve(pid)
        .map(|p| p.exe_name)
        .unwrap_or_else(|| "?".into());

    if policy::contains(policy_state, pid) {
        tracing::info!(
            "VPN socket: event={} pid={} proto={} local={} remote={} process={}",
            event_str, pid, proto_s, local, remote, exe
        );
    } else {
        tracing::debug!(
            "socket: event={} pid={} proto={} local={} remote={} process={}",
            event_str, pid, proto_s, local, remote, exe
        );
    }
}

fn proto_str(p: u8) -> &'static str {
    match p {
        6 => "TCP",
        17 => "UDP",
        1 => "ICMP",
        58 => "ICMP6",
        _ => "?",
    }
}
