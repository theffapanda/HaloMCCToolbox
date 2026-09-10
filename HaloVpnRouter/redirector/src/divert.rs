use std::borrow::Cow;
use std::collections::HashMap;
use std::net::{IpAddr, Ipv4Addr, SocketAddr, SocketAddrV4};
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::Arc;
use std::time::{Duration, Instant};

use anyhow::{Context, Result};
use etherparse::{NetSlice, SlicedPacket, TransportSlice};
use smoltcp::iface::{Config, Interface, SocketHandle, SocketSet};
use smoltcp::socket::tcp;
use smoltcp::wire::{HardwareAddress, IpAddress, IpCidr, IpEndpoint, Ipv4Address};
use tokio::io::{AsyncReadExt, AsyncWriteExt};
use tokio::sync::{mpsc, oneshot};
use windivert::layer::NetworkLayer;
use windivert::prelude::*;
use windivert_sys::ChecksumFlags;

use crate::adapter::VpnTarget;
use crate::flows::{self, FlowEntry, FlowTable, LocalEndpoint, PROTO_TCP, PROTO_UDP};
use crate::pidlookup;
use crate::policy::{self, PolicyState};
use crate::process::Resolver;
use crate::stack::VirtualDevice;
use crate::status_server::StatusBus;
use crate::tunneled::{self, TunneledPaths};
use crate::vpn_state::{self, VpnState};

const BRIDGE_CHANNEL_DEPTH: usize = 64;
const SMOLTCP_CHUNK: usize = 16 * 1024;
const UDP_IDLE_TIMEOUT: Duration = Duration::from_secs(60);
const CAPTURE_BATCH_SIZE: u8 = 64;
const CAPTURE_BUFFER_SIZE: usize = 65_535 * CAPTURE_BATCH_SIZE as usize;

#[allow(clippy::too_many_arguments)]
pub fn run(
    policy_state: PolicyState,
    flows: FlowTable,
    vpn_state: VpnState,
    bus: Arc<StatusBus>,
    resolver: Arc<Resolver>,
    tunneled_paths: TunneledPaths,
) -> Result<()> {
    // Capture IPv6 too. The bridge is IPv4-only, so selected processes have
    // IPv6 consumed below instead of leaking around the VPN.
    let capture_filter = "outbound and (ip or ipv6) and (tcp or udp) and !loopback";
    let capture = WinDivert::network(capture_filter, 1040, WinDivertFlags::new())
        .context("WinDivert::network capture handle failed (need admin)")?;
    let inject = WinDivert::network(
        "false",
        1039,
        WinDivertFlags::new().set_send_only(),
    )
    .context("WinDivert::network inject handle failed")?;

    tracing::info!(
        "divert: capture+inject open, filter={}, smoltcp mode",
        capture_filter
    );

    let tokio_handle = tokio::runtime::Handle::current();

    let mut device = VirtualDevice::new();
    let mut iface = Interface::new(
        Config::new(HardwareAddress::Ip),
        &mut device,
        smoltcp::time::Instant::now(),
    );
    iface.update_ip_addrs(|addrs| {
        let _ = addrs.push(IpCidr::new(IpAddress::Ipv4(Ipv4Address::new(0, 0, 0, 0)), 0));
    });
    iface.set_any_ip(true);

    let mut sockets: SocketSet<'static> = SocketSet::new(Vec::new());
    let mut tcp_flows: HashMap<TcpKey, TcpFlow> = HashMap::new();
    let mut udp_flows: HashMap<UdpKey, UdpFlow> = HashMap::new();
    let mut local_interfaces: HashMap<Ipv4Addr, InjectionInterface> = HashMap::new();

    let stats = Arc::new(Stats::new(bus.clone()));
    let stats_clone = stats.clone();
    std::thread::spawn(move || stats_clone.report_loop());

    // Capturing one packet at a time serializes the whole machine's network
    // traffic through this userspace loop. Receive and reinject unrelated
    // traffic in batches so selective routing cannot throttle other apps.
    let mut buffer = vec![0u8; CAPTURE_BUFFER_SIZE];
    let mut last_vpn = vpn_state::current(&vpn_state);

    loop {
        let now_vpn = vpn_state::current(&vpn_state);
        if now_vpn != last_vpn {
            on_vpn_change(last_vpn, now_vpn, &mut tcp_flows, &mut udp_flows, &mut sockets, &stats);
            last_vpn = now_vpn;
        }

        match capture.recv_wait_ex(&mut buffer, CAPTURE_BATCH_SIZE, 5) {
            Ok(packets) => {
                let mut passthrough = Vec::with_capacity(packets.len());
                for packet in packets {
                    if let Some(packet) = handle_capture(
                        packet,
                        &policy_state,
                        &flows,
                        &mut tcp_flows,
                        &mut udp_flows,
                        &mut sockets,
                        &mut device,
                        &tokio_handle,
                        now_vpn,
                        &stats,
                        &resolver,
                        &tunneled_paths,
                        &bus,
                        &mut local_interfaces,
                    ) {
                        passthrough.push(packet);
                    }
                }
                if !passthrough.is_empty() {
                    if let Err(error) = capture.send_ex(&passthrough) {
                        tracing::error!(
                            "failed to reinject {} unrelated packets: {}",
                            passthrough.len(),
                            error
                        );
                    }
                }
            }
            Err(error) => tracing::warn!("WinDivert receive failed: {}", error),
        }

        let _ = iface.poll(smoltcp::time::Instant::now(), &mut device, &mut sockets);
        drain_tx(&mut device, &inject, &stats, &local_interfaces);

        service_tcp_flows(&mut tcp_flows, &mut sockets, &tokio_handle, now_vpn, &stats);
        service_udp_flows(&mut udp_flows, &inject, &stats);

        let _ = iface.poll(smoltcp::time::Instant::now(), &mut device, &mut sockets);
        drain_tx(&mut device, &inject, &stats, &local_interfaces);
    }
}

fn on_vpn_change(
    prev: Option<VpnTarget>,
    now: Option<VpnTarget>,
    tcp_flows: &mut HashMap<TcpKey, TcpFlow>,
    udp_flows: &mut HashMap<UdpKey, UdpFlow>,
    sockets: &mut SocketSet<'static>,
    stats: &Stats,
) {
    if prev != now {
        let tcp_n = tcp_flows.len();
        let udp_n = udp_flows.len();
        if tcp_n + udp_n > 0 {
            tracing::warn!(
                "VPN target changed — closing {} TCP + {} UDP tunneled flows",
                tcp_n, udp_n
            );
        }
        for (_, flow) in tcp_flows.drain() {
            sockets.remove(flow.handle);
            stats.bridges_closed.fetch_add(1, Ordering::Relaxed);
        }
        // Dropping each flow also drops its cancellation sender, waking the
        // bridge even when its VPN socket is idle.
        let dropped = udp_flows.drain().count() as u64;
        stats.bridges_closed.fetch_add(dropped, Ordering::Relaxed);
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
struct TcpKey {
    src_ip: Ipv4Addr,
    src_port: u16,
    dst_ip: Ipv4Addr,
    dst_port: u16,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
struct UdpKey {
    src_ip: Ipv4Addr,
    src_port: u16,
    dst_ip: Ipv4Addr,
    dst_port: u16,
}

struct TcpFlow {
    handle: SocketHandle,
    pid: u32,
    original_dst: SocketAddr,
    state: TcpState,
}

enum TcpState {
    Pending,
    Bridging {
        app_to_remote: Option<mpsc::Sender<Vec<u8>>>,
        remote_to_app: mpsc::Receiver<Vec<u8>>,
        pending_to_app: Option<(Vec<u8>, usize)>,
        _cancel: oneshot::Sender<()>,
    },
    Closed,
}

struct UdpFlow {
    pid: u32,
    app_endpoint: (Ipv4Addr, u16),
    original_dst: (Ipv4Addr, u16),
    app_to_remote: mpsc::Sender<Vec<u8>>,
    remote_to_app: mpsc::Receiver<Vec<u8>>,
    injection_interface: InjectionInterface,
    _cancel: oneshot::Sender<()>,
    last_active: Instant,
}

#[derive(Debug, Clone, Copy)]
struct InjectionInterface {
    if_index: u32,
    sub_if_index: u32,
}

/// Resolve the owning PID for a freshly-seen flow. The fast path is the
/// observer-populated flow table; on miss we query the Windows TCP/UDP
/// table directly. Never wait or retry here: this runs on the global packet
/// loop, so even a small per-packet delay can stall unrelated applications.
fn resolve_pid(
    flows: &FlowTable,
    src_ip: IpAddr,
    src_port: u16,
    proto: u8,
) -> Option<u32> {
    let lookup = |flows: &FlowTable| -> Option<u32> {
        flows::lookup_with_wildcard(flows, src_ip, src_port, proto)
            .map(|e| e.pid)
            .or_else(|| pidlookup::pid_for_local(src_ip, src_port, proto))
    };

    lookup(flows)
}

#[allow(clippy::too_many_arguments)]
fn handle_capture<'a>(
    packet: WinDivertPacket<'a, NetworkLayer>,
    policy_state: &PolicyState,
    flows: &FlowTable,
    tcp_flows: &mut HashMap<TcpKey, TcpFlow>,
    udp_flows: &mut HashMap<UdpKey, UdpFlow>,
    sockets: &mut SocketSet<'static>,
    device: &mut VirtualDevice,
    tokio_handle: &tokio::runtime::Handle,
    vpn_target: Option<VpnTarget>,
    stats: &Stats,
    resolver: &Arc<Resolver>,
    tunneled_paths: &TunneledPaths,
    bus: &Arc<StatusBus>,
    local_interfaces: &mut HashMap<Ipv4Addr, InjectionInterface>,
) -> Option<WinDivertPacket<'a, NetworkLayer>> {
    if packet.data.first().is_some_and(|byte| byte >> 4 == 6) {
        return handle_ipv6_capture(
            packet,
            policy_state,
            flows,
            stats,
            resolver,
            tunneled_paths,
            bus,
        );
    }

    let parsed = match parse_ipv4_l4(&packet.data) {
        Some(p) => p,
        None => {
            stats.passed.fetch_add(1, Ordering::Relaxed);
            return Some(packet);
        }
    };

    let (src_ip, src_port, dst_ip, dst_port, is_tcp) = match parsed {
        ParsedL4::Tcp { src_ip, src_port, dst_ip, dst_port, .. } => {
            (src_ip, src_port, dst_ip, dst_port, true)
        }
        ParsedL4::Udp { src_ip, src_port, dst_ip, dst_port, .. } => {
            (src_ip, src_port, dst_ip, dst_port, false)
        }
    };

    let injection_interface = InjectionInterface {
        if_index: packet.address.interface_index(),
        sub_if_index: packet.address.subinterface_index(),
    };
    local_interfaces.insert(src_ip, injection_interface);

    let proto = if is_tcp { PROTO_TCP } else { PROTO_UDP };
    let pid = resolve_pid(flows, IpAddr::V4(src_ip), src_port, proto);
    if let Some(p) = pid {
        flows::insert(
            flows,
            LocalEndpoint { addr: IpAddr::V4(src_ip), port: src_port, proto },
            FlowEntry { pid: p },
        );
    }

    // Three-step admission:
    //   1. PID already in policy → tunnel
    //   2. PID has a known exe path matching the tunnel list → admit now
    //      (this is the path that closes the proc_watcher race — fast
    //      launchers like STOVE make their first connection before any
    //      polling tick would catch them)
    //   3. otherwise → pass through to the default route
    let is_tunneled = match pid {
        Some(p) if policy::contains(policy_state, p) => true,
        Some(p) => {
            if let Some(info) = resolver.resolve(p) {
                if tunneled::matches(tunneled_paths, std::path::Path::new(&info.exe_path)) {
                    tracing::info!(
                        "auto-tunnel (packet): pid={} matches {}",
                        p, info.exe_path
                    );
                    policy::add(policy_state, p);
                    bus.pid_paths.insert(p, info.exe_path);
                    true
                } else {
                    false
                }
            } else {
                false
            }
        }
        None => false,
    };

    if !is_tunneled {
        stats.passed.fetch_add(1, Ordering::Relaxed);
        return Some(packet);
    }
    let pid = pid.unwrap();

    // Fail closed. Once a process is selected, a tunnel outage must never
    // expose its real route. Existing bridge flows are torn down when the VPN
    // disappears; new packets are consumed until the tunnel is healthy.
    if vpn_target.is_none() {
        stats.captured.fetch_add(1, Ordering::Relaxed);
        return None;
    }

    if is_tcp {
        let key = TcpKey { src_ip, src_port, dst_ip, dst_port };
        if !tcp_flows.contains_key(&key) {
            let rx_buf = tcp::SocketBuffer::new(vec![0u8; 65536]);
            let tx_buf = tcp::SocketBuffer::new(vec![0u8; 65536]);
            let mut sock = tcp::Socket::new(rx_buf, tx_buf);
            let endpoint = IpEndpoint::new(IpAddress::Ipv4(dst_ip.into()), dst_port);
            if let Err(e) = sock.listen(endpoint) {
                tracing::warn!("smoltcp tcp::listen({}) failed: {:?}", endpoint, e);
                return None;
            }
            let handle = sockets.add(sock);
            tcp_flows.insert(
                key,
                TcpFlow {
                    handle,
                    pid,
                    original_dst: SocketAddr::V4(SocketAddrV4::new(dst_ip, dst_port)),
                    state: TcpState::Pending,
                },
            );
        }
        device.push_rx(packet.data.into_owned());
        stats.captured.fetch_add(1, Ordering::Relaxed);
    } else {
        let key = UdpKey { src_ip, src_port, dst_ip, dst_port };
        let payload = match extract_udp_payload(&packet.data) {
            Some(p) => p,
            None => {
                tracing::warn!("could not extract UDP payload");
                return None;
            }
        };

        if !udp_flows.contains_key(&key) {
            let (a2r_tx, a2r_rx) = mpsc::channel::<Vec<u8>>(BRIDGE_CHANNEL_DEPTH);
            let (r2a_tx, r2a_rx) = mpsc::channel::<Vec<u8>>(BRIDGE_CHANNEL_DEPTH);
            let (cancel_tx, cancel_rx) = oneshot::channel();

            tracing::info!(
                "udp bridge open: pid={} {}:{} -> {}:{}",
                pid, src_ip, src_port, dst_ip, dst_port
            );

            let dst = SocketAddr::V4(SocketAddrV4::new(dst_ip, dst_port));
            tokio_handle.spawn(async move {
                if let Err(e) = udp_bridge(dst, vpn_target, a2r_rx, r2a_tx, cancel_rx).await {
                    tracing::warn!("udp bridge to {} failed: {:#}", dst, e);
                }
            });
            stats.bridges_opened.fetch_add(1, Ordering::Relaxed);

            udp_flows.insert(
                key,
                UdpFlow {
                    pid,
                    app_endpoint: (src_ip, src_port),
                    original_dst: (dst_ip, dst_port),
                    app_to_remote: a2r_tx,
                    remote_to_app: r2a_rx,
                    injection_interface,
                    _cancel: cancel_tx,
                    last_active: Instant::now(),
                },
            );
        }

        if let Some(flow) = udp_flows.get_mut(&key) {
            flow.last_active = Instant::now();
            let len = packet.data.len() as u64;
            if flow.app_to_remote.try_send(payload).is_ok() {
                stats.add_app_to_remote(len);
                stats.add_pid_out(flow.pid, len);
            }
        }
        stats.captured.fetch_add(1, Ordering::Relaxed);
        // packet consumed - do NOT reinject
    }
    None
}

#[allow(clippy::too_many_arguments)]
fn handle_ipv6_capture<'a>(
    packet: WinDivertPacket<'a, NetworkLayer>,
    policy_state: &PolicyState,
    flows: &FlowTable,
    stats: &Stats,
    resolver: &Arc<Resolver>,
    tunneled_paths: &TunneledPaths,
    bus: &Arc<StatusBus>,
) -> Option<WinDivertPacket<'a, NetworkLayer>> {
    let Some((src_ip, src_port, proto)) = parse_ipv6_local_endpoint(&packet.data) else {
        stats.passed.fetch_add(1, Ordering::Relaxed);
        return Some(packet);
    };

    let pid = resolve_pid(flows, src_ip, src_port, proto);
    let is_tunneled = match pid {
        Some(pid) if policy::contains(policy_state, pid) => true,
        Some(pid) => resolver.resolve(pid).is_some_and(|info| {
            if !tunneled::matches(tunneled_paths, std::path::Path::new(&info.exe_path)) {
                return false;
            }
            policy::add(policy_state, pid);
            bus.pid_paths.insert(pid, info.exe_path);
            true
        }),
        None => false,
    };

    if is_tunneled {
        stats.captured.fetch_add(1, Ordering::Relaxed);
        None
    } else {
        stats.passed.fetch_add(1, Ordering::Relaxed);
        Some(packet)
    }
}

fn parse_ipv6_local_endpoint(data: &[u8]) -> Option<(IpAddr, u16, u8)> {
    let parsed = SlicedPacket::from_ip(data).ok()?;
    let src_ip = match parsed.net? {
        NetSlice::Ipv6(header) => IpAddr::V6(header.header().source_addr()),
        _ => return None,
    };
    let (src_port, proto) = match parsed.transport? {
        TransportSlice::Tcp(tcp) => (tcp.source_port(), PROTO_TCP),
        TransportSlice::Udp(udp) => (udp.source_port(), PROTO_UDP),
        _ => return None,
    };
    Some((src_ip, src_port, proto))
}

fn drain_tx(
    device: &mut VirtualDevice,
    inject: &WinDivert<NetworkLayer>,
    stats: &Stats,
    local_interfaces: &HashMap<Ipv4Addr, InjectionInterface>,
) {
    while let Some(data) = device.pop_tx() {
        let injection_interface = ipv4_destination(&data)
            .and_then(|destination| local_interfaces.get(&destination).copied());
        send_injected(inject, data, stats, injection_interface);
    }
}

fn send_injected(
    inject: &WinDivert<NetworkLayer>,
    data: Vec<u8>,
    stats: &Stats,
    injection_interface: Option<InjectionInterface>,
) {
    let Some(injection_interface) = injection_interface else {
        tracing::warn!("dropping synthesized reply without a valid inbound interface");
        return;
    };
    let mut addr = unsafe { WinDivertAddress::<NetworkLayer>::new() };
    addr.set_outbound(false);
    addr.set_interface_index(injection_interface.if_index);
    addr.set_subinterface_index(injection_interface.sub_if_index);
    addr.set_ip_checksum(false);
    addr.set_tcp_checksum(false);
    addr.set_udp_checksum(false);
    let mut pkt = WinDivertPacket {
        address: addr,
        data: Cow::Owned(data),
    };
    if let Err(e) = pkt.recalculate_checksums(ChecksumFlags::new()) {
        tracing::warn!("inject checksum recalc failed: {}", e);
        return;
    }
    if let Err(e) = inject.send(&pkt) {
        tracing::warn!("inject send failed: {}", e);
        return;
    }
    stats.injected.fetch_add(1, Ordering::Relaxed);
}

fn ipv4_destination(data: &[u8]) -> Option<Ipv4Addr> {
    if data.len() < 20 || data[0] >> 4 != 4 {
        return None;
    }
    Some(Ipv4Addr::new(data[16], data[17], data[18], data[19]))
}

fn service_tcp_flows(
    tcp_flows: &mut HashMap<TcpKey, TcpFlow>,
    sockets: &mut SocketSet<'static>,
    tokio_handle: &tokio::runtime::Handle,
    vpn_target: Option<VpnTarget>,
    stats: &Stats,
) {
    let mut to_remove = Vec::new();
    for (key, flow) in tcp_flows.iter_mut() {
        match &mut flow.state {
            TcpState::Pending => {
                let socket = sockets.get_mut::<tcp::Socket>(flow.handle);
                if socket.state() == tcp::State::Established {
                    let (a2r_tx, a2r_rx) = mpsc::channel::<Vec<u8>>(BRIDGE_CHANNEL_DEPTH);
                    let (r2a_tx, r2a_rx) = mpsc::channel::<Vec<u8>>(BRIDGE_CHANNEL_DEPTH);
                    let (cancel_tx, cancel_rx) = oneshot::channel();

                    tracing::info!(
                        "tcp bridge open: pid={} {}:{} -> {}",
                        flow.pid, key.src_ip, key.src_port, flow.original_dst
                    );

                    let dst = flow.original_dst;
                    tokio_handle.spawn(async move {
                        if let Err(e) = tcp_bridge(dst, vpn_target, a2r_rx, r2a_tx, cancel_rx).await {
                            tracing::warn!("tcp bridge to {} failed: {:#}", dst, e);
                        }
                    });
                    stats.bridges_opened.fetch_add(1, Ordering::Relaxed);

                    flow.state = TcpState::Bridging {
                        app_to_remote: Some(a2r_tx),
                        remote_to_app: r2a_rx,
                        pending_to_app: None,
                        _cancel: cancel_tx,
                    };
                }
            }
            TcpState::Bridging { app_to_remote, remote_to_app, pending_to_app, .. } => {
                let socket = sockets.get_mut::<tcp::Socket>(flow.handle);
                pump_tcp_app_to_remote(socket, app_to_remote, flow.pid, stats);
                pump_tcp_remote_to_app(socket, remote_to_app, pending_to_app, flow.pid, stats);
                if matches!(socket.state(), tcp::State::Closed | tcp::State::TimeWait) {
                    flow.state = TcpState::Closed;
                }
            }
            TcpState::Closed => to_remove.push(*key),
        }
    }
    for k in to_remove {
        if let Some(flow) = tcp_flows.remove(&k) {
            sockets.remove(flow.handle);
            stats.bridges_closed.fetch_add(1, Ordering::Relaxed);
        }
    }
}

fn service_udp_flows(
    udp_flows: &mut HashMap<UdpKey, UdpFlow>,
    inject: &WinDivert<NetworkLayer>,
    stats: &Stats,
) {
    let now = Instant::now();
    let mut to_remove = Vec::new();

    for (key, flow) in udp_flows.iter_mut() {
        loop {
            match flow.remote_to_app.try_recv() {
                Ok(payload) => {
                    flow.last_active = now;
                    let pkt = build_udp_packet(
                        flow.original_dst.0,
                        flow.original_dst.1,
                        flow.app_endpoint.0,
                        flow.app_endpoint.1,
                        &payload,
                    );
                    stats.add_remote_to_app(payload.len() as u64);
                    stats.add_pid_in(flow.pid, payload.len() as u64);
                    send_injected(inject, pkt, stats, Some(flow.injection_interface));
                }
                Err(mpsc::error::TryRecvError::Empty) => break,
                Err(mpsc::error::TryRecvError::Disconnected) => {
                    to_remove.push(*key);
                    break;
                }
            }
        }
        if now.duration_since(flow.last_active) > UDP_IDLE_TIMEOUT {
            to_remove.push(*key);
        }
    }

    for k in to_remove {
        if udp_flows.remove(&k).is_some() {
            stats.bridges_closed.fetch_add(1, Ordering::Relaxed);
        }
    }
}

fn pump_tcp_app_to_remote(
    socket: &mut tcp::Socket,
    sender: &mut Option<mpsc::Sender<Vec<u8>>>,
    pid: u32,
    stats: &Stats,
) {
    let mut channel_closed = false;
    if let Some(channel) = sender.as_ref() {
        while socket.can_recv() {
            let permit = match channel.try_reserve() {
                Ok(permit) => permit,
                Err(mpsc::error::TrySendError::Full(_)) => break,
                Err(mpsc::error::TrySendError::Closed(_)) => {
                    channel_closed = true;
                    break;
                }
            };
            let mut buf = vec![0u8; SMOLTCP_CHUNK];
            let n = match socket.recv_slice(&mut buf) {
                Ok(n) => n,
                Err(_) => break,
            };
            if n == 0 {
                break;
            }
            buf.truncate(n);
            permit.send(buf);
            stats.add_app_to_remote(n as u64);
            stats.add_pid_out(pid, n as u64);
        }
    }

    if channel_closed || matches!(
        socket.state(),
        tcp::State::CloseWait | tcp::State::LastAck | tcp::State::Closing
    ) {
        sender.take();
    }
}

fn pump_tcp_remote_to_app(
    socket: &mut tcp::Socket,
    receiver: &mut mpsc::Receiver<Vec<u8>>,
    pending: &mut Option<(Vec<u8>, usize)>,
    pid: u32,
    stats: &Stats,
) {
    while socket.can_send() {
        if pending.is_none() {
            match receiver.try_recv() {
                Ok(data) => *pending = Some((data, 0)),
                Err(mpsc::error::TryRecvError::Empty) => break,
                Err(mpsc::error::TryRecvError::Disconnected) => {
                    socket.close();
                    break;
                }
            }
        }

        let Some((data, written)) = pending.as_mut() else { break };
        match socket.send_slice(&data[*written..]) {
            Ok(0) => break,
            Ok(n) => {
                *written += n;
                stats.add_remote_to_app(n as u64);
                stats.add_pid_in(pid, n as u64);
                if *written == data.len() {
                    *pending = None;
                }
            }
            Err(_) => return,
        }
    }
}

/// Pin a TCP/UDP socket to a specific outgoing interface so the kernel ignores
/// the system routing table and egresses via the chosen NIC. Combined with
/// `--pull-filter ignore redirect-gateway` on the OpenVPN side, this is what
/// turns the tunnel into a true split-tunnel — only sockets we explicitly pin
/// land on the VPN, everything else stays on the default route.
fn set_unicast_if_v4<S: std::os::windows::io::AsRawSocket>(socket: &S, if_index: u32) -> Result<()> {
    use windows::Win32::Networking::WinSock::{setsockopt, IPPROTO_IP, SOCKET};

    // IP_UNICAST_IF wants the index in network byte order.
    let val = if_index.to_be();
    let bytes = val.to_ne_bytes();
    let s = SOCKET(socket.as_raw_socket() as usize);
    // IP_UNICAST_IF is not exposed as a constant in the windows crate at our
    // version — its numeric value is 31. Defined this way in the WinSock SDK.
    const IP_UNICAST_IF: i32 = 31;
    let rc = unsafe { setsockopt(s, IPPROTO_IP.0, IP_UNICAST_IF, Some(&bytes)) };
    if rc != 0 {
        anyhow::bail!(
            "setsockopt(IP_UNICAST_IF, ifindex={}) failed: {}",
            if_index,
            std::io::Error::last_os_error()
        );
    }
    Ok(())
}

async fn tcp_bridge(
    dst: SocketAddr,
    target: Option<VpnTarget>,
    mut from_app: mpsc::Receiver<Vec<u8>>,
    to_app: mpsc::Sender<Vec<u8>>,
    mut cancel: oneshot::Receiver<()>,
) -> Result<()> {
    let stream = match target {
        Some(t) => {
            let socket = tokio::net::TcpSocket::new_v4()?;
            set_unicast_if_v4(&socket, t.if_index)?;
            socket.bind(SocketAddr::V4(SocketAddrV4::new(t.ipv4, 0)))?;
            socket
                .connect(dst)
                .await
                .with_context(|| format!("tcp connect({}) via ifindex {}", dst, t.if_index))?
        }
        None => tokio::net::TcpStream::connect(dst)
            .await
            .with_context(|| format!("tcp connect({})", dst))?,
    };
    tracing::debug!("tcp bridge connected to {} (local={:?})", dst, stream.local_addr().ok());
    let (mut rd, mut wr) = stream.into_split();

    let mut write_task = tokio::spawn(async move {
        while let Some(data) = from_app.recv().await {
            if wr.write_all(&data).await.is_err() {
                break;
            }
        }
        let _ = wr.shutdown().await;
    });

    let mut read_task = tokio::spawn(async move {
        let mut buf = vec![0u8; SMOLTCP_CHUNK];
        loop {
            match rd.read(&mut buf).await {
                Ok(0) => break,
                Ok(n) => {
                    if to_app.send(buf[..n].to_vec()).await.is_err() {
                        break;
                    }
                }
                Err(_) => break,
            }
        }
    });

    tokio::select! {
        _ = &mut cancel => {
            write_task.abort();
            read_task.abort();
        }
        _ = &mut read_task => {
            write_task.abort();
            let _ = write_task.await;
        }
        _ = &mut write_task => {
            let _ = read_task.await;
        }
    }
    Ok(())
}

async fn udp_bridge(
    dst: SocketAddr,
    target: Option<VpnTarget>,
    mut from_app: mpsc::Receiver<Vec<u8>>,
    to_app: mpsc::Sender<Vec<u8>>,
    mut cancel: oneshot::Receiver<()>,
) -> Result<()> {
    let local = match target {
        Some(t) => SocketAddr::V4(SocketAddrV4::new(t.ipv4, 0)),
        None => SocketAddr::V4(SocketAddrV4::new(Ipv4Addr::UNSPECIFIED, 0)),
    };
    let socket = tokio::net::UdpSocket::bind(local)
        .await
        .with_context(|| format!("udp bind({})", local))?;
    if let Some(t) = target {
        set_unicast_if_v4(&socket, t.if_index)?;
    }
    socket.connect(dst).await
        .with_context(|| format!("udp connect({})", dst))?;
    tracing::debug!("udp bridge {} -> {} (local={:?})", local, dst, socket.local_addr().ok());
    let mut buf = vec![0u8; 65536];
    loop {
        tokio::select! {
            _ = &mut cancel => break,
            outbound = from_app.recv() => {
                let Some(data) = outbound else { break };
                match socket.try_send(&data) {
                    Ok(_) => {}
                    Err(error) if error.kind() == std::io::ErrorKind::WouldBlock => {}
                    Err(error) => return Err(error.into()),
                }
            }
            inbound = socket.recv(&mut buf) => {
                let n = inbound?;
                if n == 0 {
                    break;
                }
                match to_app.try_send(buf[..n].to_vec()) {
                    Ok(()) | Err(mpsc::error::TrySendError::Full(_)) => {}
                    Err(mpsc::error::TrySendError::Closed(_)) => break,
                }
            }
        }
    }
    Ok(())
}

enum ParsedL4 {
    Tcp {
        src_ip: Ipv4Addr,
        src_port: u16,
        dst_ip: Ipv4Addr,
        dst_port: u16,
    },
    Udp {
        src_ip: Ipv4Addr,
        src_port: u16,
        dst_ip: Ipv4Addr,
        dst_port: u16,
    },
}

fn parse_ipv4_l4(data: &[u8]) -> Option<ParsedL4> {
    let sliced = SlicedPacket::from_ip(data).ok()?;
    let net = sliced.net?;
    let ipv4 = match net {
        NetSlice::Ipv4(ip) => ip,
        _ => return None,
    };
    let src_ip = ipv4.header().source_addr();
    let dst_ip = ipv4.header().destination_addr();
    match sliced.transport? {
        TransportSlice::Tcp(tcp) => Some(ParsedL4::Tcp {
            src_ip,
            src_port: tcp.source_port(),
            dst_ip,
            dst_port: tcp.destination_port(),
        }),
        TransportSlice::Udp(udp_slice) => Some(ParsedL4::Udp {
            src_ip,
            src_port: udp_slice.source_port(),
            dst_ip,
            dst_port: udp_slice.destination_port(),
        }),
        _ => None,
    }
}

fn extract_udp_payload(data: &[u8]) -> Option<Vec<u8>> {
    let sliced = SlicedPacket::from_ip(data).ok()?;
    match sliced.transport? {
        TransportSlice::Udp(udp_slice) => Some(udp_slice.payload().to_vec()),
        _ => None,
    }
}

fn build_udp_packet(
    src_ip: Ipv4Addr,
    src_port: u16,
    dst_ip: Ipv4Addr,
    dst_port: u16,
    payload: &[u8],
) -> Vec<u8> {
    let total_len = 20 + 8 + payload.len();
    let udp_len = 8 + payload.len();
    let mut buf = Vec::with_capacity(total_len);

    // IPv4 header (20 bytes, no options, IHL=5)
    buf.push(0x45);
    buf.push(0);
    buf.extend_from_slice(&(total_len as u16).to_be_bytes());
    buf.extend_from_slice(&[0, 0]); // identification
    buf.extend_from_slice(&[0, 0]); // flags + fragment offset
    buf.push(64); // ttl
    buf.push(17); // protocol = UDP
    buf.extend_from_slice(&[0, 0]); // header checksum (WinDivert will recompute)
    buf.extend_from_slice(&src_ip.octets());
    buf.extend_from_slice(&dst_ip.octets());

    // UDP header (8 bytes)
    buf.extend_from_slice(&src_port.to_be_bytes());
    buf.extend_from_slice(&dst_port.to_be_bytes());
    buf.extend_from_slice(&(udp_len as u16).to_be_bytes());
    buf.extend_from_slice(&[0, 0]); // checksum (WinDivert will recompute)

    buf.extend_from_slice(payload);
    buf
}

struct Stats {
    passed: AtomicU64,
    captured: AtomicU64,
    injected: AtomicU64,
    bridges_opened: AtomicU64,
    bridges_closed: AtomicU64,
    bus: Arc<StatusBus>,
}

impl Stats {
    fn new(bus: Arc<StatusBus>) -> Self {
        Self {
            passed: AtomicU64::new(0),
            captured: AtomicU64::new(0),
            injected: AtomicU64::new(0),
            bridges_opened: AtomicU64::new(0),
            bridges_closed: AtomicU64::new(0),
            bus,
        }
    }

    fn add_app_to_remote(&self, n: u64) {
        self.bus.bytes_out.fetch_add(n, Ordering::Relaxed);
    }
    fn add_remote_to_app(&self, n: u64) {
        self.bus.bytes_in.fetch_add(n, Ordering::Relaxed);
    }
    fn add_pid_out(&self, pid: u32, n: u64) {
        self.bus.add_pid_out(pid, n);
    }
    fn add_pid_in(&self, pid: u32, n: u64) {
        self.bus.add_pid_in(pid, n);
    }

    fn report_loop(&self) {
        let start = Instant::now();
        let mut last = [0u64; 7];
        loop {
            std::thread::sleep(Duration::from_secs(5));
            let cur = [
                self.passed.load(Ordering::Relaxed),
                self.captured.load(Ordering::Relaxed),
                self.injected.load(Ordering::Relaxed),
                self.bridges_opened.load(Ordering::Relaxed),
                self.bridges_closed.load(Ordering::Relaxed),
                self.bus.bytes_out.load(Ordering::Relaxed),
                self.bus.bytes_in.load(Ordering::Relaxed),
            ];
            let d: Vec<u64> = cur.iter().zip(&last).map(|(a, b)| a - b).collect();
            last = cur;
            tracing::info!(
                "divert 5s: pass={} cap={} inj={} bridges +{}/-{} a→r={}B r→a={}B (uptime {:?})",
                d[0], d[1], d[2], d[3], d[4], d[5], d[6], start.elapsed()
            );
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn udp_reply_targets_the_original_application_endpoint() {
        let packet = build_udp_packet(
            Ipv4Addr::new(10, 8, 0, 1),
            3074,
            Ipv4Addr::new(192, 168, 1, 25),
            55000,
            b"halo",
        );
        assert_eq!(ipv4_destination(&packet), Some(Ipv4Addr::new(192, 168, 1, 25)));
        match parse_ipv4_l4(&packet) {
            Some(ParsedL4::Udp { src_port, dst_port, .. }) => {
                assert_eq!(src_port, 3074);
                assert_eq!(dst_port, 55000);
            }
            _ => panic!("expected a UDP packet"),
        }
    }

    #[test]
    fn non_ipv4_data_has_no_injection_interface_key() {
        assert_eq!(ipv4_destination(&[]), None);
        assert_eq!(ipv4_destination(&[0x60; 20]), None);
    }
}
