use std::net::{IpAddr, Ipv4Addr, Ipv6Addr};

use windows::Win32::Foundation::NO_ERROR;
use windows::Win32::NetworkManagement::IpHelper::{
    GetExtendedTcpTable, GetExtendedUdpTable, MIB_TCP6ROW_OWNER_PID, MIB_TCP6TABLE_OWNER_PID,
    MIB_TCPROW_OWNER_PID, MIB_TCPTABLE_OWNER_PID, MIB_UDP6ROW_OWNER_PID,
    MIB_UDP6TABLE_OWNER_PID, MIB_UDPROW_OWNER_PID, MIB_UDPTABLE_OWNER_PID,
    TCP_TABLE_OWNER_PID_ALL, UDP_TABLE_OWNER_PID,
};
use windows::Win32::Networking::WinSock::{AF_INET, AF_INET6};

/// Synchronous fallback when the observer-populated flow table hasn't caught
/// up with a fresh socket yet. Hits the Windows IP helper to find the PID
/// that owns (proto, local_addr, local_port).
///
/// Cost: ~hundreds of microseconds. We only call it on first-packet misses,
/// because the second the flow lands in the policy, every subsequent packet
/// from that PID takes the fast path.
pub fn pid_for_local_v4(local_addr: Ipv4Addr, local_port: u16, proto: u8) -> Option<u32> {
    match proto {
        6 => pid_for_tcp_v4(local_addr, local_port),
        17 => pid_for_udp_v4(local_addr, local_port),
        _ => None,
    }
}

pub fn pid_for_local(local_addr: IpAddr, local_port: u16, proto: u8) -> Option<u32> {
    match local_addr {
        IpAddr::V4(address) => pid_for_local_v4(address, local_port, proto),
        IpAddr::V6(address) => match proto {
            6 => pid_for_tcp_v6(address, local_port),
            17 => pid_for_udp_v6(address, local_port),
            _ => None,
        },
    }
}

fn pid_for_tcp_v4(local_addr: Ipv4Addr, local_port: u16) -> Option<u32> {
    let mut size: u32 = 0;
    unsafe {
        let _ = GetExtendedTcpTable(
            None,
            &mut size,
            false,
            AF_INET.0 as u32,
            TCP_TABLE_OWNER_PID_ALL,
            0,
        );
    }
    if size == 0 {
        return None;
    }
    let mut buf = vec![0u8; size as usize];
    let rc = unsafe {
        GetExtendedTcpTable(
            Some(buf.as_mut_ptr() as *mut _),
            &mut size,
            false,
            AF_INET.0 as u32,
            TCP_TABLE_OWNER_PID_ALL,
            0,
        )
    };
    if rc != NO_ERROR.0 {
        return None;
    }
    let table = unsafe { &*(buf.as_ptr() as *const MIB_TCPTABLE_OWNER_PID) };
    let count = table.dwNumEntries as usize;
    if count == 0 {
        return None;
    }
    let rows: &[MIB_TCPROW_OWNER_PID] =
        unsafe { std::slice::from_raw_parts(table.table.as_ptr(), count) };
    let target_addr_be = u32::from(local_addr).to_be();
    for row in rows {
        let port = port_from_dword(row.dwLocalPort);
        if port != local_port {
            continue;
        }
        if row.dwLocalAddr != 0 && row.dwLocalAddr != target_addr_be {
            continue;
        }
        return Some(row.dwOwningPid);
    }
    None
}

fn pid_for_udp_v4(local_addr: Ipv4Addr, local_port: u16) -> Option<u32> {
    let mut size: u32 = 0;
    unsafe {
        let _ = GetExtendedUdpTable(
            None,
            &mut size,
            false,
            AF_INET.0 as u32,
            UDP_TABLE_OWNER_PID,
            0,
        );
    }
    if size == 0 {
        return None;
    }
    let mut buf = vec![0u8; size as usize];
    let rc = unsafe {
        GetExtendedUdpTable(
            Some(buf.as_mut_ptr() as *mut _),
            &mut size,
            false,
            AF_INET.0 as u32,
            UDP_TABLE_OWNER_PID,
            0,
        )
    };
    if rc != NO_ERROR.0 {
        return None;
    }
    let table = unsafe { &*(buf.as_ptr() as *const MIB_UDPTABLE_OWNER_PID) };
    let count = table.dwNumEntries as usize;
    if count == 0 {
        return None;
    }
    let rows: &[MIB_UDPROW_OWNER_PID] =
        unsafe { std::slice::from_raw_parts(table.table.as_ptr(), count) };
    let target_addr_be = u32::from(local_addr).to_be();
    for row in rows {
        let port = port_from_dword(row.dwLocalPort);
        if port != local_port {
            continue;
        }
        if row.dwLocalAddr != 0 && row.dwLocalAddr != target_addr_be {
            continue;
        }
        return Some(row.dwOwningPid);
    }
    None
}

fn pid_for_tcp_v6(local_addr: Ipv6Addr, local_port: u16) -> Option<u32> {
    let mut size: u32 = 0;
    unsafe {
        let _ = GetExtendedTcpTable(
            None,
            &mut size,
            false,
            AF_INET6.0 as u32,
            TCP_TABLE_OWNER_PID_ALL,
            0,
        );
    }
    if size == 0 {
        return None;
    }
    let mut buf = vec![0u8; size as usize];
    let rc = unsafe {
        GetExtendedTcpTable(
            Some(buf.as_mut_ptr() as *mut _),
            &mut size,
            false,
            AF_INET6.0 as u32,
            TCP_TABLE_OWNER_PID_ALL,
            0,
        )
    };
    if rc != NO_ERROR.0 {
        return None;
    }
    let table = unsafe { &*(buf.as_ptr() as *const MIB_TCP6TABLE_OWNER_PID) };
    let rows: &[MIB_TCP6ROW_OWNER_PID] = unsafe {
        std::slice::from_raw_parts(table.table.as_ptr(), table.dwNumEntries as usize)
    };
    let target = local_addr.octets();
    rows.iter()
        .find(|row| {
            port_from_dword(row.dwLocalPort) == local_port &&
                (row.ucLocalAddr == [0; 16] || row.ucLocalAddr == target)
        })
        .map(|row| row.dwOwningPid)
}

fn pid_for_udp_v6(local_addr: Ipv6Addr, local_port: u16) -> Option<u32> {
    let mut size: u32 = 0;
    unsafe {
        let _ = GetExtendedUdpTable(
            None,
            &mut size,
            false,
            AF_INET6.0 as u32,
            UDP_TABLE_OWNER_PID,
            0,
        );
    }
    if size == 0 {
        return None;
    }
    let mut buf = vec![0u8; size as usize];
    let rc = unsafe {
        GetExtendedUdpTable(
            Some(buf.as_mut_ptr() as *mut _),
            &mut size,
            false,
            AF_INET6.0 as u32,
            UDP_TABLE_OWNER_PID,
            0,
        )
    };
    if rc != NO_ERROR.0 {
        return None;
    }
    let table = unsafe { &*(buf.as_ptr() as *const MIB_UDP6TABLE_OWNER_PID) };
    let rows: &[MIB_UDP6ROW_OWNER_PID] = unsafe {
        std::slice::from_raw_parts(table.table.as_ptr(), table.dwNumEntries as usize)
    };
    let target = local_addr.octets();
    rows.iter()
        .find(|row| {
            port_from_dword(row.dwLocalPort) == local_port &&
                (row.ucLocalAddr == [0; 16] || row.ucLocalAddr == target)
        })
        .map(|row| row.dwOwningPid)
}

/// dwLocalPort is a DWORD whose low 16 bits hold the port in network byte
/// order. Mask + byte-swap to get the host-order port.
fn port_from_dword(d: u32) -> u16 {
    u16::from_be((d & 0xFFFF) as u16)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn finds_current_process_tcp_listener() {
        let listener = std::net::TcpListener::bind((Ipv4Addr::LOCALHOST, 0)).unwrap();
        let port = listener.local_addr().unwrap().port();
        assert_eq!(
            pid_for_local_v4(Ipv4Addr::LOCALHOST, port, 6),
            Some(std::process::id())
        );
    }

    #[test]
    fn finds_current_process_udp_socket() {
        let socket = std::net::UdpSocket::bind((Ipv4Addr::LOCALHOST, 0)).unwrap();
        let port = socket.local_addr().unwrap().port();
        assert_eq!(
            pid_for_local_v4(Ipv4Addr::LOCALHOST, port, 17),
            Some(std::process::id())
        );
    }
}
