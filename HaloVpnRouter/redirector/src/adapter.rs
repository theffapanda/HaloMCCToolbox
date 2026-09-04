use std::net::{IpAddr, Ipv4Addr, Ipv6Addr};

use anyhow::{bail, Result};
use windows::Win32::Foundation::ERROR_BUFFER_OVERFLOW;
use windows::Win32::NetworkManagement::IpHelper::{
    GetAdaptersAddresses, GAA_FLAG_SKIP_ANYCAST, GAA_FLAG_SKIP_DNS_SERVER,
    GAA_FLAG_SKIP_MULTICAST, IP_ADAPTER_ADDRESSES_LH,
};
use windows::Win32::Networking::WinSock::{AF_INET, AF_INET6, AF_UNSPEC, SOCKADDR_IN, SOCKADDR_IN6};

#[derive(Debug, Clone)]
pub struct Adapter {
    pub friendly_name: String,
    pub description: String,
    pub adapter_name: String,
    pub addresses: Vec<IpAddr>,
    pub is_up: bool,
    /// IPv4 interface index from IP_ADAPTER_ADDRESSES_LH.IfIndex. Used for
    /// IP_UNICAST_IF socket-option pinning.
    pub if_index: u32,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct VpnTarget {
    pub ipv4: std::net::Ipv4Addr,
    pub if_index: u32,
}

impl Adapter {
    pub fn is_vpn_candidate(&self) -> bool {
        // wintun is our preferred path, but openvpn 2.7 falls back to
        // TAP-Windows6 when wintun isn't installed on the box — we still
        // need to recognize that as our tunnel. Disambiguation from the
        // OpenVPN Connect commercial client's TAP/DCO adapters is done by
        // the is_up gate in find_vpn_target: theirs sit Disconnected when
        // their app isn't running.
        let d = self.description.to_lowercase();
        d.contains("wintun") || d.contains("tap-windows") || d.contains("openvpn")
    }

    pub fn routable_ipv4(&self) -> Option<std::net::Ipv4Addr> {
        self.addresses.iter().find_map(|ip| match ip {
            IpAddr::V4(v4)
                if !v4.is_link_local() && !v4.is_unspecified() && !v4.is_loopback() =>
            {
                Some(*v4)
            }
            _ => None,
        })
    }
}

pub fn find_vpn_target(adapters: &[Adapter]) -> Option<VpnTarget> {
    // Without the is_up gate, an adapter with a stale IPv4 from a previous
    // session reads as "VPN up" even after openvpn has exited — the redirector
    // would then keep tunneling flows to a dead interface and report a
    // bogus uptime to the UI.
    for a in adapters.iter().filter(|a| a.is_vpn_candidate() && a.is_up) {
        if let Some(ip) = a.routable_ipv4() {
            return Some(VpnTarget {
                ipv4: ip,
                if_index: a.if_index,
            });
        }
    }
    None
}

pub fn enumerate() -> Result<Vec<Adapter>> {
    let flags = GAA_FLAG_SKIP_ANYCAST | GAA_FLAG_SKIP_MULTICAST | GAA_FLAG_SKIP_DNS_SERVER;
    let mut size: u32 = 16 * 1024;
    let mut buf: Vec<u8> = Vec::new();

    for _ in 0..4 {
        buf.resize(size as usize, 0);
        let rc = unsafe {
            GetAdaptersAddresses(
                AF_UNSPEC.0 as u32,
                flags,
                None,
                Some(buf.as_mut_ptr() as *mut IP_ADAPTER_ADDRESSES_LH),
                &mut size,
            )
        };
        match rc {
            0 => return Ok(parse(&buf)),
            x if x == ERROR_BUFFER_OVERFLOW.0 => continue,
            x => bail!("GetAdaptersAddresses failed: 0x{:x}", x),
        }
    }
    bail!("GetAdaptersAddresses kept asking for more buffer space")
}

fn parse(buf: &[u8]) -> Vec<Adapter> {
    let mut out = Vec::new();
    let mut cur = buf.as_ptr() as *const IP_ADAPTER_ADDRESSES_LH;
    while !cur.is_null() {
        let a = unsafe { &*cur };
        let friendly_name = unsafe { a.FriendlyName.to_string() }.unwrap_or_default();
        let description = unsafe { a.Description.to_string() }.unwrap_or_default();
        let adapter_name = unsafe { a.AdapterName.to_string() }.unwrap_or_default();
        let is_up = a.OperStatus.0 == 1;
        let if_index = unsafe { a.Anonymous1.Anonymous.IfIndex };

        let mut addrs = Vec::new();
        let mut uc = a.FirstUnicastAddress;
        while !uc.is_null() {
            let u = unsafe { &*uc };
            let sa = u.Address.lpSockaddr;
            if !sa.is_null() {
                let family = unsafe { (*sa).sa_family };
                if family == AF_INET {
                    let v4 = unsafe { &*(sa as *const SOCKADDR_IN) };
                    let raw = unsafe { v4.sin_addr.S_un.S_addr };
                    addrs.push(IpAddr::V4(Ipv4Addr::from(raw.to_ne_bytes())));
                } else if family == AF_INET6 {
                    let v6 = unsafe { &*(sa as *const SOCKADDR_IN6) };
                    let bytes = unsafe { v6.sin6_addr.u.Byte };
                    addrs.push(IpAddr::V6(Ipv6Addr::from(bytes)));
                }
            }
            uc = u.Next;
        }

        out.push(Adapter {
            friendly_name,
            description,
            adapter_name,
            addresses: addrs,
            is_up,
            if_index,
        });
        cur = a.Next;
    }
    out
}
