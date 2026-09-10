use std::collections::VecDeque;

use smoltcp::phy::{Checksum, ChecksumCapabilities, Device, DeviceCapabilities, Medium, RxToken, TxToken};
use smoltcp::time::Instant;

pub struct VirtualDevice {
    rx: VecDeque<Vec<u8>>,
    tx: VecDeque<Vec<u8>>,
}

impl VirtualDevice {
    pub fn new() -> Self {
        Self {
            rx: VecDeque::new(),
            tx: VecDeque::new(),
        }
    }

    pub fn push_rx(&mut self, packet: Vec<u8>) {
        self.rx.push_back(packet);
    }

    pub fn pop_tx(&mut self) -> Option<Vec<u8>> {
        self.tx.pop_front()
    }
}

impl Device for VirtualDevice {
    type RxToken<'a> = VirtRxToken;
    type TxToken<'a> = VirtTxToken<'a>;

    fn capabilities(&self) -> DeviceCapabilities {
        let mut caps = DeviceCapabilities::default();
        caps.medium = Medium::Ip;
        caps.max_transmission_unit = 1500;
        let mut ck = ChecksumCapabilities::default();
        ck.ipv4 = Checksum::Tx;
        ck.tcp = Checksum::Tx;
        ck.udp = Checksum::Tx;
        caps.checksum = ck;
        caps
    }

    fn receive(&mut self, _: Instant) -> Option<(VirtRxToken, VirtTxToken<'_>)> {
        let pkt = self.rx.pop_front()?;
        Some((
            VirtRxToken { data: pkt },
            VirtTxToken { queue: &mut self.tx },
        ))
    }

    fn transmit(&mut self, _: Instant) -> Option<VirtTxToken<'_>> {
        Some(VirtTxToken {
            queue: &mut self.tx,
        })
    }
}

pub struct VirtRxToken {
    data: Vec<u8>,
}

impl RxToken for VirtRxToken {
    fn consume<R, F: FnOnce(&[u8]) -> R>(self, f: F) -> R {
        f(&self.data)
    }
}

pub struct VirtTxToken<'a> {
    queue: &'a mut VecDeque<Vec<u8>>,
}

impl<'a> TxToken for VirtTxToken<'a> {
    fn consume<R, F: FnOnce(&mut [u8]) -> R>(self, len: usize, f: F) -> R {
        let mut buf = vec![0u8; len];
        let r = f(&mut buf);
        self.queue.push_back(buf);
        r
    }
}
