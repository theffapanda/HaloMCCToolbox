pub mod policy {
    include!(concat!(env!("OUT_DIR"), "/vpnclient.policy.rs"));
}

pub mod status {
    include!(concat!(env!("OUT_DIR"), "/vpnclient.status.rs"));
}

pub use policy::{policy_message::Body, AddPid, PolicyMessage, RemovePid, Snapshot};

use prost::Message;

pub const PIPE_NAME: &str = r"\\.\pipe\halo-mcc-toolbox-vpn-policy";
pub const STATUS_PIPE_NAME: &str = r"\\.\pipe\halo-mcc-toolbox-vpn-status";

#[derive(Debug, thiserror::Error)]
pub enum FrameError {
    #[error("io error: {0}")]
    Io(#[from] std::io::Error),
    #[error("protobuf decode error: {0}")]
    Decode(#[from] prost::DecodeError),
    #[error("frame too large: {0} bytes (limit {1})")]
    TooLarge(u32, u32),
}

pub const MAX_FRAME_LEN: u32 = 1 << 20;

pub fn encode_frame(msg: &PolicyMessage, out: &mut Vec<u8>) {
    let len = msg.encoded_len();
    out.reserve(4 + len);
    out.extend_from_slice(&(len as u32).to_le_bytes());
    msg.encode(out).expect("Vec<u8> cannot fail encoding");
}

pub fn encode_status_frame(msg: &status::StatusMessage, out: &mut Vec<u8>) {
    let len = msg.encoded_len();
    out.reserve(4 + len);
    out.extend_from_slice(&(len as u32).to_le_bytes());
    msg.encode(out).expect("Vec<u8> cannot fail encoding");
}

pub fn read_frame<R: std::io::Read>(reader: &mut R) -> Result<PolicyMessage, FrameError> {
    let mut len_buf = [0u8; 4];
    reader.read_exact(&mut len_buf)?;
    let len = u32::from_le_bytes(len_buf);
    if len > MAX_FRAME_LEN {
        return Err(FrameError::TooLarge(len, MAX_FRAME_LEN));
    }
    let mut body = vec![0u8; len as usize];
    reader.read_exact(&mut body)?;
    Ok(PolicyMessage::decode(body.as_slice())?)
}

pub fn add_pid(pid: u32) -> PolicyMessage {
    PolicyMessage {
        body: Some(Body::AddPid(AddPid { pid })),
    }
}

pub fn remove_pid(pid: u32) -> PolicyMessage {
    PolicyMessage {
        body: Some(Body::RemovePid(RemovePid { pid })),
    }
}

pub fn snapshot(pids: Vec<u32>) -> PolicyMessage {
    PolicyMessage {
        body: Some(Body::Snapshot(Snapshot { pids })),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn roundtrip_add_pid() {
        let msg = add_pid(1234);
        let mut buf = Vec::new();
        encode_frame(&msg, &mut buf);
        let decoded = read_frame(&mut buf.as_slice()).unwrap();
        match decoded.body {
            Some(Body::AddPid(AddPid { pid })) => assert_eq!(pid, 1234),
            other => panic!("unexpected body: {:?}", other),
        }
    }

    #[test]
    fn roundtrip_snapshot() {
        let msg = snapshot(vec![1, 2, 3, 99]);
        let mut buf = Vec::new();
        encode_frame(&msg, &mut buf);
        let decoded = read_frame(&mut buf.as_slice()).unwrap();
        match decoded.body {
            Some(Body::Snapshot(Snapshot { pids })) => assert_eq!(pids, vec![1, 2, 3, 99]),
            other => panic!("unexpected body: {:?}", other),
        }
    }
}
