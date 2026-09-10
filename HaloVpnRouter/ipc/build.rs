use std::io::Result;

fn main() -> Result<()> {
    std::env::set_var("PROTOC", protoc_bin_vendored::protoc_bin_path().unwrap());
    prost_build::compile_protos(
        &["proto/policy.proto", "proto/status.proto"],
        &["proto"],
    )?;
    println!("cargo:rerun-if-changed=proto/policy.proto");
    println!("cargo:rerun-if-changed=proto/status.proto");
    Ok(())
}
