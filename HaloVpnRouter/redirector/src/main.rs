mod adapter;
mod divert;
mod flows;
mod observer;
mod pidlookup;
mod policy;
mod proc_watcher;
mod process;
mod stack;
mod status_server;
mod tunneled;
mod vpn_state;

use std::sync::Arc;

use anyhow::Result;

fn main() -> Result<()> {
    tracing_subscriber::fmt()
        .with_ansi(false)
        .with_env_filter(
            tracing_subscriber::EnvFilter::try_from_default_env()
                .unwrap_or_else(|_| tracing_subscriber::EnvFilter::new("info")),
        )
        .init();

    let args: Vec<String> = std::env::args().collect();
    match args.get(1).map(|s| s.as_str()) {
        None | Some("adapters") => print_adapters(),
        Some("observe") => run_observe(),
        Some("help") | Some("--help") | Some("-h") => {
            print_help();
            Ok(())
        }
        Some(cmd) => {
            eprintln!("unknown command: {}", cmd);
            print_help();
            std::process::exit(2);
        }
    }
}

fn run_observe() -> Result<()> {
    let rt = tokio::runtime::Builder::new_multi_thread()
        .enable_all()
        .build()?;
    rt.block_on(async {
        let policy_state = policy::new();
        let flows = flows::new();
        let resolver = Arc::new(process::Resolver::new());
        let tunneled_paths = tunneled::new();
        tunneled::replace(&tunneled_paths, tunneled::load_from_disk());

        let vpn_state = vpn_state::new();
        let bus = Arc::new(status_server::StatusBus::new());
        proc_watcher::seed(&policy_state, &bus, &tunneled_paths);

        let vpn_task = tokio::spawn(vpn_state::watcher(vpn_state.clone(), bus.clone()));
        let watcher_task = tokio::spawn(proc_watcher::run(
            policy_state.clone(),
            bus.clone(),
            resolver.clone(),
            flows.clone(),
            tunneled_paths.clone(),
        ));
        let observer_handle = {
            let resolver = resolver.clone();
            let policy_state = policy_state.clone();
            let flows = flows.clone();
            tokio::task::spawn_blocking(move || observer::run(resolver, policy_state, flows))
        };
        let divert_handle = {
            let policy_state = policy_state.clone();
            let flows = flows.clone();
            let vpn_state = vpn_state.clone();
            let bus = bus.clone();
            let resolver = resolver.clone();
            let tunneled_paths = tunneled_paths.clone();
            tokio::task::spawn_blocking(move || {
                divert::run(policy_state, flows, vpn_state, bus, resolver, tunneled_paths)
            })
        };

        tokio::select! {
            r = vpn_task => r??,
            r = watcher_task => r??,
            r = observer_handle => r??,
            r = divert_handle => r??,
        }
        Ok(())
    })
}

fn print_help() {
    println!("HaloVpnRouter - Halo MCC per-app VPN routing service");
    println!();
    println!("USAGE:");
    println!("  HaloVpnRouter [adapters]    List network adapters and VPN candidates");
    println!("  HaloVpnRouter observe       Route configured MCC processes (requires admin)");
    println!("  HaloVpnRouter help          Show this message");
}

fn print_adapters() -> Result<()> {
    let adapters = adapter::enumerate()?;
    println!("Found {} adapters\n", adapters.len());

    for a in &adapters {
        let status = if a.is_up { "UP  " } else { "DOWN" };
        let marker = if a.is_vpn_candidate() { "  <-- VPN candidate" } else { "" };
        println!("[{}] {}{}", status, a.friendly_name, marker);
        println!("       description: {}", a.description);
        println!("       guid:        {}", a.adapter_name);
        if a.addresses.is_empty() {
            println!("       (no addresses)");
        } else {
            for ip in &a.addresses {
                println!("       address:     {}", ip);
            }
        }
        println!();
    }

    let vpn: Vec<_> = adapters
        .iter()
        .filter(|a| a.is_vpn_candidate() && a.is_up)
        .collect();
    match vpn.as_slice() {
        [] => println!("No active VPN adapter detected. Start OpenVPN before continuing."),
        [one] => println!(
            "VPN adapter ready: {} ({} addresses)",
            one.friendly_name,
            one.addresses.len()
        ),
        many => println!("Multiple VPN adapter candidates ({}) — narrow your config", many.len()),
    }
    Ok(())
}
