# Halo-only VPN

The VPN tab is an optional NordVPN/OpenVPN client built specifically for Halo MCC. It does not use or automate the NordVPN desktop application.

## User flow

1. Open **VPN** and click **Connect MCC only**.
2. The first connection downloads OpenVPN Community 2.7.6 from the official OpenVPN release host, verifies the pinned SHA-256, and installs only the OpenVPN engine, services, OpenSSL runtime, and signed network drivers.
3. Enter NordVPN **service credentials** from Nord Account → Manual setup.
4. Choose the Halo MCC server region you want: East US, East US 2, South Central US, North Central US, West US, West Europe, and the other MCC/Azure regions. The Toolbox tests direct TCP handshake latency to nearby Nord exits and selects the fastest responsive server. If a latency test is unavailable, it falls back to distance from the matching MCC datacenter and then load. The initial UDP profile comes from Nord's configuration CDN; live switches can prepare the replacement from the validated Nord profile already on disk.
5. Approve the Windows administrator prompt. The Toolbox relaunches, establishes the tunnel, verifies initialization, starts the Advanced Features service bundle with core Rejoin Recovery, and then optionally launches MCC.

Rejoin Recovery is not a separate VPN option. It is a core component that remains active whenever Advanced Features are running and stops when those services stop. If the VPN connects but Advanced Features cannot start, Toolbox leaves the VPN connected, reports the service problem, and does not auto-launch MCC.

To change regions while connected, select another MCC region. The nearby exits are retested, and the connection button becomes **Switch to &lt;region&gt;**. Selecting a different measured server in the current region changes it to **Switch server**. The Toolbox downloads the replacement profile before dropping the old tunnel, keeps the process router active and fail-closed during the handoff, then updates Network Stats when the new tunnel is ready. The Toolbox application does not restart.

The displayed latency is the direct TCP connection time from the player's PC to the Nord endpoint. The probe runs outside the MCC-only router, including while switching regions, so an existing tunnel does not distort the ranking. It is useful for comparing Nord servers in the selected area, but it is not a promise of in-game ping because the Nord-to-Azure hop and Internet peering still affect the final route.

Normal applications keep their existing default route. OpenVPN creates a high-metric route on its own interface and automatically removes it on disconnect; only sockets explicitly pinned to that interface use it. `HaloVpnRouter` tunnels only the configured MCC executables. Toolbox control-plane requests, browsers, Discord, Steam, and other applications continue to use the normal Windows route.

While connected, the in-game and OBS Network Stats overlays show the selected MCC-facing route as `VPN: <region>`. The line is removed immediately when the tunnel disconnects or fails.

## Safety behavior

- The router handles outbound IPv4 TCP and UDP.
- IPv6 from selected processes is blocked rather than allowed to leak outside the tunnel.
- Selected traffic fails closed while the VPN adapter is unavailable.
- The router uses WinDivert and a user-space TCP/UDP bridge; it does not inject code into MCC.
- OpenVPN and the router are attached to a Windows Job Object and terminate when the Toolbox closes or crashes.
- Nord service passwords are encrypted with Windows DPAPI for the current user. The temporary OpenVPN authentication file is restricted to that user and deleted immediately after the connection attempt.
- The NordVPN desktop application must be uninstalled because Nord currently documents that it prevents third-party OpenVPN tunnel adapters from initializing.

## Building the router

Install the minimal stable Rust toolchain and run:

```powershell
.\tools\build-vpn-router.ps1
```

The script builds `HaloVpnRouter` and refreshes the native resources embedded in `HaloMCCToolbox.exe`. Normal .NET builds use the checked-in, reviewed native bundle.

## Third-party components

- The routing core was adapted from ENA's MIT-licensed OpenVPN Split Tunneling Client. Its source and MIT notice are under `HaloVpnRouter/`.
- WinDivert is distributed under LGPLv3 (or GPLv2 at the recipient's option). Its unmodified DLL, driver, and complete license are included in `HaloToolbox/Native/Vpn/` and extracted beside the router at runtime so they can be replaced with compatible modified versions.
- OpenVPN Community is GPLv2. The Toolbox does not embed or redistribute OpenVPN; it downloads the official MSI and installs OpenVPN as an independent Windows product.

This feature requires Windows 10/11 x64 and administrator approval for packet routing.
