using System.Collections.Concurrent;
using System.Diagnostics;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Runtime.InteropServices;

namespace HaloToolbox;

public sealed class GameServerConnectionMonitor : IDisposable
{
    private const int AfInet = 2;
    private const int SioRcvAll = unchecked((int)0x98000001);
    private static readonly TimeSpan ObservationWindow = TimeSpan.FromSeconds(10);
    private static readonly TimeSpan TrafficRateWindow = TimeSpan.FromSeconds(2);
    private static readonly TimeSpan CandidateStaleAfter = TimeSpan.FromSeconds(4);

    private readonly object _lock = new();
    private readonly Queue<UdpObservation> _observations = new();
    private readonly ConcurrentBag<Socket> _sockets = new();
    private readonly ConcurrentDictionary<string, byte> _localAddressStrings = new(StringComparer.OrdinalIgnoreCase);
    private CancellationTokenSource? _cts;
    private Task? _captureTask;
    private Task? _evaluateTask;
    private string _activeEndpoint = "";
    private string _lastStatus = "";

    public event EventHandler<GameServerInfo?>? ActiveServerChanged;
    public event EventHandler<NetworkTrafficSnapshot>? TrafficStatsUpdated;
    public event EventHandler<string>? StatusChanged;

    public bool IsRunning => _cts is not null;
    public bool HasActiveServer
    {
        get
        {
            lock (_lock)
                return !string.IsNullOrWhiteSpace(_activeEndpoint);
        }
    }

    public void Start()
    {
        lock (_lock)
        {
            if (_cts is not null)
                return;

            _observations.Clear();
            _activeEndpoint = "";
            _cts = new CancellationTokenSource();
            _captureTask = Task.Run(() => CaptureLoopAsync(_cts.Token));
            _evaluateTask = Task.Run(() => EvaluateLoopAsync(_cts.Token));
        }
    }

    public void Stop()
    {
        CancellationTokenSource? cts;
        Task? captureTask;
        Task? evaluateTask;

        lock (_lock)
        {
            cts = _cts;
            captureTask = _captureTask;
            evaluateTask = _evaluateTask;
            _cts = null;
            _captureTask = null;
            _evaluateTask = null;
            _observations.Clear();
            _activeEndpoint = "";
        }

        if (cts is null)
            return;

        try { cts.Cancel(); } catch { }

        while (_sockets.TryTake(out var socket))
        {
            try { socket.Close(); } catch { }
            try { socket.Dispose(); } catch { }
        }

        _ = Task.WhenAll(
                captureTask ?? Task.CompletedTask,
                evaluateTask ?? Task.CompletedTask)
            .ContinueWith(_ => cts.Dispose(), TaskScheduler.Default);
    }

    private async Task CaptureLoopAsync(CancellationToken token)
    {
        var localAddresses = GetCaptureAddresses();
        if (localAddresses.Count == 0)
        {
            PublishStatus("No active IPv4 interface found for packet observation.");
            return;
        }

        _localAddressStrings.Clear();
        foreach (var address in localAddresses)
            _localAddressStrings.TryAdd(address.ToString(), 0);

        var tasks = localAddresses.Select(address => Task.Run(() => CaptureAddressLoop(address, token), token)).ToArray();
        try
        {
            await Task.WhenAll(tasks);
        }
        catch (OperationCanceledException)
        {
        }
    }

    private void CaptureAddressLoop(IPAddress localAddress, CancellationToken token)
    {
        Socket? socket = null;

        try
        {
            socket = new Socket(AddressFamily.InterNetwork, SocketType.Raw, ProtocolType.IP);
            _sockets.Add(socket);
            socket.Bind(new IPEndPoint(localAddress, 0));
            socket.SetSocketOption(SocketOptionLevel.IP, SocketOptionName.HeaderIncluded, true);
            socket.ReceiveTimeout = 1000;
            socket.IOControl(SioRcvAll, BitConverter.GetBytes(1), new byte[4]);
            PublishStatus($"Watching MCC UDP traffic on {localAddress}.");
        }
        catch (Exception ex)
        {
            PublishStatus($"Packet observer could not start on {localAddress}: {ex.Message}");
            try { socket?.Dispose(); } catch { }
            return;
        }

        var buffer = new byte[65535];

        while (!token.IsCancellationRequested)
        {
            int length;
            try
            {
                length = socket.Receive(buffer);
            }
            catch (SocketException ex) when (ex.SocketErrorCode == SocketError.TimedOut)
            {
                continue;
            }
            catch
            {
                if (!token.IsCancellationRequested)
                    PublishStatus($"Packet observer stopped on {localAddress}.");
                break;
            }

            if (TryParseUdpPacket(buffer, length, out var packet))
                AddObservation(packet);
        }
    }

    private async Task EvaluateLoopAsync(CancellationToken token)
    {
        while (!token.IsCancellationRequested)
        {
            try
            {
                await Task.Delay(TimeSpan.FromSeconds(1), token);
            }
            catch (TaskCanceledException)
            {
                break;
            }

            var ports = GetMccUdpPorts();
            var candidate = EvaluateCandidate(ports);
            var regionLabel = candidate is null
                ? ""
                : GameServerRegionResolver.GetCachedOrQueueIpRegion(candidate.RemoteIp);
            var endpoint = candidate is null ? "" : $"{candidate.RemoteIp}:{candidate.RemotePort}";
            var activeKey = string.IsNullOrWhiteSpace(regionLabel) ? endpoint : $"{endpoint}|{regionLabel}";
            var activeChanged = false;

            lock (_lock)
            {
                if (!string.Equals(activeKey, _activeEndpoint, StringComparison.OrdinalIgnoreCase))
                {
                    _activeEndpoint = activeKey;
                    activeChanged = true;
                }
            }

            if (activeChanged)
                ActiveServerChanged?.Invoke(this, candidate?.ToGameServerInfo(regionLabel));

            if (candidate is not null)
                PublishTrafficStats(candidate, ports);
        }
    }

    private void AddObservation(UdpPacket packet)
    {
        if (!TryClassifyMccTraffic(packet, out var observation))
            return;

        var now = DateTime.UtcNow;
        lock (_lock)
        {
            _observations.Enqueue(observation with { SeenAtUtc = now });
            TrimObservations(now);
        }
    }

    private ServerCandidate? EvaluateCandidate(IReadOnlySet<int> mccUdpPorts)
    {
        if (mccUdpPorts.Count == 0)
            return null;

        var now = DateTime.UtcNow;
        List<UdpObservation> observations;

        lock (_lock)
        {
            TrimObservations(now);
            observations = _observations
                .Where(x => mccUdpPorts.Contains(x.LocalPort))
                .ToList();
        }

        var candidates = observations
            .Where(IsGameServerLike)
            .GroupBy(x => $"{x.RemoteIp}:{x.RemotePort}", StringComparer.OrdinalIgnoreCase)
            .Select(group =>
            {
                var items = group.ToList();
                var latest = items.Max(x => x.SeenAtUtc);
                int packets = items.Count;
                int bytes = items.Sum(x => x.PayloadBytes);
                int appPackets = items.Count(x => x.FirstPayloadByte is 0x17);
                int handshakePackets = items.Count(x => x.FirstPayloadByte is 0x16 or 0x14);
                bool inbound = items.Any(x => !x.Outbound);
                bool outbound = items.Any(x => x.Outbound);
                var first = items[0];
                double score = packets + (bytes / 900.0) + (appPackets * 5) + (handshakePackets * 3);

                if (first.RemotePort is >= 30000 and <= 32000)
                    score += 60;
                if (inbound && outbound)
                    score += 80;

                return new ServerCandidate(
                    first.RemoteIp,
                    first.RemotePort,
                    packets,
                    bytes,
                    latest,
                    score);
            })
            .Where(x => now - x.LastSeenUtc <= CandidateStaleAfter)
            .Where(x => x.Packets >= 10 || x.Bytes >= 2500)
            .OrderByDescending(x => x.Score)
            .FirstOrDefault();

        return candidates;
    }

    private void TrimObservations(DateTime nowUtc)
    {
        while (_observations.Count > 0 && nowUtc - _observations.Peek().SeenAtUtc > ObservationWindow)
            _observations.Dequeue();
    }

    private void PublishTrafficStats(ServerCandidate candidate, IReadOnlySet<int> mccUdpPorts)
    {
        var now = DateTime.UtcNow;
        List<UdpObservation> observations;

        lock (_lock)
        {
            TrimObservations(now);
            observations = _observations
                .Where(x =>
                    mccUdpPorts.Contains(x.LocalPort) &&
                    string.Equals(x.RemoteIp, candidate.RemoteIp, StringComparison.OrdinalIgnoreCase) &&
                    x.RemotePort == candidate.RemotePort &&
                    now - x.SeenAtUtc <= TrafficRateWindow)
                .ToList();
        }

        double seconds = TrafficRateWindow.TotalSeconds;
        int uploadPackets = observations.Count(x => x.Outbound);
        int downloadPackets = observations.Count(x => !x.Outbound);
        int uploadBytes = observations.Where(x => x.Outbound).Sum(x => x.PayloadBytes);
        int downloadBytes = observations.Where(x => !x.Outbound).Sum(x => x.PayloadBytes);

        TrafficStatsUpdated?.Invoke(this, new NetworkTrafficSnapshot(
            candidate.RemoteIp,
            candidate.RemotePort,
            uploadBytes / 1024.0 / seconds,
            downloadBytes / 1024.0 / seconds,
            uploadPackets / seconds,
            downloadPackets / seconds));
    }

    private bool TryClassifyMccTraffic(UdpPacket packet, out UdpObservation observation)
    {
        observation = default;

        bool sourceLocal = IsLocalAddress(packet.SourceIp);
        bool destinationLocal = IsLocalAddress(packet.DestinationIp);
        if (sourceLocal == destinationLocal)
            return false;

        string remoteIp = sourceLocal ? packet.DestinationIp : packet.SourceIp;
        int remotePort = sourceLocal ? packet.DestinationPort : packet.SourcePort;
        int localPort = sourceLocal ? packet.SourcePort : packet.DestinationPort;

        if (!IsPublicAddress(remoteIp))
            return false;

        observation = new UdpObservation(
            DateTime.UtcNow,
            remoteIp,
            remotePort,
            localPort,
            packet.PayloadBytes,
            packet.FirstPayloadByte,
            sourceLocal);
        return true;
    }

    private static bool IsGameServerLike(UdpObservation observation)
    {
        if (observation.RemotePort is 3075 or 3478)
            return false;

        if (observation.RemotePort is >= 30000 and <= 32000)
            return true;

        return observation.FirstPayloadByte is 0x16 or 0x17 or 0x14;
    }

    private static IReadOnlySet<int> GetMccUdpPorts()
    {
        var processIds = GetMccProcessIds();
        if (processIds.Count == 0)
            return new HashSet<int>();

        var ports = new HashSet<int>();
        int size = 0;
        _ = GetExtendedUdpTable(IntPtr.Zero, ref size, true, AfInet, UdpTableClass.OwnerPid, 0);

        var buffer = Marshal.AllocHGlobal(size);
        try
        {
            if (GetExtendedUdpTable(buffer, ref size, true, AfInet, UdpTableClass.OwnerPid, 0) != 0)
                return ports;

            int count = Marshal.ReadInt32(buffer);
            int rowSize = Marshal.SizeOf<MibUdpRowOwnerPid>();

            for (int i = 0; i < count; i++)
            {
                var row = Marshal.PtrToStructure<MibUdpRowOwnerPid>(IntPtr.Add(buffer, sizeof(int) + (i * rowSize)));
                if (processIds.Contains((int)row.OwningPid))
                    ports.Add(PortFromNetworkOrder(row.LocalPort));
            }
        }
        finally
        {
            Marshal.FreeHGlobal(buffer);
        }

        ports.Add(4379);
        return ports;
    }

    private static HashSet<int> GetMccProcessIds()
    {
        try
        {
            return Process.GetProcessesByName("MCC-Win64-Shipping")
                .Concat(Process.GetProcessesByName("MCC"))
                .Select(x => x.Id)
                .ToHashSet();
        }
        catch
        {
            return new HashSet<int>();
        }
    }

    private static int PortFromNetworkOrder(uint port) =>
        (int)(((port & 0xFF) << 8) | ((port >> 8) & 0xFF));

    private static List<IPAddress> GetCaptureAddresses()
    {
        return NetworkInterface.GetAllNetworkInterfaces()
            .Where(x => x.OperationalStatus == OperationalStatus.Up)
            .SelectMany(x => x.GetIPProperties().UnicastAddresses)
            .Where(x => x.Address.AddressFamily == AddressFamily.InterNetwork)
            .Select(x => x.Address)
            .Where(x => !IPAddress.IsLoopback(x))
            .Distinct()
            .ToList();
    }

    private static bool TryParseUdpPacket(byte[] buffer, int length, out UdpPacket packet)
    {
        packet = default;
        if (length < 28)
            return false;

        int version = buffer[0] >> 4;
        if (version != 4)
            return false;

        int headerLength = (buffer[0] & 0x0F) * 4;
        if (headerLength < 20 || length < headerLength + 8 || buffer[9] != 17)
            return false;

        string sourceIp = $"{buffer[12]}.{buffer[13]}.{buffer[14]}.{buffer[15]}";
        string destinationIp = $"{buffer[16]}.{buffer[17]}.{buffer[18]}.{buffer[19]}";
        int udpOffset = headerLength;
        int sourcePort = ReadUInt16BigEndian(buffer, udpOffset);
        int destinationPort = ReadUInt16BigEndian(buffer, udpOffset + 2);
        int udpLength = ReadUInt16BigEndian(buffer, udpOffset + 4);
        int payloadBytes = Math.Max(0, Math.Min(udpLength - 8, length - udpOffset - 8));
        byte? firstPayloadByte = payloadBytes > 0 ? buffer[udpOffset + 8] : null;

        packet = new UdpPacket(sourceIp, sourcePort, destinationIp, destinationPort, payloadBytes, firstPayloadByte);
        return true;
    }

    private static int ReadUInt16BigEndian(byte[] buffer, int offset) =>
        (buffer[offset] << 8) | buffer[offset + 1];

    private bool IsLocalAddress(string ip) => _localAddressStrings.ContainsKey(ip);

    private static bool IsPublicAddress(string ip)
    {
        if (!IPAddress.TryParse(ip, out var address))
            return false;

        var bytes = address.GetAddressBytes();
        if (bytes.Length != 4)
            return false;

        return bytes[0] switch
        {
            0 or 10 or 127 or >= 224 => false,
            100 when bytes[1] is >= 64 and <= 127 => false,
            169 when bytes[1] == 254 => false,
            172 when bytes[1] is >= 16 and <= 31 => false,
            192 when bytes[1] == 168 => false,
            _ => true
        };
    }

    private void PublishStatus(string status)
    {
        if (string.Equals(status, _lastStatus, StringComparison.Ordinal))
            return;

        _lastStatus = status;
        StatusChanged?.Invoke(this, status);
    }

    public void Dispose()
    {
        Stop();
    }

    private readonly record struct UdpPacket(
        string SourceIp,
        int SourcePort,
        string DestinationIp,
        int DestinationPort,
        int PayloadBytes,
        byte? FirstPayloadByte);

    private readonly record struct UdpObservation(
        DateTime SeenAtUtc,
        string RemoteIp,
        int RemotePort,
        int LocalPort,
        int PayloadBytes,
        byte? FirstPayloadByte,
        bool Outbound);

    private sealed record ServerCandidate(
        string RemoteIp,
        int RemotePort,
        int Packets,
        int Bytes,
        DateTime LastSeenUtc,
        double Score)
    {
        public GameServerInfo ToGameServerInfo(string regionLabel) => new()
        {
            IPv4Address = RemoteIp,
            Region = regionLabel,
            State = "Observed",
            CachedAt = DateTime.UtcNow,
            Ports =
            [
                new GameServerPort
                {
                    Name = "udp",
                    Num = RemotePort,
                    Protocol = "UDP"
                }
            ]
        };
    }

    private enum UdpTableClass
    {
        Basic = 0,
        OwnerPid = 1,
        OwnerModule = 2
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MibUdpRowOwnerPid
    {
        public uint LocalAddr;
        public uint LocalPort;
        public uint OwningPid;
    }

    [DllImport("iphlpapi.dll", SetLastError = true)]
    private static extern uint GetExtendedUdpTable(
        IntPtr pUdpTable,
        ref int dwOutBufLen,
        bool sort,
        int ipVersion,
        UdpTableClass tblClass,
        uint reserved);
}

public sealed record NetworkTrafficSnapshot(
    string RemoteIp,
    int RemotePort,
    double UploadKilobytesPerSecond,
    double DownloadKilobytesPerSecond,
    double UploadPacketsPerSecond,
    double DownloadPacketsPerSecond);
