using System.Net;
using System.Net.NetworkInformation;

namespace HaloToolbox;

public sealed class NetworkStatsMonitor : IDisposable
{
    private readonly object _lock = new();
    private readonly Queue<PingSample> _samples = new();
    private readonly Queue<PingSample> _graphSamples = new();
    private readonly TimeSpan _window = TimeSpan.FromSeconds(10);
    private const int MaxGraphSamples = 36;
    private CancellationTokenSource? _cts;
    private Task? _loopTask;

    public event EventHandler<NetworkStatsSnapshot>? StatsUpdated;

    public string TargetIp { get; private set; } = "";

    public void Start(string targetIp)
    {
        if (!IPAddress.TryParse(targetIp, out _))
            return;

        lock (_lock)
        {
            if (_cts is not null && string.Equals(TargetIp, targetIp, StringComparison.OrdinalIgnoreCase))
                return;
        }

        Stop();

        TargetIp = targetIp;
        _cts = new CancellationTokenSource();
        _loopTask = Task.Run(() => PingLoopAsync(targetIp, _cts.Token));
    }

    public void Stop()
    {
        CancellationTokenSource? cts;
        Task? loopTask;

        lock (_lock)
        {
            _samples.Clear();
            _graphSamples.Clear();
            cts = _cts;
            loopTask = _loopTask;
            _cts = null;
            _loopTask = null;
            TargetIp = "";
        }

        if (cts is null)
            return;

        try
        {
            cts.Cancel();
        }
        catch
        {
            // Best effort; the ping loop exits on the next cancellation check.
        }

        if (loopTask is null)
            cts.Dispose();
        else
            _ = loopTask.ContinueWith(_ => cts.Dispose(), TaskScheduler.Default);
    }

    private async Task PingLoopAsync(string targetIp, CancellationToken token)
    {
        using var ping = new Ping();

        while (!token.IsCancellationRequested)
        {
            var sentAtUtc = DateTime.UtcNow;
            long? rttMs = null;

            try
            {
                var reply = await ping.SendPingAsync(targetIp, timeout: 1000);
                if (reply.Status == IPStatus.Success)
                    rttMs = reply.RoundtripTime;
            }
            catch
            {
                rttMs = null;
            }

            if (token.IsCancellationRequested)
                break;

            AddSample(new PingSample(sentAtUtc, rttMs));

            try
            {
                await Task.Delay(TimeSpan.FromSeconds(1), token);
            }
            catch (TaskCanceledException)
            {
                break;
            }
        }
    }

    private void AddSample(PingSample sample)
    {
        NetworkStatsSnapshot snapshot;

        lock (_lock)
        {
            _samples.Enqueue(sample);
            _graphSamples.Enqueue(sample);
            TrimSamples(sample.SentAtUtc);
            TrimGraphSamples();

            int sent = _samples.Count;
            int received = _samples.Count(x => x.RttMs.HasValue);
            long? latestRtt = sample.RttMs;
            double packetLoss = sent == 0 ? 0 : ((sent - received) * 100.0) / sent;
            var history = _graphSamples.Select(x => x.RttMs).ToArray();

            snapshot = new NetworkStatsSnapshot(TargetIp, latestRtt, packetLoss, sent, received, history);
        }

        StatsUpdated?.Invoke(this, snapshot);
    }

    private void TrimSamples(DateTime nowUtc)
    {
        while (_samples.Count > 0 && nowUtc - _samples.Peek().SentAtUtc > _window)
            _samples.Dequeue();
    }

    private void TrimGraphSamples()
    {
        while (_graphSamples.Count > MaxGraphSamples)
            _graphSamples.Dequeue();
    }

    public void Dispose()
    {
        Stop();
        _loopTask = null;
    }

    private sealed record PingSample(DateTime SentAtUtc, long? RttMs);
}

public sealed record NetworkStatsSnapshot(
    string TargetIp,
    long? RttMs,
    double PacketLossPercent,
    int SentCount,
    int ReceivedCount,
    IReadOnlyList<long?> RttHistory);
