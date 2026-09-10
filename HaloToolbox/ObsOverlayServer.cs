using System.Globalization;
using System.Net;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace HaloToolbox;

internal sealed class ObsOverlayServer : IDisposable
{
    private const int Port = 19998;
    private const string OverlayVersion = "isolated-game-overlay-v6";
    private readonly object _sync = new();
    private readonly JsonSerializerOptions _jsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    private HttpListener? _listener;
    private CancellationTokenSource? _cts;
    private Task? _listenTask;
    private ObsOverlaySnapshot _snapshot = ObsOverlaySnapshot.Empty;

    public string Url => $"http://127.0.0.1:{Port}/overlay/?mode=obs&v={OverlayVersion}";
    public string GameOverlayUrl => $"http://127.0.0.1:{Port}/overlay/?mode=game&v={OverlayVersion}";
    public bool IsRunning => _listener?.IsListening == true;

    public void Start()
    {
        if (IsRunning)
            return;

        _cts = new CancellationTokenSource();
        _listener = new HttpListener();
        _listener.Prefixes.Add($"http://127.0.0.1:{Port}/");
        _listener.Start();
        _listenTask = Task.Run(() => ListenLoopAsync(_cts.Token));
    }

    public void Stop()
    {
        try { _cts?.Cancel(); } catch { }
        try { _listener?.Stop(); } catch { }
        try { _listener?.Close(); } catch { }
        _listener = null;
        _cts?.Dispose();
        _cts = null;
        _listenTask = null;
    }

    public void Update(ObsOverlaySnapshot snapshot)
    {
        lock (_sync)
        {
            _snapshot = snapshot;
        }
    }

    public void Dispose() => Stop();

    private async Task ListenLoopAsync(CancellationToken token)
    {
        while (!token.IsCancellationRequested && _listener is { IsListening: true } listener)
        {
            HttpListenerContext context;
            try
            {
                context = await listener.GetContextAsync();
            }
            catch
            {
                if (token.IsCancellationRequested)
                    return;
                continue;
            }

            _ = Task.Run(() => HandleRequestAsync(context), token);
        }
    }

    private async Task HandleRequestAsync(HttpListenerContext context)
    {
        try
        {
            string path = context.Request.Url?.AbsolutePath.TrimEnd('/') ?? "";
            if (path.Length == 0 || path.Equals("/overlay", StringComparison.OrdinalIgnoreCase))
            {
                await WriteTextAsync(context, GetOverlayHtml(), "text/html; charset=utf-8");
                return;
            }

            if (path.Equals("/state.json", StringComparison.OrdinalIgnoreCase))
            {
                ObsOverlaySnapshot snapshot;
                lock (_sync)
                {
                    snapshot = _snapshot;
                }

                await WriteTextAsync(
                    context,
                    JsonSerializer.Serialize(snapshot, _jsonOptions),
                    "application/json; charset=utf-8");
                return;
            }

            context.Response.StatusCode = 404;
            await WriteTextAsync(context, "Not found", "text/plain; charset=utf-8");
        }
        catch
        {
            try { context.Response.StatusCode = 500; } catch { }
        }
        finally
        {
            try { context.Response.Close(); } catch { }
        }
    }

    private static async Task WriteTextAsync(HttpListenerContext context, string text, string contentType)
    {
        byte[] bytes = Encoding.UTF8.GetBytes(text);
        context.Response.ContentType = contentType;
        context.Response.ContentLength64 = bytes.Length;
        context.Response.Headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0";
        context.Response.Headers["Pragma"] = "no-cache";
        context.Response.Headers["Expires"] = "0";
        await context.Response.OutputStream.WriteAsync(bytes);
    }

    private static string GetOverlayHtml() => """
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Halo Toolbox OBS Overlay</title>
  <style>
    :root {
      color-scheme: dark;
      font-family: Consolas, "Cascadia Mono", monospace;
      background: transparent;
      --cyan: #00c8ff;
      --green: #39ff14;
      --red: #ff2d55;
      --muted: #8fb8d8;
      --text: #c8d8e8;
      --shadow: 0 2px 5px rgba(0, 0, 0, 0.95), 0 0 10px rgba(0, 0, 0, 0.75);
    }

    * { box-sizing: border-box; }
    html, body { margin: 0; width: 100%; min-height: 100%; overflow: hidden; background: transparent; }
    body { padding: 8px; }
    .overlay {
      width: 760px;
      padding: 0;
      background: transparent;
      transform-origin: top left;
    }
    body.obs-mode .overlay { position: fixed; left: 0; top: 0; }
    .inner { display: flex; align-items: flex-start; gap: 24px; padding: 10px; background: transparent; }
    .network { width: 392px; flex: 0 0 392px; }
    .session { width: 300px; flex: 0 0 300px; padding-top: 1px; }
    .row { display: flex; align-items: baseline; justify-content: space-between; gap: 12px; }
    .label { color: var(--muted); font-size: 12px; font-weight: 700; text-shadow: var(--shadow); }
    .section-title { color: var(--cyan); font-size: 13px; font-weight: 700; text-shadow: var(--shadow); }
    .server { color: var(--muted); font-size: 13px; text-align: right; text-shadow: var(--shadow); }
    .ping { color: var(--green); font-size: 26px; font-weight: 700; line-height: 1.25; text-shadow: var(--shadow); }
    .loss { color: var(--text); font-size: 21px; font-weight: 700; text-shadow: var(--shadow); }
    .traffic { margin-top: 8px; color: var(--text); font-size: 13px; font-weight: 700; text-shadow: var(--shadow); }
    .traffic .up { color: var(--green); }
    .traffic .down { color: var(--cyan); }
    .packets { margin-top: 4px; color: var(--muted); font-size: 12px; text-shadow: var(--shadow); }
    .graph { width: 100%; height: 34px; margin-top: 8px; }
    .divider { height: 1px; margin: 10px 0; background: rgba(0, 200, 255, 0.32); box-shadow: 0 1px 4px rgba(0, 0, 0, 0.95); }
    .session.hidden { display: none; }
    .record { margin-top: 8px; }
    .wins, .losses, .sep { font-size: 30px; font-weight: 700; text-shadow: var(--shadow); }
    .wins { color: var(--green); }
    .losses { color: var(--red); }
    .sep { color: #53636e; padding: 0 6px; }
    .winrate-label { color: var(--muted); font-size: 13px; text-align: right; text-shadow: var(--shadow); }
    .winrate { color: var(--cyan); font-size: 26px; font-weight: 700; text-align: right; text-shadow: var(--shadow); }
    .stat-box {
      margin-top: 8px;
      padding: 3px 0;
      background: transparent;
    }
    .kd-value { color: var(--text); font-size: 20px; font-weight: 700; margin-left: 10px; text-shadow: var(--shadow); }
    .kills { color: var(--green); font-size: 16px; font-weight: 700; text-shadow: var(--shadow); }
    .deaths { color: var(--red); font-size: 16px; font-weight: 700; text-shadow: var(--shadow); }
  </style>
</head>
<body>
  <main class="overlay">
    <div class="inner">
      <div class="network">
        <div class="row">
          <div class="section-title">NETWORK</div>
          <div id="server" class="server">SERVER: --</div>
        </div>
        <div class="row">
          <div id="ping" class="ping">Ping: -- ms</div>
          <div id="loss" class="loss">Loss: --%</div>
        </div>
        <div class="row traffic">
          <div><span class="up">UP</span> <span id="up">--</span></div>
          <div><span class="down">DOWN</span> <span id="down">--</span></div>
        </div>
        <div class="row packets">
          <div id="upPackets">-- packets/s</div>
          <div id="downPackets">-- packets/s</div>
        </div>
        <svg class="graph" viewBox="0 0 352 34" preserveAspectRatio="none">
          <line x1="0" y1="10" x2="352" y2="10" stroke="rgba(0,200,255,.14)" />
          <line x1="0" y1="24" x2="352" y2="24" stroke="rgba(0,200,255,.14)" />
          <polyline id="graphGlow" points="" fill="none" stroke="rgba(0,200,255,.32)" stroke-width="6" stroke-linejoin="round" />
          <polyline id="graphLine" points="" fill="none" stroke="#39ff14" stroke-width="2.4" stroke-linejoin="round" />
        </svg>
      </div>

      <section id="session" class="session">
        <div class="row">
          <div class="label">SESSION</div>
          <div id="games" class="label">0 games</div>
        </div>
        <div class="row record">
          <div><span id="wins" class="wins">0W</span><span class="sep">-</span><span id="losses" class="losses">0L</span></div>
          <div><div class="winrate-label">WIN RATE</div><div id="winRate" class="winrate">--</div></div>
        </div>
        <div class="stat-box row">
          <div><span class="label">K/D</span><span id="kd" class="kd-value">--</span></div>
          <div id="kills" class="kills">0 Kills</div>
          <div id="deaths" class="deaths">0 Deaths</div>
        </div>
      </section>
    </div>
  </main>

  <script>
    const $ = id => document.getElementById(id);
    const set = (id, value) => { $(id).textContent = value || "--"; };
    const setBlankable = (id, value) => { $(id).textContent = value || ""; };
    const params = new URLSearchParams(window.location.search);
    const isObsMode = params.get("mode") === "obs";
    const urlScale = Math.max(0.25, Math.min(4, Number(params.get("scale")) || 1));
    document.body.classList.toggle("obs-mode", isObsMode);
    const BASE_WIDTH = 760;
    const BASE_HEIGHT = 112;

    function formatKb(value) {
      if (value === null || value === undefined) return "--";
      if (value >= 1024) return (value / 1024).toFixed(1) + " MB/s";
      return value.toFixed(value >= 10 ? 0 : 1) + " KB/s";
    }

    function graphPoints(history) {
      if (!history || history.length < 2) return "";
      const max = Math.max(80, ...history.filter(x => x !== null).map(Number));
      const step = 352 / Math.max(1, history.length - 1);
      return history.map((v, i) => {
        const value = v === null ? max : Number(v);
        const y = 30 - Math.max(0, Math.min(1, value / max)) * 26;
        return `${(i * step).toFixed(1)},${y.toFixed(1)}`;
      }).join(" ");
    }

    function applyPlacement(data) {
      const overlay = document.querySelector(".overlay");
      if (!overlay) return;

      if (isObsMode) {
        const left = (data.overlayLeftRatio ?? 0) * window.innerWidth;
        const top = (data.overlayTopRatio ?? 0) * window.innerHeight;
        const width = Math.max(1, (data.overlayWidthRatio ?? 0) * window.innerWidth);
        const scale = (width > 1 ? width / BASE_WIDTH : 1) * urlScale;
        overlay.style.left = `${left}px`;
        overlay.style.top = `${top}px`;
        overlay.style.transform = `scale(${scale})`;
        return;
      }

      const scale = Math.min(
        window.innerWidth / BASE_WIDTH,
        window.innerHeight / BASE_HEIGHT
      );
      overlay.style.transform = `scale(${Math.max(0.1, scale * urlScale)})`;
    }

    async function refresh() {
      try {
        const data = await fetch("/state.json", { cache: "no-store" }).then(r => r.json());
        applyPlacement(data);
        const hasPing = data.rttMs !== null && data.rttMs !== undefined;
        const hasNetwork = hasPing || data.uploadKilobytesPerSecond !== null || data.downloadKilobytesPerSecond !== null;
        setBlankable("server", hasNetwork && data.serverLabel !== "SERVER: --" ? data.serverLabel : "");
        setBlankable("ping", hasPing ? `Ping: ${data.rttMs} ms` : "");
        setBlankable("loss", hasNetwork ? `Loss: ${(data.packetLossPercent ?? 0).toFixed(0)}%` : "");
        setBlankable("up", data.uploadKilobytesPerSecond !== null && data.uploadKilobytesPerSecond !== undefined ? formatKb(data.uploadKilobytesPerSecond) : "");
        setBlankable("down", data.downloadKilobytesPerSecond !== null && data.downloadKilobytesPerSecond !== undefined ? formatKb(data.downloadKilobytesPerSecond) : "");
        setBlankable("upPackets", data.uploadPacketsPerSecond !== null && data.uploadPacketsPerSecond !== undefined ? `${data.uploadPacketsPerSecond.toFixed(0)} packets/s` : "");
        setBlankable("downPackets", data.downloadPacketsPerSecond !== null && data.downloadPacketsPerSecond !== undefined ? `${data.downloadPacketsPerSecond.toFixed(0)} packets/s` : "");
        const points = graphPoints(data.rttHistoryMs);
        $("graphLine").setAttribute("points", points);
        $("graphGlow").setAttribute("points", points);

        $("session").classList.toggle("hidden", !data.showSessionStats);
        const games = data.gamesPlayed ?? 0;
        set("wins", `${data.wins ?? 0}W`);
        setBlankable("losses", games > 0 || (data.losses ?? 0) > 0 ? `${data.losses ?? 0}L` : "");
        setBlankable("games", games > 0 ? `${games} game${games === 1 ? "" : "s"}` : "");
        setBlankable("winRate", games > 0 ? `${Math.round(((data.wins ?? 0) / games) * 100)}%` : "--");
        set("kd", data.sessionKd || "--");
        setBlankable("kills", (data.kills ?? 0) > 0 ? `${data.kills} Kills` : "");
        setBlankable("deaths", (data.deaths ?? 0) > 0 ? `${data.deaths} Deaths` : "");
      } catch {
      }
    }

    refresh();
    setInterval(refresh, 500);
  </script>
</body>
</html>
""";
}

internal sealed record ObsOverlaySnapshot(
    bool ShowSessionStats,
    string ServerLabel,
    int? RttMs,
    double PacketLossPercent,
    IReadOnlyList<int?> RttHistoryMs,
    double? UploadKilobytesPerSecond,
    double? DownloadKilobytesPerSecond,
    double? UploadPacketsPerSecond,
    double? DownloadPacketsPerSecond,
    int Wins,
    int Losses,
    int GamesPlayed,
    long Kills,
    long Deaths,
    string SessionKd,
    double OverlayLeftRatio,
    double OverlayTopRatio,
    double OverlayWidthRatio,
    double OverlayHeightRatio)
{
    public static ObsOverlaySnapshot Empty { get; } = new(
        ShowSessionStats: true,
        ServerLabel: "SERVER: --",
        RttMs: null,
        PacketLossPercent: 0,
        RttHistoryMs: Array.Empty<int?>(),
        UploadKilobytesPerSecond: null,
        DownloadKilobytesPerSecond: null,
        UploadPacketsPerSecond: null,
        DownloadPacketsPerSecond: null,
        Wins: 0,
        Losses: 0,
        GamesPlayed: 0,
        Kills: 0,
        Deaths: 0,
        SessionKd: "--",
        OverlayLeftRatio: 0,
        OverlayTopRatio: 0,
        OverlayWidthRatio: 760.0 / 1920.0,
        OverlayHeightRatio: 132.0 / 1080.0);
}
