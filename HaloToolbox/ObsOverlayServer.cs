using System.Globalization;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Resources;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace HaloToolbox;

internal sealed class ObsOverlayServer : IDisposable
{
    private const int Port = 19998;
    private const string OverlayVersion = "component-placement-v19";
    private static readonly ResourceManager AppResources =
        new("HaloMCCToolbox.g", typeof(ObsOverlayServer).Assembly);
    private readonly object _sync = new();
    private readonly JsonSerializerOptions _jsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    private TcpListener? _listener;
    private CancellationTokenSource? _cts;
    private Task? _listenTask;
    private ObsOverlaySnapshot _snapshot = ObsOverlaySnapshot.Empty;

    // Keep browser-source URLs stable. OverlayVersion is an implementation
    // detail exposed by /health.json, not part of a user's permanent OBS link.
    // Dynamic overlay responses already use no-cache headers.
    public string Url => $"http://127.0.0.1:{Port}/overlay/?mode=obs";
    public string GameOverlayUrl => $"http://127.0.0.1:{Port}/overlay/?mode=game";
    public string ComponentUrl(string component, string mode = "game") =>
        $"http://127.0.0.1:{Port}/overlay/?mode={mode}&component={component}";
    public bool IsRunning => _listener is not null && _cts is { IsCancellationRequested: false };

    public void Start()
    {
        if (IsRunning)
            return;

        var cts = new CancellationTokenSource();
        var listener = new TcpListener(IPAddress.Loopback, Port);
        try
        {
            listener.Start();
            _cts = cts;
            _listener = listener;
            _listenTask = Task.Run(() => ListenLoopAsync(listener, cts.Token));
        }
        catch
        {
            listener.Stop();
            cts.Dispose();
            throw;
        }
    }

    public void Stop()
    {
        try { _cts?.Cancel(); } catch { }
        try { _listener?.Stop(); } catch { }
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

    private async Task ListenLoopAsync(TcpListener listener, CancellationToken token)
    {
        while (!token.IsCancellationRequested)
        {
            TcpClient client;
            try
            {
                client = await listener.AcceptTcpClientAsync(token);
            }
            catch (OperationCanceledException)
            {
                return;
            }
            catch (SocketException) when (token.IsCancellationRequested)
            {
                return;
            }
            catch when (!token.IsCancellationRequested)
            {
                continue;
            }

            _ = HandleRequestAsync(client, token);
        }
    }

    private async Task HandleRequestAsync(TcpClient client, CancellationToken token)
    {
        using (client)
        {
            try
            {
                client.NoDelay = true;
                await using NetworkStream stream = client.GetStream();
                using var reader = new StreamReader(
                    stream,
                    Encoding.ASCII,
                    detectEncodingFromByteOrderMarks: false,
                    bufferSize: 4096,
                    leaveOpen: true);

                string? requestLine = await reader.ReadLineAsync(token);
                if (string.IsNullOrWhiteSpace(requestLine))
                    return;

                for (int i = 0; i < 100; i++)
                {
                    string? header = await reader.ReadLineAsync(token);
                    if (string.IsNullOrEmpty(header))
                        break;
                }

                string[] requestParts = requestLine.Split(' ', 3, StringSplitOptions.RemoveEmptyEntries);
                if (requestParts.Length < 2)
                {
                    await WriteResponseAsync(stream, 400, "Bad Request", "text/plain; charset=utf-8", Encoding.UTF8.GetBytes("Bad request"), false, token);
                    return;
                }

                bool headOnly = requestParts[0].Equals("HEAD", StringComparison.OrdinalIgnoreCase);
                if (!headOnly && !requestParts[0].Equals("GET", StringComparison.OrdinalIgnoreCase))
                {
                    await WriteResponseAsync(stream, 405, "Method Not Allowed", "text/plain; charset=utf-8", Encoding.UTF8.GetBytes("Method not allowed"), false, token, headOnly);
                    return;
                }

                Uri requestUri = Uri.TryCreate(requestParts[1], UriKind.Absolute, out var absoluteUri)
                    ? absoluteUri
                    : new Uri(new Uri($"http://127.0.0.1:{Port}"), requestParts[1]);
                string path = requestUri.AbsolutePath.TrimEnd('/');

                if (path.Length == 0 || path.Equals("/overlay", StringComparison.OrdinalIgnoreCase))
                {
                    await WriteTextAsync(stream, GetOverlayHtml(), "text/html; charset=utf-8", token, headOnly);
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
                        stream,
                        JsonSerializer.Serialize(snapshot, _jsonOptions),
                        "application/json; charset=utf-8",
                        token,
                        headOnly);
                    return;
                }

                if (path.Equals("/health.json", StringComparison.OrdinalIgnoreCase))
                {
                    await WriteTextAsync(
                        stream,
                        $"{{\"status\":\"ok\",\"version\":\"{OverlayVersion}\"}}",
                        "application/json; charset=utf-8",
                        token,
                        headOnly);
                    return;
                }

                if (path.StartsWith("/medals/", StringComparison.OrdinalIgnoreCase))
                {
                    string fileName = Path.GetFileName(path);
                    var allowed = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                    {
                        "double-kill.png", "triple-kill.png", "overkill.png",
                        "killtacular.png", "killtrocity.png", "killimanjaro.png",
                        "killtastrophe.png", "killpocalypse.png", "killionaire.png"
                    };
                    if (allowed.Contains(fileName))
                    {
                        if (AppResources.GetObject(
                                $"resources/medals/{fileName}",
                                CultureInfo.InvariantCulture) is Stream resource)
                        {
                            using (resource)
                            {
                                await using var memory = new MemoryStream();
                                await resource.CopyToAsync(memory, token);
                                await WriteResponseAsync(stream, 200, "OK", "image/png", memory.ToArray(), true, token, headOnly);
                                return;
                            }
                        }
                    }
                }

                await WriteResponseAsync(stream, 404, "Not Found", "text/plain; charset=utf-8", Encoding.UTF8.GetBytes("Not found"), false, token, headOnly);
            }
            catch (Exception ex)
            {
                // OBS and WebView browser sources poll continuously. A client
                // disappearing between polls must not stop the listener.
                System.Diagnostics.Trace.WriteLine($"OBS overlay request failed: {ex}");
            }
        }
    }

    private static Task WriteTextAsync(
        NetworkStream stream,
        string text,
        string contentType,
        CancellationToken token,
        bool headOnly = false)
    {
        return WriteResponseAsync(
            stream,
            200,
            "OK",
            contentType,
            Encoding.UTF8.GetBytes(text),
            cacheForDay: false,
            token,
            headOnly);
    }

    private static async Task WriteResponseAsync(
        NetworkStream stream,
        int statusCode,
        string statusText,
        string contentType,
        byte[] body,
        bool cacheForDay,
        CancellationToken token,
        bool headOnly = false)
    {
        string cacheHeaders = cacheForDay
            ? "Cache-Control: public, max-age=86400\r\n"
            : "Cache-Control: no-store, no-cache, must-revalidate, max-age=0\r\nPragma: no-cache\r\nExpires: 0\r\n";
        string headers =
            $"HTTP/1.1 {statusCode} {statusText}\r\n" +
            $"Content-Type: {contentType}\r\n" +
            $"Content-Length: {body.Length.ToString(CultureInfo.InvariantCulture)}\r\n" +
            cacheHeaders +
            "Access-Control-Allow-Origin: *\r\n" +
            "X-Content-Type-Options: nosniff\r\n" +
            "Connection: close\r\n\r\n";
        byte[] headerBytes = Encoding.ASCII.GetBytes(headers);
        await stream.WriteAsync(headerBytes, token);
        if (!headOnly && body.Length > 0)
            await stream.WriteAsync(body, token);
        await stream.FlushAsync(token);
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
    body.component-mode { padding: 0; }
    .overlay {
      width: 1280px;
      padding: 0;
      background: transparent;
      transform-origin: top left;
    }
    body.obs-mode .overlay { position: fixed; left: 0; top: 0; }
    .inner { display: flex; align-items: flex-start; gap: 24px; padding: 10px; background: transparent; }
    .network { width: 392px; flex: 0 0 392px; }
    .network.hidden { display: none; }
    .matchmaking { width: 300px; flex: 0 0 300px; padding-top: 1px; }
    .matchmaking.hidden { display: none; }
    .wait-value { color: var(--cyan); font-size: 25px; font-weight: 700; margin-top: 7px; text-shadow: var(--shadow); }
    .wait-value.warn { color: #ffb020; }
    .wait-value.long { color: var(--red); }
    .wait-detail { color: var(--muted); font-size: 12px; margin-top: 5px; text-shadow: var(--shadow); }
    .session { width: 840px; flex: 0 0 840px; padding-top: 1px; }
    .row { display: flex; align-items: baseline; justify-content: space-between; gap: 12px; }
    .label { color: var(--muted); font-size: 12px; font-weight: 700; text-shadow: var(--shadow); }
    .section-title { color: var(--cyan); font-size: 13px; font-weight: 700; text-shadow: var(--shadow); }
    .server { color: var(--muted); font-size: 13px; text-align: right; text-shadow: var(--shadow); }
    .ping { color: var(--green); font-size: 26px; font-weight: 700; line-height: 1.25; text-shadow: var(--shadow); }
    .quality { display: flex; flex-direction: column; align-items: flex-end; line-height: 1.15; }
    .jitter { color: var(--green); font-size: 12px; font-weight: 700; text-align: right; text-shadow: var(--shadow); }
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
    .medals { display: grid; grid-template-columns: repeat(9, 1fr); gap: 7px; margin-top: 8px; }
    .medal { text-align: center; min-width: 0; opacity: .34; }
    .medal.earned { opacity: 1; }
    .medal img { width: 40px; height: 40px; object-fit: contain; display: block; margin: 0 auto 2px; }
    .medal-name { color: var(--muted); font-size: 8px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; text-shadow: var(--shadow); }
    .medal-count { color: var(--cyan); font-size: 14px; font-weight: 700; text-shadow: var(--shadow); }
    .recap { width: 840px; flex: 0 0 840px; color: var(--text); animation: recapIn .32s ease-out both; }
    .recap.hidden { display: none; }
    .recap-head { display: flex; align-items: center; justify-content: space-between; margin-bottom: 6px; }
    .recap-title { color: var(--cyan); font-size: 14px; font-weight: 700; letter-spacing: 1.5px; text-shadow: var(--shadow); }
    .recap-result { color: var(--green); font-size: 22px; font-weight: 700; text-shadow: var(--shadow); }
    .recap-result.loss { color: var(--red); }
    .recap-grid { display: grid; grid-template-columns: 1fr 1.05fr 1.55fr; gap: 14px; align-items: center; }
    .recap-block { min-width: 0; border-left: 1px solid rgba(0,200,255,.34); padding-left: 12px; }
    .recap-game { font-size: 20px; font-weight: 700; text-shadow: var(--shadow); }
    .recap-game span { color: var(--muted); font-size: 10px; font-weight: 400; }
    .recap-kd-change { color: var(--muted); font-size: 10px; margin-top: 7px; }
    .recap-kd-change strong { color: var(--cyan); font-size: 15px; }
    .recap-best { color: var(--green); font-size: 9px; margin-top: 3px; }
    .recap-feature { display: flex; align-items: center; gap: 10px; }
    .recap-feature img { width: 66px; height: 66px; object-fit: contain; }
    .recap-feature-name { color: var(--cyan); font-size: 15px; font-weight: 700; text-shadow: var(--shadow); }
    .recap-feature-detail { color: var(--muted); font-size: 9px; margin-top: 3px; }
    .recap-deltas { display: grid; grid-template-columns: 1fr 1fr; gap: 5px 10px; }
    .recap-delta { display: grid; grid-template-columns: 28px 1fr auto; align-items: center; gap: 6px; }
    .recap-delta img { width: 28px; height: 28px; object-fit: contain; }
    .recap-delta-name { color: var(--muted); font-size: 8px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .recap-delta-count { color: var(--green); font-size: 12px; font-weight: 700; }
    .recap-progress { height: 2px; margin-top: 8px; background: rgba(0,200,255,.18); overflow: hidden; }
    .recap-progress > div { height: 100%; background: var(--cyan); transform-origin: left; }
    @keyframes recapIn { from { opacity: 0; transform: translateY(8px); } to { opacity: 1; transform: translateY(0); } }
  </style>
</head>
<body>
  <main class="overlay">
    <div class="inner">
      <div id="network" class="network">
        <div class="row">
          <div class="section-title">NETWORK</div>
          <div id="server" class="server">SERVER: --</div>
        </div>
        <div class="row">
          <div id="ping" class="ping">Ping: -- ms</div>
          <div class="quality">
            <div id="loss" class="loss">Loss: --%</div>
            <div id="jitter" class="jitter">Jitter: -- ms</div>
          </div>
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

      <section id="matchmaking" class="matchmaking hidden">
        <div class="row">
          <div class="section-title">MATCHMAKING</div>
          <div class="label">ESTIMATE</div>
        </div>
        <div id="waitValue" class="wait-value">EST. WAIT --</div>
        <div id="waitDetail" class="wait-detail"></div>
      </section>

      <section id="session" class="session">
        <div class="row">
          <div class="label">SESSION</div>
          <div id="games" class="label">0 games</div>
        </div>
        <div class="row record">
          <div><span id="wins" class="wins">0W</span><span class="sep">-</span><span id="losses" class="losses">0L</span></div>
          <div><div class="winrate-label">WIN RATE</div><div id="winRate" class="winrate">0%</div></div>
        </div>
        <div class="stat-box row">
          <div><span class="label">K/D</span><span id="kd" class="kd-value">0.00</span></div>
          <div id="kills" class="kills">0 Kills</div>
          <div id="deaths" class="deaths">0 Deaths</div>
        </div>
        <div id="medals" class="medals">
          <div class="medal" data-field="doubleKills"><img src="/medals/double-kill.png"><div class="medal-name">DOUBLE</div><div class="medal-count">×0</div></div>
          <div class="medal" data-field="tripleKills"><img src="/medals/triple-kill.png"><div class="medal-name">TRIPLE</div><div class="medal-count">×0</div></div>
          <div class="medal" data-field="overkills"><img src="/medals/overkill.png"><div class="medal-name">OVERKILL</div><div class="medal-count">×0</div></div>
          <div class="medal" data-field="killtaculars"><img src="/medals/killtacular.png"><div class="medal-name">KILLTACULAR</div><div class="medal-count">×0</div></div>
          <div class="medal" data-field="killtrocities"><img src="/medals/killtrocity.png"><div class="medal-name">KILLTROCITY</div><div class="medal-count">×0</div></div>
          <div class="medal" data-field="killimanjaros"><img src="/medals/killimanjaro.png"><div class="medal-name">KILLIMANJARO</div><div class="medal-count">×0</div></div>
          <div class="medal" data-field="killtastrophes"><img src="/medals/killtastrophe.png"><div class="medal-name">KILLTASTROPHE</div><div class="medal-count">×0</div></div>
          <div class="medal" data-field="killpocalypses"><img src="/medals/killpocalypse.png"><div class="medal-name">KILLPOCALYPSE</div><div class="medal-count">×0</div></div>
          <div class="medal" data-field="killionaires"><img src="/medals/killionaire.png"><div class="medal-name">KILLIONAIRE</div><div class="medal-count">×0</div></div>
        </div>
      </section>

      <section id="recap" class="recap hidden">
        <div class="recap-head">
          <div class="recap-title">POST-GAME RECAP</div>
          <div id="recapResult" class="recap-result">VICTORY</div>
        </div>
        <div class="recap-grid">
          <div class="recap-block">
            <div id="recapGame" class="recap-game">0 KILLS &nbsp; 0 DEATHS &nbsp; 0.00 K/D</div>
            <div class="recap-kd-change">SESSION K/D &nbsp; <strong id="recapKd">0.00 → 0.00</strong></div>
            <div id="recapBest" class="recap-best"></div>
          </div>
          <div class="recap-block recap-feature">
            <img id="recapFeatureIcon" src="/medals/double-kill.png">
            <div><div id="recapFeatureName" class="recap-feature-name">SESSION UPDATE</div><div id="recapFeatureDetail" class="recap-feature-detail"></div></div>
          </div>
          <div id="recapDeltas" class="recap-block recap-deltas"></div>
        </div>
        <div class="recap-progress"><div id="recapProgress"></div></div>
      </section>
    </div>
  </main>

  <script>
    const $ = id => document.getElementById(id);
    const set = (id, value) => {
      const next = value || "--";
      if ($(id).textContent !== next) $(id).textContent = next;
    };
    const setBlankable = (id, value) => {
      const next = value || "";
      if ($(id).textContent !== next) $(id).textContent = next;
    };
    const params = new URLSearchParams(window.location.search);
    const isObsMode = params.get("mode") === "obs";
    const component = params.get("component") || "all";
    const urlScale = Math.max(0.25, Math.min(4, Number(params.get("scale")) || 1));
    document.body.classList.toggle("obs-mode", isObsMode);
    document.body.classList.toggle("component-mode", component !== "all");
    const componentSize = {
      network: [430, 132],
      wait: [360, 112],
      session: [920, 230],
      all: [1280, 150]
    }[component] || [1280, 150];
    const BASE_WIDTH = componentSize[0];
    const BASE_HEIGHT = componentSize[1];
    let lastPlacementKey = "";
    let lastOverlayData = null;
    const componentOverlay = document.querySelector(".overlay");
    componentOverlay.style.width = `${BASE_WIDTH}px`;
    componentOverlay.style.minHeight = `${BASE_HEIGHT}px`;

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

    function medalFile(name) {
      const files = {
        "Double Kill": "double-kill.png", "Triple Kill": "triple-kill.png",
        "Overkill": "overkill.png", "Killtacular": "killtacular.png",
        "Killtrocity": "killtrocity.png", "Killimanjaro": "killimanjaro.png",
        "Killtastrophe": "killtastrophe.png", "Killpocalypse": "killpocalypse.png",
        "Killionaire": "killionaire.png"
      };
      return files[name] || "double-kill.png";
    }

    function applyPlacement(data) {
      const overlay = document.querySelector(".overlay");
      if (!overlay) return;

      const componentPlacement = data.overlayPlacements?.[component];
      const placement = component !== "all" && componentPlacement
        ? componentPlacement
        : {
            leftRatio: data.overlayLeftRatio,
            topRatio: data.overlayTopRatio,
            widthRatio: data.overlayWidthRatio,
            heightRatio: data.overlayHeightRatio
          };

      const contentHeight = Math.max(BASE_HEIGHT, overlay.scrollHeight);
      const placementKey = isObsMode
        ? `${component}|obs|${urlScale}|${window.innerWidth}|${window.innerHeight}|${contentHeight}|${placement.leftRatio}|${placement.topRatio}|${placement.widthRatio}|${placement.heightRatio}`
        : `${component}|game|${urlScale}|${window.innerWidth}|${window.innerHeight}`;
      if (placementKey === lastPlacementKey) return;
      lastPlacementKey = placementKey;

      if (isObsMode) {
        const left = (placement.leftRatio ?? 0) * window.innerWidth;
        const top = (placement.topRatio ?? 0) * window.innerHeight;
        const width = Math.max(1, (placement.widthRatio ?? 0) * window.innerWidth);
        const height = Math.max(1, (placement.heightRatio ?? 0) * window.innerHeight);
        const widthScale = width > 1 ? width / BASE_WIDTH : 1;
        const heightScale = height > 1 ? height / contentHeight : widthScale;
        const availableHeight = Math.max(1, window.innerHeight - top - 4);
        const viewportScale = availableHeight / contentHeight;
        const scale = Math.max(0.1, Math.min(
          Math.min(widthScale, heightScale) * urlScale,
          viewportScale
        ));
        overlay.style.left = `${left}px`;
        overlay.style.top = `${top}px`;
        overlay.style.zoom = 1;
        overlay.style.transform = `scale(${scale})`;
        return;
      }

      const scale = Math.min(
        window.innerWidth / BASE_WIDTH,
        window.innerHeight / contentHeight
      );
      const fittedScale = Math.max(0.1, scale * urlScale);
      overlay.style.zoom = 1;
      overlay.style.transform = `scale(${fittedScale})`;
    }

    async function refresh() {
      try {
        const data = await fetch("/state.json", { cache: "no-store" }).then(r => r.json());
        lastOverlayData = data;
        applyPlacement(data);
        const hasPing = data.rttMs !== null && data.rttMs !== undefined;
        const hasNetwork = hasPing || data.uploadKilobytesPerSecond !== null || data.downloadKilobytesPerSecond !== null;
        $("network").classList.toggle("hidden", component !== "all" && component !== "network" || !data.showNetworkStats);
        setBlankable("server", hasNetwork && data.serverLabel !== "SERVER: --" ? data.serverLabel : "");
        setBlankable("ping", hasPing ? `Ping: ${data.rttMs} ms` : "");
        const hasJitter = data.jitterMs !== null && data.jitterMs !== undefined;
        setBlankable("jitter", hasJitter ? `Jitter: ${data.jitterMs.toFixed(1)} ms` : "");
        $("jitter").style.color = !hasJitter ? "" : data.jitterMs < 10 ? "var(--green)" : data.jitterMs < 20 ? "var(--cyan)" : "var(--red)";
        setBlankable("loss", hasNetwork ? `Loss: ${(data.packetLossPercent ?? 0).toFixed(0)}%` : "");
        setBlankable("up", data.uploadKilobytesPerSecond !== null && data.uploadKilobytesPerSecond !== undefined ? formatKb(data.uploadKilobytesPerSecond) : "");
        setBlankable("down", data.downloadKilobytesPerSecond !== null && data.downloadKilobytesPerSecond !== undefined ? formatKb(data.downloadKilobytesPerSecond) : "");
        setBlankable("upPackets", data.uploadPacketsPerSecond !== null && data.uploadPacketsPerSecond !== undefined ? `${data.uploadPacketsPerSecond.toFixed(0)} packets/s` : "");
        setBlankable("downPackets", data.downloadPacketsPerSecond !== null && data.downloadPacketsPerSecond !== undefined ? `${data.downloadPacketsPerSecond.toFixed(0)} packets/s` : "");
        const points = graphPoints(data.rttHistoryMs);
        $("graphLine").setAttribute("points", points);
        $("graphGlow").setAttribute("points", points);

        const showWait = data.showMatchmakingWait &&
          data.matchmakingWaitSeconds !== null && data.matchmakingWaitSeconds !== undefined &&
          (!data.matchmakingExpiresAtUtc || Date.now() < Date.parse(data.matchmakingExpiresAtUtc));
        $("matchmaking").classList.toggle("hidden", (component !== "all" && component !== "wait") || !showWait);
        if (showWait) {
          const estimate = Math.max(0, Number(data.matchmakingWaitSeconds));
          const elapsed = data.matchmakingStartedAtUtc
            ? Math.max(0, Math.floor((Date.now() - Date.parse(data.matchmakingStartedAtUtc)) / 1000))
            : 0;
          const formatDuration = seconds => seconds < 60 ? `${seconds} SEC` : seconds < 120 ? `~1 MIN` : `~${Math.ceil(seconds / 60)} MIN`;
          set("waitValue", `EST. WAIT ${formatDuration(estimate)}`);
          $("waitValue").classList.toggle("warn", estimate >= 30 && estimate < 90);
          $("waitValue").classList.toggle("long", estimate >= 90);
          const hasPopulation = data.matchmakingPopulation !== null && data.matchmakingPopulation !== undefined;
          const playlistName = data.matchmakingPlaylistName || "this playlist";
          const searchScope = data.matchmakingSearchScope || "all gametypes";
          set("waitDetail", hasPopulation
            ? `${data.matchmakingPopulation} player${data.matchmakingPopulation === 1 ? "" : "s"} searching ${playlistName} across ${searchScope} - your wait time may vary`
            : `elapsed ${Math.floor(elapsed / 60)}:${String(elapsed % 60).padStart(2, "0")}`);
        }

        // A dedicated session window exists only while that component is enabled,
        // so keep it rendered during startup and empty-state snapshots. The
        // combined page still honors the visibility setting from app state.
        const sessionAllowed = component === "session" ||
          (component === "all" && data.showSessionStats);
        // Keep the empty session on the exact same render path as a completed,
        // expired recap. Without this normalization the && expression returns
        // null (not false) until the first game creates a recap object.
        const recap = data.postGameRecap || {
          won: false,
          kills: 0,
          deaths: 0,
          gameKd: "0.00",
          previousSessionKd: "0.00",
          sessionKd: "0.00",
          bestSpree: 0,
          isNewBestSpree: false,
          featuredMedal: "\u2014",
          medalDeltas: [],
          capturedAtUtc: "1970-01-01T00:00:00Z",
          expiresAtUtc: "1970-01-01T00:00:00Z"
        };
        const showRecap = Boolean(
          sessionAllowed && Date.now() < Date.parse(recap.expiresAtUtc));
        $("session").classList.toggle("hidden", !sessionAllowed || showRecap);
        $("recap").classList.toggle("hidden", !showRecap);
        if (showRecap) {
          set("recapResult", recap.won ? "VICTORY" : "DEFEAT");
          $("recapResult").classList.toggle("loss", !recap.won);
          set("recapGame", `${recap.kills} KILLS   ${recap.deaths} DEATHS   ${recap.gameKd} K/D`);
          set("recapKd", `${recap.previousSessionKd} \u2192 ${recap.sessionKd}`);
          setBlankable("recapBest", recap.isNewBestSpree ? `NEW SESSION BEST  \u2022  ${recap.bestSpree} KILL SPREE` : `${recap.bestSpree} KILL SPREE`);
          const featured = recap.featuredMedal && recap.featuredMedal !== "\u2014" ? recap.featuredMedal : "SESSION UPDATE";
          set("recapFeatureName", featured.toUpperCase());
          const featureSource = `/medals/${medalFile(recap.featuredMedal)}`;
          if (!$("recapFeatureIcon").src.endsWith(featureSource)) $("recapFeatureIcon").src = featureSource;
          set("recapFeatureDetail", recap.featuredMedal && recap.featuredMedal !== "\u2014" ? "HIGHEST MULTIKILL THIS GAME" : "GAME ADDED TO SESSION");
          const deltaMarkup = (recap.medalDeltas || []).map(delta => `
            <div class="recap-delta">
              <img src="/medals/${medalFile(delta.name)}">
              <div class="recap-delta-name">${delta.name.toUpperCase()}</div>
              <div class="recap-delta-count">${delta.previousCount} \u2192 ${delta.newCount}</div>
            </div>`).join("") || `<div class="recap-delta-name">NO MULTIKILLS THIS GAME</div>`;
          if ($("recapDeltas").innerHTML !== deltaMarkup) $("recapDeltas").innerHTML = deltaMarkup;
          const start = Date.parse(recap.capturedAtUtc);
          const end = Date.parse(recap.expiresAtUtc);
          const remaining = Math.max(0, Math.min(1, (end - Date.now()) / Math.max(1, end - start)));
          $("recapProgress").style.transform = `scaleX(${remaining})`;
        }
        const games = data.gamesPlayed ?? 0;
        set("wins", `${data.wins ?? 0}W`);
        set("losses", `${data.losses ?? 0}L`);
        set("games", `${games} game${games === 1 ? "" : "s"}`);
        set("winRate", games > 0 ? `${Math.round(((data.wins ?? 0) / games) * 100)}%` : "0%");
        set("kd", data.sessionKd || "--");
        set("kills", `${data.kills ?? 0} Kills`);
        set("deaths", `${data.deaths ?? 0} Deaths`);
        document.querySelectorAll(".medal").forEach(el => {
          const count = Number(data[el.dataset.field] ?? 0);
          el.classList.toggle("earned", count > 0);
          const countElement = el.querySelector(".medal-count");
          const countText = `\u00D7${count}`;
          if (countElement.textContent !== countText) countElement.textContent = countText;
        });
      } catch {
      }
    }

    refresh();
    setInterval(refresh, 500);
    window.addEventListener("resize", () => {
      if (!lastOverlayData) return;
      lastPlacementKey = "";
      requestAnimationFrame(() => applyPlacement(lastOverlayData));
    });
  </script>
</body>
</html>
""";
}

internal sealed record ObsOverlaySnapshot(
    bool ShowSessionStats,
    bool ShowNetworkStats,
    bool ShowMatchmakingWait,
    int? MatchmakingWaitSeconds,
    int? MatchmakingPopulation,
    string MatchmakingPlaylistName,
    string MatchmakingSearchScope,
    DateTimeOffset? MatchmakingStartedAtUtc,
    DateTimeOffset? MatchmakingExpiresAtUtc,
    string ServerLabel,
    int? RttMs,
    double? JitterMs,
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
    ObsPostGameRecap? PostGameRecap,
    int BestSpree,
    int DoubleKills,
    int TripleKills,
    int Overkills,
    int Killtaculars,
    int Killtrocities,
    int Killimanjaros,
    int Killtastrophes,
    int Killpocalypses,
    int Killionaires,
    double OverlayLeftRatio,
    double OverlayTopRatio,
    double OverlayWidthRatio,
    double OverlayHeightRatio,
    IReadOnlyDictionary<string, ObsOverlayPlacement> OverlayPlacements)
{
    public static ObsOverlaySnapshot Empty { get; } = new(
        ShowSessionStats: true,
        ShowNetworkStats: true,
        ShowMatchmakingWait: false,
        MatchmakingWaitSeconds: null,
        MatchmakingPopulation: null,
        MatchmakingPlaylistName: "",
        MatchmakingSearchScope: "all gametypes",
        MatchmakingStartedAtUtc: null,
        MatchmakingExpiresAtUtc: null,
        ServerLabel: "SERVER: --",
        RttMs: null,
        JitterMs: null,
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
        PostGameRecap: null,
        BestSpree: 0,
        DoubleKills: 0,
        TripleKills: 0,
        Overkills: 0,
        Killtaculars: 0,
        Killtrocities: 0,
        Killimanjaros: 0,
        Killtastrophes: 0,
        Killpocalypses: 0,
        Killionaires: 0,
        OverlayLeftRatio: 0,
        OverlayTopRatio: 0,
        OverlayWidthRatio: 1280.0 / 1920.0,
        OverlayHeightRatio: 170.0 / 1080.0,
        OverlayPlacements: new Dictionary<string, ObsOverlayPlacement>());
}

internal sealed record ObsOverlayPlacement(
    double LeftRatio,
    double TopRatio,
    double WidthRatio,
    double HeightRatio);

internal sealed record ObsPostGameRecap(
    bool Won,
    long Kills,
    long Deaths,
    string GameKd,
    string PreviousSessionKd,
    string SessionKd,
    int BestSpree,
    bool IsNewBestSpree,
    string FeaturedMedal,
    IReadOnlyList<ObsMedalDelta> MedalDeltas,
    DateTimeOffset CapturedAtUtc,
    DateTimeOffset ExpiresAtUtc);

internal sealed record ObsMedalDelta(
    string Name,
    int PreviousCount,
    int NewCount,
    int EarnedCount);
