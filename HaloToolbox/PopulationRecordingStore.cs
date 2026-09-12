using System.IO;
using System.Text;
using System.Text.Json;

namespace HaloToolbox;

// Only population counts and queue names are stored here, never authorization.
internal sealed class PopulationRecordingStore
{
    internal const long MaximumBytes = 100 * 1024 * 1024;
    private readonly string _directory;
    private readonly string _preference;
    private string? _sessionFile;
    public bool Enabled { get; private set; }

    public PopulationRecordingStore(string directory)
    {
        _directory = directory;
        _preference = Path.Combine(directory, "enabled.txt");
        // A fresh install performs no writes and defaults to live-only history.
        Enabled = File.Exists(_preference) && File.ReadAllText(_preference).Trim() == "true";
    }

    public void SetEnabled(bool enabled)
    {
        Directory.CreateDirectory(_directory);
        File.WriteAllText(_preference + ".tmp", enabled ? "true" : "false");
        File.Move(_preference + ".tmp", _preference, true);
        Enabled = enabled;
        _sessionFile = null;
    }

    public void Append(IReadOnlyList<MatchmakingPopulationSample> samples)
    {
        if (!Enabled || samples.Count == 0) return;
        Directory.CreateDirectory(_directory);
        // One complete JSON line per poll. A partial final line from a crash is skipped on read.
        string line = JsonSerializer.Serialize(samples) + "\n";
        long bytes = Encoding.UTF8.GetByteCount(line);
        if (Directory.EnumerateFiles(_directory, "*.jsonl").Sum(f => new FileInfo(f).Length) + bytes > MaximumBytes)
            throw new IOException("Recording storage reached 100 MB. Move saved recordings out of the recording folder to continue.");
        _sessionFile ??= Path.Combine(_directory, $"{DateTime.UtcNow:yyyyMMdd-HHmmss}-{Guid.NewGuid():N}.jsonl");
        using var stream = new FileStream(_sessionFile, FileMode.Append, FileAccess.Write, FileShare.Read);
        byte[] payload = Encoding.UTF8.GetBytes(line);
        stream.Write(payload);
        stream.Flush(flushToDisk: true);
    }

    public List<MatchmakingPopulationSample> ReadAll(out int skippedLines)
    {
        var samples = new List<MatchmakingPopulationSample>();
        skippedLines = 0;
        if (!Directory.Exists(_directory)) return samples;
        foreach (string file in Directory.EnumerateFiles(_directory, "*.jsonl").OrderBy(f => f))
        {
            foreach (string line in File.ReadLines(file))
            {
                try
                {
                    var batch = JsonSerializer.Deserialize<List<MatchmakingPopulationSample>>(line);
                    if (batch is null) { skippedLines++; continue; }
                    samples.AddRange(batch.Where(s => s is not null && s.Population >= 0 &&
                        !string.IsNullOrWhiteSpace(s.HopperName)));
                }
                catch (JsonException) { skippedLines++; }
            }
        }
        return samples.OrderBy(s => s.CapturedAt).ToList();
    }
}
