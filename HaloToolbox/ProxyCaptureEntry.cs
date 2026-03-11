using System;
using System.Collections.Generic;

namespace HaloToolbox;

public class ProxyCaptureEntry
{
    public DateTime Timestamp { get; init; } = DateTime.Now;
    public string Method      { get; init; } = "";
    public string Url         { get; init; } = "";
    public string Host        { get; init; } = "";
    public string Path        { get; init; } = "";

    // Request
    public Dictionary<string, string> RequestHeaders { get; init; } = new();
    public string RequestBody { get; set; } = "";

    // Response (filled after round-trip completes)
    public int StatusCode { get; set; }
    public Dictionary<string, string> ResponseHeaders { get; set; } = new();
    public string ResponseBody { get; set; } = "";

    // Display helpers
    public string TimestampStr  => Timestamp.ToString("HH:mm:ss.fff");
    public string StatusDisplay => StatusCode == 0 ? "…" : StatusCode.ToString();

    public bool HasSpartanToken =>
        RequestHeaders.TryGetValue("x-343-authorization-spartan", out var v)
        && !string.IsNullOrEmpty(v);

    public string? ExtractedSpartanToken =>
        RequestHeaders.TryGetValue("x-343-authorization-spartan", out var v) ? v : null;
}
