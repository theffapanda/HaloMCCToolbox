using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace HaloToolbox;

internal static class RejoinFixPaths
{
    public static string RootDirectory =>
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "HaloMCCToolbox",
            "RejoinFix");

    public static string LastHandleFile => Path.Combine(RootDirectory, "last-handle.json");
    public static string LastMatchSessionFile => Path.Combine(RootDirectory, "last-match-session.json");
    public static string LastSquadStateFile => Path.Combine(RootDirectory, "last-squad-state.json");
    public static string LastGameServerFile => Path.Combine(RootDirectory, "last-game-server.json");
    public static string ProxyRootCertificateFile => Path.Combine(RootDirectory, "proxy-root.pfx");
    public static string RecentLogFile => Path.Combine(RootDirectory, "rejoin-fix.log");

    public static void EnsureRootDirectory() => Directory.CreateDirectory(RootDirectory);

    public static IReadOnlyList<string> GetExportFiles()
    {
        EnsureRootDirectory();

        string[] preferred =
        {
            RecentLogFile,
            LastHandleFile,
            LastMatchSessionFile,
            LastSquadStateFile,
            LastGameServerFile,
        };

        return preferred.Where(File.Exists).ToArray();
    }
}

internal static class RejoinFixDiagnostics
{
    private static readonly object Sync = new();
    private const int MaxEntries = 100;

    public static void Info(string category, string message) => Write("INFO", category, message);
    public static void Warn(string category, string message) => Write("WARN", category, message);
    public static void Error(string category, string message) => Write("ERROR", category, message);

    public static void Write(string level, string category, string message)
    {
        lock (Sync)
        {
            RejoinFixPaths.EnsureRootDirectory();

            var line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{level}] [{category}] {message}";
            List<string> lines;
            try
            {
                lines = File.Exists(RejoinFixPaths.RecentLogFile)
                    ? File.ReadAllLines(RejoinFixPaths.RecentLogFile).Where(l => !string.IsNullOrWhiteSpace(l)).ToList()
                    : new List<string>();
            }
            catch
            {
                lines = new List<string>();
            }

            lines.Add(line);
            if (lines.Count > MaxEntries)
                lines = lines.Skip(lines.Count - MaxEntries).ToList();

            File.WriteAllLines(RejoinFixPaths.RecentLogFile, lines);
        }
    }
}
