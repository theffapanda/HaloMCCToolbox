param(
    [string]$OldDll = 'C:\Program Files (x86)\Steam\steamapps\content\app_976730\depot_1064225\haloreach\haloreach.dll',
    [string]$CurrentDll = 'C:\Program Files (x86)\Steam\steamapps\common\Halo The Master Chief Collection\haloreach\haloreach.dll',
    [string]$RawFile,
    [string]$RawDirectory,
    [uint32[]]$TargetHashes = @(180688984,714809486,756507193,1143893577,757649136,981019040)
)

$ErrorActionPreference = 'Stop'

if (-not ('ReachShaderDiff.Scanner' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.IO;

namespace ReachShaderDiff {
    public sealed class Entry {
        public long Offset;
        public int Length;
        public uint Crc32;
    }

    public static class Scanner {
        static readonly uint[] Table = BuildTable();
        static uint[] BuildTable() {
            var table = new uint[256];
            for (uint i = 0; i < table.Length; ++i) {
                uint c = i;
                for (int j = 0; j < 8; ++j) c = (c & 1) != 0 ? 0xEDB88320U ^ (c >> 1) : c >> 1;
                table[i] = c;
            }
            return table;
        }
        static uint Crc(byte[] data, int offset, int length) {
            uint crc = 0xFFFFFFFFU;
            for (int i = 0; i < length; ++i) crc = Table[(crc ^ data[offset + i]) & 0xFF] ^ (crc >> 8);
            return crc ^ 0xFFFFFFFFU;
        }
        public static uint HashFile(string path) {
            byte[] data = File.ReadAllBytes(path);
            return Crc(data, 0, data.Length);
        }
        public static string[] FindHashes(string directory, uint[] targets) {
            var wanted = new HashSet<uint>(targets);
            var matches = new List<string>();
            foreach (string path in Directory.EnumerateFiles(directory)) {
                byte[] data = File.ReadAllBytes(path);
                uint crc = Crc(data, 0, data.Length);
                if (wanted.Contains(crc)) matches.Add(Path.GetFileName(path) + "=" + crc);
            }
            return matches.ToArray();
        }
        public static Entry[] Scan(string path) {
            byte[] data = File.ReadAllBytes(path);
            var result = new List<Entry>();
            for (int i = 0; i + 32 <= data.Length; ++i) {
                if (data[i] != 0x44 || data[i+1] != 0x58 || data[i+2] != 0x42 || data[i+3] != 0x43) continue;
                uint length = BitConverter.ToUInt32(data, i + 24);
                uint chunks = BitConverter.ToUInt32(data, i + 28);
                if (length >= 32 && length <= 2000000 && chunks < 128 && (long)i + length <= data.Length)
                    result.Add(new Entry { Offset = i, Length = (int)length, Crc32 = Crc(data, i, (int)length) });
                i += 3;
            }
            return result.ToArray();
        }
    }
}
'@
}

if ($RawFile) {
    "raw_file=$RawFile crc32=$([ReachShaderDiff.Scanner]::HashFile($RawFile))"
    exit 0
}
if ($RawDirectory) {
    [ReachShaderDiff.Scanner]::FindHashes($RawDirectory, $TargetHashes)
    exit 0
}

$old = [ReachShaderDiff.Scanner]::Scan($OldDll)
$current = [ReachShaderDiff.Scanner]::Scan($CurrentDll)
$oldHashes = @($old | ForEach-Object Crc32 | Sort-Object -Unique)
$currentHashes = @($current | ForEach-Object Crc32 | Sort-Object -Unique)
$oldOnly = @($oldHashes | Where-Object { $_ -notin $currentHashes })
$currentOnly = @($currentHashes | Where-Object { $_ -notin $oldHashes })
$shared = @($oldHashes | Where-Object { $_ -in $currentHashes })

"old_containers=$($old.Count) old_unique=$($oldHashes.Count)"
"current_containers=$($current.Count) current_unique=$($currentHashes.Count)"
"shared=$($shared.Count) old_only=$($oldOnly.Count) current_only=$($currentOnly.Count)"
'OLD_ONLY=' + ($oldOnly -join ',')
'CURRENT_ONLY=' + ($currentOnly -join ',')
