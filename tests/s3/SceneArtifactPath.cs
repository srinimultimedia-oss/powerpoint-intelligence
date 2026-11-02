using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PowerPointIntelligence.Tests.Security
{
    // Only data-derived child names are restricted; the caller's output root remains configurable.
    public static class SceneArtifactPath
    {
        public static string Resolve(string root, string name)
        {
            // Inspect raw characters before path APIs: .NET Framework IsPathRooted
            // throws on invalid characters. Separators/colon below already reject roots.
            if (String.IsNullOrWhiteSpace(name) || name.Length > 128 || name == "." || name == ".." ||
                name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 ||
                name.IndexOfAny(new[] { '/', '\\', ':' }) >= 0 || name.Any(Char.IsControl) ||
                name.EndsWith(".", StringComparison.Ordinal) || name.EndsWith(" ", StringComparison.Ordinal))
                throw new InvalidDataException("Corpus name must be a bounded safe basename.");
            string stem = name.Split('.')[0].TrimEnd(' ', '.').ToUpperInvariant();
            if (new[] { "CON", "PRN", "AUX", "NUL" }.Contains(stem) ||
                (stem.Length == 4 && (stem.StartsWith("COM", StringComparison.Ordinal) || stem.StartsWith("LPT", StringComparison.Ordinal)) &&
                 "123456789\u00b9\u00b2\u00b3".IndexOf(stem[3]) >= 0))
                throw new InvalidDataException("Corpus name cannot identify a Windows device.");
            string directory = Path.GetFullPath(root).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar;
            string candidate = Path.GetFullPath(Path.Combine(directory, name + ".v1.json"));
            if (!candidate.StartsWith(directory, StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("Corpus artifact escaped its output directory.");
            return candidate;
        }

        public static void ValidateNames(string root, IEnumerable<string> names)
        {
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string name in names)
            {
                Resolve(root, name);
                if (!seen.Add(name)) throw new InvalidDataException("Corpus artifact names must be unique.");
            }
        }
    }
}
