using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace PowerPointIntelligence.Security
{
    // Input selection and generated-output ownership are different authorities.
    // An explicitly selected source may be outside the repository; CLI evidence
    // may only be written in a repository workspace or the OS temporary area.
    internal static class PathPolicy
    {
        static void Syntax(string value)
        {
            if (String.IsNullOrWhiteSpace(value) || value.Length > 32767 || value.Any(Char.IsControl) ||
                value.StartsWith(@"\\?\", StringComparison.Ordinal) || value.StartsWith(@"\\.\", StringComparison.Ordinal))
                throw new InvalidDataException("Invalid filesystem path.");
            string normalized = value.Replace('/', '\\');
            foreach (string component in normalized.Split('\\'))
            {
                if (component == "." || component == ".." || component.EndsWith(".", StringComparison.Ordinal) || component.EndsWith(" ", StringComparison.Ordinal))
                    throw new InvalidDataException("Ambiguous or traversing path component.");
                if (component.Length == 0) continue;
                bool drive = component.Length == 2 && Char.IsLetter(component[0]) && component[1] == ':';
                if ((!drive && component.IndexOf(':') >= 0) || component.IndexOfAny(new[] { '<', '>', '"', '|', '?', '*' }) >= 0)
                    throw new InvalidDataException("Device, stream or wildcard path is unsupported.");
                string stem = component.Split('.')[0].TrimEnd(' ', '.').ToUpperInvariant();
                if (new[] { "CON", "PRN", "AUX", "NUL" }.Contains(stem) ||
                    (stem.Length == 4 && (stem.StartsWith("COM", StringComparison.Ordinal) || stem.StartsWith("LPT", StringComparison.Ordinal)) && "123456789\u00b9\u00b2\u00b3".IndexOf(stem[3]) >= 0))
                    throw new InvalidDataException("Windows device names are unsupported.");
            }
        }

        internal static string Selected(string path)
        {
            Syntax(path);
            string windows = path.Replace('/', '\\');
            bool drive = windows.Length >= 3 && Char.IsLetter(windows[0]) && windows[1] == ':' && windows[2] == '\\';
            bool unc = windows.StartsWith("\\\\", StringComparison.Ordinal) && windows.Substring(2).Split('\\').Length >= 2;
            if (!drive && !unc)
                throw new InvalidDataException("Select a fully qualified filesystem path.");
            string full = Path.GetFullPath(path);
            Syntax(full);
            return full;
        }

        internal static bool IsWithin(string root, string candidate)
        {
            string fullRoot = Path.GetFullPath(root).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            string full = Path.GetFullPath(candidate);
            return String.Equals(fullRoot, full.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar), StringComparison.OrdinalIgnoreCase) ||
                full.StartsWith(fullRoot + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase);
        }

        internal static string Within(string root, string path)
        {
            string approved = Selected(root);
            Syntax(path);
            string full = Selected(Path.IsPathRooted(path) ? path : Path.Combine(approved, path));
            if (!IsWithin(approved, full)) throw new InvalidDataException("Path escapes its authorized root.");
            string relative = full.Substring(approved.TrimEnd('\\', '/').Length).TrimStart('\\', '/');
            string safe = approved;
            foreach (string component in relative.Split(new[] { '\\', '/' }, StringSplitOptions.RemoveEmptyEntries))
            {
                // GetFileName removes path authority; Syntax rejects dot/device/ADS
                // names before this explicit component-by-component reconstruction.
                string name = Path.GetFileName(component);
                if (!String.Equals(name, component, StringComparison.Ordinal)) throw new InvalidDataException("Invalid filename component.");
                safe = Path.Combine(safe, name);
            }
            if (!String.Equals(safe.TrimEnd('\\', '/'), full.TrimEnd('\\', '/'), StringComparison.OrdinalIgnoreCase)) throw new InvalidDataException("Path reconstruction changed the requested location.");
            return safe;
        }

        internal static string ApprovedRoot(string requested, IEnumerable<string> approvedRoots)
        {
            // Stored legacy data may select an independently approved root. It
            // cannot authorize a new root, and is never returned to an IO sink.
            string comparison = Selected(requested).TrimEnd('\\', '/');
            foreach (string approved in approvedRoots)
            {
                string root = Selected(approved);
                if (String.Equals(root.TrimEnd('\\', '/'), comparison, StringComparison.OrdinalIgnoreCase)) return root;
            }
            throw new InvalidDataException("Legacy source is not an approved library root. Add or configure the folder in Library before retrying migration; the original registration was preserved.");
        }

        internal static string Configured(string value, string baseDirectory)
        {
            return Path.IsPathRooted(value) ? Selected(value) : Within(baseDirectory, value);
        }

        internal static void RejectLinkedAncestors(string path)
        {
            // Generated workspaces/index/configuration must not follow directory
            // junctions. This is intentionally not imposed on selected input files.
            string current = Selected(path);
            while (!String.IsNullOrEmpty(current))
            {
                if ((Directory.Exists(current) || File.Exists(current)) && (File.GetAttributes(current) & FileAttributes.ReparsePoint) != 0)
                    throw new InvalidDataException("Owned filesystem paths cannot traverse reparse points.");
                current = Path.GetDirectoryName(current);
            }
        }

        static IEnumerable<string> TestRoots()
        {
            var roots = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            string location = Assembly.GetExecutingAssembly().Location;
            foreach (string start in new[] { String.IsNullOrEmpty(location) ? null : Path.GetDirectoryName(location), Environment.CurrentDirectory })
            {
                string current = start;
                while (!String.IsNullOrEmpty(current))
                {
                    if (File.Exists(Path.Combine(current, "apps", "powerpoint-vsto", "PowerPointIntelligence.csproj"))) { roots.Add(Selected(current)); break; }
                    current = Path.GetDirectoryName(current);
                }
            }
            roots.Add(Selected(Path.GetTempPath()));
            return roots;
        }

        internal static string TestPath(string path)
        {
            Syntax(path);
            string full = Selected(Path.GetFullPath(path));
            // The requested string selects a trusted workspace but never becomes
            // the filesystem root. Rebuild only validated filename components.
            foreach (string root in TestRoots())
            {
                if (!IsWithin(root, full) || String.Equals(root.TrimEnd('\\', '/'), full.TrimEnd('\\', '/'), StringComparison.OrdinalIgnoreCase)) continue;
                string relative = full.Substring(root.TrimEnd('\\', '/').Length).TrimStart('\\', '/');
                string safe = Within(root, relative);
                RejectLinkedAncestors(safe);
                return safe;
            }
            throw new InvalidDataException("Test paths must belong to a repository or OS temporary workspace.");
        }

        internal static string ConfigFile(string path)
        {
            string full = Path.IsPathRooted(path) ? Selected(path) : Within(Environment.CurrentDirectory, path);
            if (!new[] { ".config", ".xml" }.Contains(Path.GetExtension(full), StringComparer.OrdinalIgnoreCase))
                throw new InvalidDataException("Expected a configuration file.");
            RejectLinkedAncestors(full);
            return full;
        }
    }
}
