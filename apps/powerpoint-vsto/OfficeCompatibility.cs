using System;
using System.Collections.Specialized;
using System.Configuration;
using System.IO;

namespace PowerPointIntelligence.Compatibility
{
    public sealed class OfficeSettings
    {
        public string Profile { get; private set; }
        public string Architecture { get; private set; }
        public Version MinimumVersion { get; private set; }

        public static OfficeSettings Parse(NameValueCollection values)
        {
            if (values == null) throw new ArgumentNullException("values");
            string profile = (values["OfficeProfile"] ?? "auto").Trim().ToLowerInvariant();
            string architecture = (values["OfficeArchitecture"] ?? "any").Trim().ToLowerInvariant();
            string minimum = (values["OfficeMinimumVersion"] ?? "16.0").Trim();
            if (profile != "auto" && profile != "office2016" && profile != "office2019" && profile != "microsoft365")
                throw new FormatException("Unknown Office compatibility profile.");
            if (architecture != "any" && architecture != "x86" && architecture != "x64")
                throw new FormatException("Unknown Office architecture setting.");
            Version version;
            if (!Version.TryParse(minimum, out version) || version.Major != 16 || version.Minor != 0)
                throw new FormatException("Office minimum version must be a 16.0 version or build.");
            return new OfficeSettings { Profile = profile, Architecture = architecture, MinimumVersion = version };
        }

        public static OfficeSettings Load(string configurationPath)
        {
            // Map the add-in's config explicitly: never read or modify POWERPNT.EXE.config.
            if (!File.Exists(configurationPath)) return Parse(new NameValueCollection());
            var config = ConfigurationManager.OpenMappedExeConfiguration(
                new ExeConfigurationFileMap { ExeConfigFilename = configurationPath }, ConfigurationUserLevel.None);
            var values = new NameValueCollection();
            foreach (string key in config.AppSettings.Settings.AllKeys)
            {
                if (key != "OfficeProfile" && key != "OfficeArchitecture" && key != "OfficeMinimumVersion")
                    throw new FormatException("Unknown add-in compatibility setting.");
                values.Add(key, config.AppSettings.Settings[key].Value);
            }
            return Parse(values);
        }
    }

    public sealed class OfficeCompatibilityResult
    {
        public bool Allowed { get; internal set; }
        public string Profile { get; internal set; }
        public string Architecture { get; internal set; }
        public string HostVersion { get; internal set; }
        public string Message { get; internal set; }
    }

    public static class OfficeCompatibility
    {
        public static OfficeCompatibilityResult Evaluate(OfficeSettings settings, string applicationVersion,
            string executableVersion, bool is64BitProcess)
        {
            if (settings == null) throw new ArgumentNullException("settings");
            var result = new OfficeCompatibilityResult { Profile = settings.Profile,
                Architecture = is64BitProcess ? "x64" : "x86", HostVersion = "unknown" };
            Version api;
            if (!Version.TryParse(applicationVersion, out api) || api.Major != 16 || api.Minor != 0)
                return Deny(result, "PowerPoint must expose the Office 16.0 API.");
            Version actual = api;
            if (!String.IsNullOrWhiteSpace(executableVersion))
            {
                if (!Version.TryParse(executableVersion, out actual) || actual.Major != api.Major || actual.Minor != api.Minor)
                    return Deny(result, "PowerPoint executable version could not be verified.");
            }
            result.HostVersion = actual.ToString();
            if (settings.Architecture != "any" && settings.Architecture != result.Architecture)
                return Deny(result, "Configured Office architecture does not match the running PowerPoint process.");
            if (settings.MinimumVersion.Build >= 0 && actual.Build < 0)
                return Deny(result, "A full PowerPoint executable build is required for the configured build-specific minimum.");
            if (Normalize(actual) < Normalize(settings.MinimumVersion))
                return Deny(result, "PowerPoint does not meet the configured minimum version; a full executable build is required for build-specific limits.");
            result.Allowed = true;
            result.Message = "Office profile: " + result.Profile + "; host: " + result.HostVersion + "; " + result.Architecture +
                ". Common Office 16 API enabled; edition not auto-detected. Runtime compatibility still requires validation on this Office build.";
            return result;
        }
        private static Version Normalize(Version version)
        { return new Version(version.Major, version.Minor, Math.Max(0, version.Build), Math.Max(0, version.Revision)); }
        private static OfficeCompatibilityResult Deny(OfficeCompatibilityResult result, string reason)
        { result.Allowed = false; result.Message = reason; return result; }
    }
}
