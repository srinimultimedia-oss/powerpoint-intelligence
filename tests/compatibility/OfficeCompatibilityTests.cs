using System;
using System.Collections.Specialized;
using System.IO;
using PowerPointIntelligence.Compatibility;

internal static class OfficeCompatibilityTests
{
    private static int passed;
    private static void Test(string name, Action action) { action(); Console.WriteLine("PASS " + name); passed++; }
    private static void Require(bool value) { if (!value) throw new Exception("Compatibility assertion failed."); }
    private static void Reject(Action action)
    { try { action(); } catch (FormatException) { return; } throw new Exception("Expected invalid configuration rejection."); }
    private static OfficeSettings Settings(string profile, string architecture = "any", string minimum = "16.0")
    {
        return OfficeSettings.Parse(new NameValueCollection {
            { "OfficeProfile", profile }, { "OfficeArchitecture", architecture }, { "OfficeMinimumVersion", minimum }
        });
    }
    public static int Main(string[] args)
    {
        try
        {
            foreach (var profile in new[] { "auto", "office2016", "office2019", "microsoft365" })
            {
                foreach (bool is64 in new[] { false, true })
                {
                    var localProfile = profile; var local64 = is64;
                    Test(localProfile + " common API on " + (local64 ? "x64" : "x86"), () => {
                        var result = OfficeCompatibility.Evaluate(Settings(localProfile), "16.0", "16.0.20326.20144", local64);
                        Require(result.Allowed && result.HostVersion == "16.0.20326.20144");
                        Require(result.Profile == localProfile && result.Architecture == (local64 ? "x64" : "x86"));
                    });
                }
            }
            Test("defaults choose common Office 16 API and either bitness", () => {
                var s = OfficeSettings.Parse(new NameValueCollection());
                Require(s.Profile == "auto" && s.Architecture == "any" && s.MinimumVersion.ToString() == "16.0");
            });
            Test("setting values normalize case and whitespace", () => Require(Settings(" OFFICE2019 ", " X64 ").Profile == "office2019"));
            Test("invalid profile rejected", () => Reject(() => Settings("office19")));
            Test("blank explicit profile rejected", () => Reject(() => Settings("")));
            Test("invalid architecture rejected", () => Reject(() => Settings("auto", "arm64")));
            Test("minimum must stay on Office 16 API", () => Reject(() => Settings("auto", "any", "19.0")));
            Test("malformed minimum rejected", () => Reject(() => Settings("auto", "any", "16.bad")));
            Test("x86 restriction permits x86 process", () => Require(OfficeCompatibility.Evaluate(Settings("office2016", "x86"), "16.0", null, false).Allowed));
            Test("x86 restriction blocks x64 process", () => Require(!OfficeCompatibility.Evaluate(Settings("office2016", "x86"), "16.0", null, true).Allowed));
            Test("x64 restriction permits x64 process", () => Require(OfficeCompatibility.Evaluate(Settings("office2019", "x64"), "16.0", null, true).Allowed));
            Test("x64 restriction blocks x86 process", () => Require(!OfficeCompatibility.Evaluate(Settings("office2019", "x64"), "16.0", null, false).Allowed));
            Test("Office 2013 API rejected", () => Require(!OfficeCompatibility.Evaluate(Settings("auto"), "15.0", "15.0.1.0", true).Allowed));
            Test("unknown future API rejected", () => Require(!OfficeCompatibility.Evaluate(Settings("microsoft365"), "17.0", "17.0.1.0", true).Allowed));
            Test("invalid application version rejected", () => Require(!OfficeCompatibility.Evaluate(Settings("auto"), "unknown", null, true).Allowed));
            Test("minimum build permits equal version", () => Require(OfficeCompatibility.Evaluate(Settings("office2019", "any", "16.0.10396.20023"), "16.0", "16.0.10396.20023", true).Allowed));
            Test("minimum build blocks older version", () => Require(!OfficeCompatibility.Evaluate(Settings("office2019", "any", "16.0.10396.20023"), "16.0", "16.0.10396.20022", true).Allowed));
            Test("no full build cannot satisfy full minimum", () => Require(!OfficeCompatibility.Evaluate(Settings("auto", "any", "16.0.10000.0"), "16.0", null, true).Allowed));
            Test("even zero build floor requires known build", () => Require(!OfficeCompatibility.Evaluate(Settings("auto", "any", "16.0.0.0"), "16.0", null, true).Allowed));
            Test("missing executable build can use API baseline", () => Require(OfficeCompatibility.Evaluate(Settings("auto"), "16.0", null, true).Allowed));
            Test("inconsistent executable version is not trusted", () => Require(!OfficeCompatibility.Evaluate(Settings("auto"), "16.0", "15.0.1234.0", true).Allowed));
            Test("malformed executable version is not trusted", () => Require(!OfficeCompatibility.Evaluate(Settings("auto"), "16.0", "bad", true).Allowed));
            Test("API major is not a product detector", () => {
                var result = OfficeCompatibility.Evaluate(Settings("office2016"), "16.0", "16.0.20326.20144", true);
                Require(result.Allowed && result.Message.Contains("edition not auto-detected"));
            });
            Test("shipped add-in config parses", () => {
                var s = OfficeSettings.Load(args[0]); Require(s.Profile == "auto" && s.Architecture == "any");
            });
            Test("missing configuration uses portable defaults", () => {
                Require(OfficeSettings.Load(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".config")).Profile == "auto");
            });
            Test("invalid deployed setting rejected", () => {
                var file = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".config");
                try {
                    File.WriteAllText(file, "<configuration><appSettings><add key='OfficeProfile' value='office19'/></appSettings></configuration>");
                    Reject(() => OfficeSettings.Load(file));
                } finally { File.Delete(file); }
            });
            Test("unknown deployed setting rejected", () => {
                var file = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".config");
                try {
                    File.WriteAllText(file, "<configuration><appSettings><add key='OfficeProfiel' value='office2019'/></appSettings></configuration>");
                    Reject(() => OfficeSettings.Load(file));
                } finally { File.Delete(file); }
            });
            Test("malformed deployed XML does not use defaults", () => {
                var file = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".config");
                try {
                    File.WriteAllText(file, "<configuration><broken>");
                    bool rejected = false;
                    try { OfficeSettings.Load(file); } catch (System.Configuration.ConfigurationErrorsException) { rejected = true; }
                    Require(rejected);
                } finally { File.Delete(file); }
            });
            Test("test executable uses requested process architecture", () => {
                Require(args[1] == "AnyCPU" || Environment.Is64BitProcess == (args[1] == "x64"));
            });
            Console.WriteLine("Office compatibility tests passed: " + passed); return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex.GetType().Name + ": compatibility tests failed."); return 1; }
    }
}
