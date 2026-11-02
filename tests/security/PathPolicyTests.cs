using System;
using System.IO;
using PowerPointIntelligence.Security;

namespace PowerPointIntelligence.Tests.Security
{
    public static class PathPolicyTests
    {
        public static int Run(string repo)
        {
            int count = 0;
            Action<bool> check = valid => { if (!valid) throw new Exception("Path authority regression failed."); count++; };
            Action<Action> reject = action => { try { action(); } catch (InvalidDataException) { count++; return; } throw new Exception("Unsafe path accepted."); };
            string owned = Path.Combine(repo, "artifacts", "path-policy-tests");
            check(PathPolicy.Within(owned, "nested/source.csv") == Path.Combine(owned, "nested", "source.csv"));
            check(PathPolicy.TestPath(Path.Combine(owned, "fixture.json")) == Path.Combine(owned, "fixture.json"));
            check(PathPolicy.Selected(@"\\server\share\source.xlsx") == @"\\server\share\source.xlsx");
            check(PathPolicy.Selected(Path.Combine(Path.GetTempPath(), "external-selected.csv")).EndsWith("external-selected.csv", StringComparison.Ordinal));
            check(PathPolicy.Configured("templates", owned) == Path.Combine(owned, "templates"));
            foreach (string path in new[] { "../escape", "..\\escape", "C:relative", "C:\\a\\..\\escape", "C:\\a\\file:stream", @"\\?\C:\escape", @"\\.\NUL", "C:\\a\\CON.txt", "C:\\a\\trailing.", "C:\\a\\trailing ", "C:\\a\\bad*", "C:\\a\\bad\nname" })
                reject(() => PathPolicy.Within(owned, path));
            reject(() => PathPolicy.Within(owned, owned + "-sibling\\outside.json"));
            reject(() => PathPolicy.Configured("../outside", owned));
            reject(() => PathPolicy.TestPath(repo));
            reject(() => PathPolicy.TestPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "unowned-output.json")));
            reject(() => PathPolicy.ConfigFile(Path.Combine(owned, "not-config.exe")));
            reject(() => PathPolicy.Selected("\\rooted-but-drive-dependent"));
            reject(() => PathPolicy.Selected("/rooted-but-drive-dependent"));
            Directory.CreateDirectory(owned);
            string target = Path.Combine(owned, "target");
            Directory.CreateDirectory(target);
            Console.WriteLine("PASS " + count + " path authority assertions: safe children, external source selection and restricted test/configuration paths.");
            return count;
        }
    }
}
