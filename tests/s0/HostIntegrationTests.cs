using System;
using System.IO;
using System.Linq;
using PowerPointIntelligence.Host;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

internal static class HostIntegrationTests
{
    private static void CloseFixture(PowerPoint.Presentation presentation)
    {
        if (presentation == null) return;
        try
        {
            presentation.Saved = Office.MsoTriState.msoTrue;
            presentation.Close();
        }
        catch (System.Runtime.InteropServices.COMException ex)
        {
            // Preserve the primary test failure when Office is already disconnected.
            Console.Error.WriteLine("Fixture cleanup failed: HRESULT=0x" + ex.ErrorCode.ToString("X8"));
            throw;
        }
    }
    private static void Require(bool condition, string name)
    { if (!condition) throw new Exception("FAIL " + name); Console.WriteLine("PASS " + name); }

    private static void WaitForNativeEdit(PowerPoint.Chart chart)
    {
        var timer = System.Diagnostics.Stopwatch.StartNew();
        do
        {
            // An external STA runner has no VSTO UI message loop. Let Office
            // finish propagating the embedded workbook change before testing capture.
            System.Windows.Forms.Application.DoEvents();
            chart.Refresh();
            var series = (PowerPoint.SeriesCollection)chart.SeriesCollection();
            var values = PowerPointHost.NativeSeriesValues(series,1);
            if (values.Length > 0 && values[0] != null
                && Convert.ToDouble(values[0], System.Globalization.CultureInfo.InvariantCulture) == 42.0)
            { Console.WriteLine("PASS native chart observes committed workbook edit"); return; }
            System.Threading.Thread.Sleep(50);
        } while (timer.Elapsed < TimeSpan.FromSeconds(10));
        throw new Exception("FAIL native chart did not observe workbook edit within 10 seconds; metadata capture not attempted.");
    }

    // Explicit opt-in only. Runs visibly on an interactive workstation, never in CI.
    // Does not quit PowerPoint or close presentations owned by the user.
    [STAThread]
    public static int Main(string[] args)
    {
        if (args.Length == 1 && args[0] == "--series-vector-tests")
        {
            bool eventRan = false;
            Require(OfficeEventBoundary.Run(() => eventRan = true, s => { }) && eventRan,
                "Office event boundary runs successful callback");
            string eventError = null;
            Require(!OfficeEventBoundary.Run(() => { throw new System.Runtime.InteropServices.COMException("private workbook text", unchecked((int)0x800706BA)); }, s => eventError = s)
                && eventError.Contains("800706BA") && !eventError.Contains("private workbook text"),
                "Office event boundary contains COM failure without leaking contents");
            int calls = 0, pauses = 0;
            PowerPointHost.RetryRejectedCall(() => {
                if (++calls < 3) throw new System.Runtime.InteropServices.COMException("busy", unchecked((int)0x80010001));
            }, () => pauses++, 3);
            Require(calls == 3 && pauses == 2, "rejected call retries then succeeds");
            calls = 0;
            var rejected = new System.Runtime.InteropServices.COMException("busy", unchecked((int)0x80010001));
            bool limitPropagated = false;
            try { PowerPointHost.RetryRejectedCall(() => { calls++; throw rejected; }, () => { }, 3); }
            catch (System.Runtime.InteropServices.COMException ex)
            { limitPropagated = Object.ReferenceEquals(ex, rejected); }
            Require(limitPropagated, "rejected call retry limit preserves failure");
            Require(calls == 3, "rejected call attempts bounded");
            calls = 0;
            var disconnected = new System.Runtime.InteropServices.COMException("disconnected", unchecked((int)0x800706BA));
            bool propagated = false;
            try { PowerPointHost.RetryRejectedCall(() => { calls++; throw disconnected; }, () => { }, 3); }
            catch (System.Runtime.InteropServices.COMException ex) { propagated = Object.ReferenceEquals(ex, disconnected); }
            Require(propagated && calls == 1, "RPC disconnection is never retried");
            Require(PowerPointHost.SeriesVector(new double[] { 42, -2 }).SequenceEqual(new object[] { 42.0, -2.0 }), "typed numeric SAFEARRAY");
            var bounded = Array.CreateInstance(typeof(object), new[] { 3 }, new[] { 1 });
            bounded.SetValue("A", 1); bounded.SetValue(null, 2); bounded.SetValue("C", 3);
            Require(PowerPointHost.SeriesVector(bounded).SequenceEqual(new object[] { "A", null, "C" }), "one-based SAFEARRAY and missing value");
            Require(PowerPointHost.SeriesVector(42.0).Single().Equals(42.0), "scalar series value");
            Require(PowerPointHost.SeriesVector("Category").Single().Equals("Category"), "scalar category remains whole string");
            Require(PowerPointHost.SeriesVector(null).Single() == null, "missing scalar");
            try { PowerPointHost.SeriesVector(new object[1, 2]); }
            catch (InvalidDataException) { Console.WriteLine("PASS rejects multidimensional series"); return 0; }
            throw new Exception("FAIL rejects multidimensional series");
        }
        if (args.Length != 2 || args[0] != "--interactive-office")
        { Console.Error.WriteLine("Usage: HostIntegrationTests --interactive-office <new-output-directory>"); return 2; }
        var directory = Path.GetFullPath(args[1]);
        if (Directory.Exists(directory)) throw new IOException("Use a new fixture directory.");
        Directory.CreateDirectory(directory);
        PowerPoint.Presentation deck = null, copyDeck = null;
        bool testFailed = false;
        try
        {
            var app = new PowerPoint.Application();
            if (app.Presentations.Count > 0) throw new InvalidOperationException("Save and close user presentations before runtime validation. No fixture was created.");
            app.Visible = Office.MsoTriState.msoTrue;
            Console.WriteLine("Office version=" + app.Version + "; process bitness=" + (Environment.Is64BitProcess ? "x64" : "x86"));
            deck = app.Presentations.Add(Office.MsoTriState.msoTrue);
            deck.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank).Select();
            var host = new PowerPointHost(app); string status = "";
            host.Status = s => { status = s; Console.WriteLine(s); };
            string diagnostic = "";
            host.Diagnostic = s => { diagnostic = s; Console.WriteLine(s); };
            var sample = MetadataCodec.Sample();
            host.InsertEditableChart(sample);
            var shape = deck.Slides[1].Shapes[1];
            Require(shape.HasChart == Office.MsoTriState.msoTrue, "native chart");
            Require(shape.Chart.ChartType == Office.XlChartType.xlColumnClustered, "clustered column");
            Require(((PowerPoint.SeriesCollection)shape.Chart.SeriesCollection()).Count == 2, "two native series");
            Require(host.ReadChartMetadata(sample.ChartId).Categories.Count == 3, "three categories and attached metadata");
            // A native formatting edit must survive metadata reconciliation and reopen.
            shape.Chart.ChartStyle = 4;
            dynamic workbook;
            shape.Chart.ChartData.Activate(); workbook = shape.Chart.ChartData.Workbook;
            try
            {
                workbook.Worksheets[1].Cells[2, 2].Value2 = 42.0;
                Require(Convert.ToDouble(workbook.Worksheets[1].Cells[2, 2].Value2) == 42.0,
                    "embedded workbook accepts test edit");
                // The embedded workbook is persisted by its PowerPoint owner.
                // Standalone Workbook.Save can fail here; verify persistence by
                // reading the embedded workbook after presentation save/reopen.
            }
            finally
            {
                Console.WriteLine("STEP workbook.close");
                try
                {
                    PowerPointHost.RetryRejectedCall(() => workbook.Close(true), () => {
                        System.Windows.Forms.Application.DoEvents();
                        System.Threading.Thread.Sleep(100);
                    });
                }
                catch (System.Runtime.InteropServices.COMException ex) when (
                    ex.ErrorCode == unchecked((int)0x80010108) || ex.ErrorCode == unchecked((int)0x800706BA))
                {
                    // A disconnected Close has an ambiguous outcome. Do not replay
                    // it; independently reopen and verify the saved embedded data.
                    Console.WriteLine("STEP workbook.close-disconnected; verifying through a fresh workbook");
                    shape.Chart.ChartData.Activate();
                    dynamic freshWorkbook = shape.Chart.ChartData.Workbook;
                    Require(Convert.ToDouble(freshWorkbook.Worksheets[1].Cells[2, 2].Value2) == 42.0,
                        "fresh embedded workbook retains edit after disconnected close");
                }
            }
            Console.WriteLine("STEP chart.wait-for-edit");
            WaitForNativeEdit(shape.Chart);
            host.InspectAndRecover(true);
            Require(host.ReadChartMetadata(sample.ChartId).Series[0].Values[0] == 42, "native workbook edit captured");
            var file = Path.Combine(directory, "roundtrip.pptx");
            deck.SaveAs(file, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
            deck.Close(); deck = null;
            deck = app.Presentations.Open(file, Office.MsoTriState.msoFalse, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue);
            deck.Slides[1].Select();
            deck.Slides[1].Shapes[1].Chart.ChartData.Activate();
            dynamic persistedWorkbook = deck.Slides[1].Shapes[1].Chart.ChartData.Workbook;
            Require(Convert.ToDouble(persistedWorkbook.Worksheets[1].Cells[2, 2].Value2) == 42.0,
                "embedded workbook edit survives presentation save reopen");
            host.InspectAndRecover(false);
            Require(host.ReadChartMetadata(sample.ChartId).Series[0].Values[0] == 42, "save close reopen restores ID and data");
            Require(Convert.ToInt32(deck.Slides[1].Shapes[1].Chart.ChartStyle) == 4, "native formatting survives reopen");
            var duplicate = deck.Slides[1].Duplicate();
            host.InspectAndRecover(false);
            var copyId = duplicate[1].Shapes[1].Tags[PowerPointHost.IdTag];
            Require(!String.Equals(copyId, sample.ChartId, StringComparison.OrdinalIgnoreCase), "duplicate slide gets distinct ID");
            Require(host.ReadChartMetadata(copyId).Series[0].Values[0] == 42, "duplicate retains metadata");
            Require(host.ReadChartMetadata(sample.ChartId).Series[0].Values[0] == 42,
                "original metadata survives duplicate recovery");
            copyDeck = app.Presentations.Add(Office.MsoTriState.msoTrue);
            copyDeck.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank).Select();
            var transferredId = host.CopyEditableChartToActiveSlide(deck.Slides[1].Shapes[1]);
            host.InspectAndRecover(false);
            Require(host.ReadChartMetadata(transferredId).Series[0].Values[0] == 42, "explicit adapter copy to new deck carries XML");
            Require(copyDeck.Slides[1].Shapes[1].HasChart == Office.MsoTriState.msoTrue
                && Convert.ToInt32(copyDeck.Slides[1].Shapes[1].Chart.ChartStyle) == 4, "adapter copy preserves native chart and formatting");
            var copiedFile = Path.Combine(directory, "copied.pptx");
            copyDeck.SaveAs(copiedFile, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
            copyDeck.Close(); copyDeck = null;
            copyDeck = app.Presentations.Open(copiedFile); copyDeck.Slides[1].Select();
            Require(host.ReadChartMetadata(transferredId).Series[0].Values[0] == 42, "adapter copy metadata survives save reopen");
            copyDeck.Close(); copyDeck = null;
            deck.Windows[1].Activate(); deck.Slides[1].Select();
            Require(host.ReadChartMetadata(sample.ChartId).Series[0].Values[0] == 42,
                "source metadata preserved after cross-presentation transfer");
            // Corrupt only the test copy, keep the native chart intact.
            var damaged = duplicate[1].Shapes[1];
            var originalPartIds = new System.Collections.Generic.List<string>();
            foreach (Office.CustomXMLPart p in damaged.CustomerData)
                if (p.NamespaceURI == MetadataCodec.Namespace) originalPartIds.Add(p.Id);
            foreach (var partId in originalPartIds) damaged.CustomerData.Delete(partId);
            var corruptPart = damaged.CustomerData.Add();
            Require(corruptPart.LoadXML("<ChartDocument xmlns='" + MetadataCodec.Namespace + "'><SchemaVersion>999</SchemaVersion></ChartDocument>"),
                "corrupt metadata fixture created in fresh XML part");
            host.InspectAndRecover(true);
            Require(status.Contains("diagnostics: 1"), "corrupt metadata diagnostic");
            Require(diagnostic.Contains("operation=metadata.read") && diagnostic.Contains("HRESULT=0x")
                && !diagnostic.Contains("<ChartDocument"), "recovery diagnostic identifies operation without metadata");
            Require(damaged.HasChart == Office.MsoTriState.msoTrue, "corrupt metadata preserves native chart");
            var damagedPartIds = new System.Collections.Generic.List<string>();
            foreach (Office.CustomXMLPart p in damaged.CustomerData)
                if (p.NamespaceURI == MetadataCodec.Namespace) damagedPartIds.Add(p.Id);
            foreach (var partId in damagedPartIds) damaged.CustomerData.Delete(partId);
            host.InspectAndRecover(false);
            Require(status.Contains("diagnostics: 1") && damaged.HasChart == Office.MsoTriState.msoTrue, "missing metadata preserves native chart");
            damaged.Delete(); host.InspectAndRecover(false);
            Require(status.Contains("diagnostics: 0"), "deleted objects leave no stale registry");
            var edited = host.ReadChartMetadata(sample.ChartId); edited.Series[0].Values[0] = 55; edited.Revision++;
            host.UpdateEditableChart(sample.ChartId, edited);
            deck.Save(); deck.Close(); deck = null;
            deck = app.Presentations.Open(file); deck.Slides[1].Select();
            Require(host.ReadChartMetadata(sample.ChartId).Series[0].Values[0] == 55, "add-in edit survives second save reopen");
            Console.WriteLine("Adapter integration passed. VSTO/ribbon/task pane and full process restart NOT tested.");
            return 0;
        }
        catch (Exception ex) { testFailed = true; Console.Error.WriteLine(ex); return 1; }
        finally
        {
            try { CloseFixture(copyDeck); }
            catch (System.Runtime.InteropServices.COMException) when (testFailed) { }
            try { CloseFixture(deck); }
            catch (System.Runtime.InteropServices.COMException) when (testFailed) { }
        }
    }
}
