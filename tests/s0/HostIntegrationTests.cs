using System;
using System.IO;
using System.Linq;
using PowerPointIntelligence.Host;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

internal static class HostIntegrationTests
{
    private static void Require(bool condition, string name)
    { if (!condition) throw new Exception("FAIL " + name); Console.WriteLine("PASS " + name); }

    // Explicit opt-in only. Runs visibly on an interactive workstation, never in CI.
    // Does not quit PowerPoint or close presentations owned by the user.
    [STAThread]
    public static int Main(string[] args)
    {
        if (args.Length != 2 || args[0] != "--interactive-office")
        { Console.Error.WriteLine("Usage: HostIntegrationTests --interactive-office <new-output-directory>"); return 2; }
        var directory = Path.GetFullPath(args[1]);
        if (Directory.Exists(directory)) throw new IOException("Use a new fixture directory.");
        Directory.CreateDirectory(directory);
        PowerPoint.Presentation deck = null, copyDeck = null;
        try
        {
            var app = new PowerPoint.Application();
            app.Visible = Office.MsoTriState.msoTrue;
            Console.WriteLine("Office version=" + app.Version + "; process bitness=" + (Environment.Is64BitProcess ? "x64" : "x86"));
            deck = app.Presentations.Add(Office.MsoTriState.msoTrue);
            deck.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank).Select();
            var host = new PowerPointHost(app); string status = "";
            host.Status = s => { status = s; Console.WriteLine(s); };
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
            workbook.Worksheets[1].Cells[2, 2].Value2 = 42.0;
            workbook.Close(true); shape.Chart.Refresh();
            host.InspectAndRecover(true);
            Require(host.ReadChartMetadata(sample.ChartId).Series[0].Values[0] == 42, "native workbook edit captured");
            var file = Path.Combine(directory, "roundtrip.pptx");
            deck.SaveAs(file, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
            deck.Close(); deck = null;
            deck = app.Presentations.Open(file, Office.MsoTriState.msoFalse, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue);
            deck.Slides[1].Select();
            host.InspectAndRecover(false);
            Require(host.ReadChartMetadata(sample.ChartId).Series[0].Values[0] == 42, "save close reopen restores ID and data");
            Require(Convert.ToInt32(deck.Slides[1].Shapes[1].Chart.ChartStyle) == 4, "native formatting survives reopen");
            var duplicate = deck.Slides[1].Duplicate();
            host.InspectAndRecover(false);
            var copyId = duplicate[1].Shapes[1].Tags[PowerPointHost.IdTag];
            Require(!String.Equals(copyId, sample.ChartId, StringComparison.OrdinalIgnoreCase), "duplicate slide gets distinct ID");
            Require(host.ReadChartMetadata(copyId).Series[0].Values[0] == 42, "duplicate retains metadata");
            copyDeck = app.Presentations.Add(Office.MsoTriState.msoTrue);
            copyDeck.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank).Select();
            deck.Slides[1].Shapes[1].Copy(); copyDeck.Slides[1].Shapes.Paste();
            host.InspectAndRecover(false);
            Require(host.ReadChartMetadata(sample.ChartId).Series[0].Values[0] == 42, "copy to new deck carries XML");
            copyDeck.SaveAs(Path.Combine(directory, "copied.pptx"), PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
            copyDeck.Close(); copyDeck = null;
            deck.Windows[1].Activate(); deck.Slides[1].Select();
            // Corrupt only the test copy, keep the native chart intact.
            var damaged = duplicate[1].Shapes[1];
            foreach (Office.CustomXMLPart p in damaged.CustomerData)
                if (p.NamespaceURI == MetadataCodec.Namespace) p.LoadXML("<ChartDocument xmlns='" + MetadataCodec.Namespace + "'><SchemaVersion>999</SchemaVersion></ChartDocument>");
            host.InspectAndRecover(true);
            Require(status.Contains("diagnostics: 1"), "corrupt metadata diagnostic");
            Require(damaged.HasChart == Office.MsoTriState.msoTrue, "corrupt metadata preserves native chart");
            foreach (Office.CustomXMLPart p in damaged.CustomerData)
                if (p.NamespaceURI == MetadataCodec.Namespace) p.Delete();
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
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
        finally
        {
            if (copyDeck != null) { copyDeck.Saved = Office.MsoTriState.msoTrue; copyDeck.Close(); }
            if (deck != null) { deck.Saved = Office.MsoTriState.msoTrue; deck.Close(); }
        }
    }
}
