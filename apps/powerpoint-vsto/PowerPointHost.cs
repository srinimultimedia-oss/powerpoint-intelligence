using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointIntelligence.Host
{
    // Serializes this add-in's writes with explicitly coordinated desktop automation.
    // The kernel releases ownership if the automation process exits unexpectedly.
    internal sealed class OfficeMutationLease : IDisposable
    {
        private Mutex mutex;
        private OfficeMutationLease(Mutex value){mutex=value;}
        [System.Runtime.InteropServices.DllImport("user32.dll",ExactSpelling=true)]
        private static extern uint GetWindowThreadProcessId(IntPtr window,out uint processId);
        internal static OfficeMutationLease TryEnter(int processId,int timeoutMilliseconds=0)
        {
            if(processId<=0||timeoutMilliseconds<0||timeoutMilliseconds>10000)throw new ArgumentOutOfRangeException("processId/timeout");
            var gate=new Mutex(false,"Local\\PowerPointIntelligence.OfficeMutation.v1."+processId.ToString(CultureInfo.InvariantCulture));bool entered=false;
            try{try{entered=gate.WaitOne(timeoutMilliseconds);}catch(AbandonedMutexException){entered=true;}if(!entered){gate.Dispose();return null;}return new OfficeMutationLease(gate);}
            catch{gate.Dispose();throw;}
        }
        internal static OfficeMutationLease ForAutomation(PowerPoint.Application app)
        {
            uint processId;GetWindowThreadProcessId(new IntPtr(app.HWND),out processId);
            var lease=TryEnter(checked((int)processId),10000);if(lease==null)throw new InvalidOperationException("PowerPoint add-in mutation is busy; automation did not start.");return lease;
        }
        public void Dispose(){var gate=mutex;if(gate==null)return;mutex=null;try{gate.ReleaseMutex();}finally{gate.Dispose();}}
    }
    internal static class OfficeEventBoundary
    {
        // Exceptions must not escape callbacks invoked by the native Office host.
        internal static bool Run(Action action, Action<string> diagnostic)
        {
            try { action(); return true; }
            catch (Exception ex)
            {
                diagnostic("Office event failed: " + ex.GetType().Name + "; HRESULT=0x" + ex.HResult.ToString("X8", CultureInfo.InvariantCulture));
                return false;
            }
        }
    }

    public sealed class PowerPointHost : IPowerPointHost
    {
        public const string IdTag = "PPI_CHART_ID";
        private readonly PowerPoint.Application application;
        private readonly int threadId = Thread.CurrentThread.ManagedThreadId;
        public Action<string> Status = delegate { };
        // Fixed operation names and error codes only; never expose workbook/metadata text.
        public Action<string> Diagnostic = delegate { };

        public PowerPointHost(PowerPoint.Application application) { this.application = application; }
        private void CheckThread()
        {
            if (Thread.CurrentThread.ManagedThreadId != threadId || Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                throw new InvalidOperationException("PowerPoint access requires the add-in UI STA.");
        }
        private PowerPoint.Presentation Presentation { get { CheckThread(); return application.ActivePresentation; } }
        private PowerPoint.Slide Slide { get { CheckThread(); return (PowerPoint.Slide)application.ActiveWindow.View.Slide; } }
        public string GetSelectedSlideId() { return Slide.SlideID.ToString(CultureInfo.InvariantCulture); }
        public IReadOnlyList<string> GetSelectedObjectIds()
        {
            CheckThread();
            var result = new List<string>();
            var selection = application.ActiveWindow.Selection;
            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                for (int i = 1; i <= selection.ShapeRange.Count; i++)
                {
                    var id = selection.ShapeRange[i].Tags[IdTag];
                    if (!String.IsNullOrEmpty(id)) result.Add(id);
                }
            return result;
        }

        private IEnumerable<PowerPoint.Shape> Shapes()
        {
            var presentation = Presentation;
            for (int s = 1; s <= presentation.Slides.Count; s++)
                for (int i = 1; i <= presentation.Slides[s].Shapes.Count; i++)
                    yield return presentation.Slides[s].Shapes[i];
        }

        // CustomerData binds the XML part to the shape. Office copies it with the shape,
        // including between presentations; shape IDs are never canonical identities.
        private Office.CustomXMLPart Part(PowerPoint.Shape shape)
        {
            Office.CustomXMLPart found = null;
            var data = shape.CustomerData;
            foreach (Office.CustomXMLPart part in data)
            {
                if (part != null && part.NamespaceURI == MetadataCodec.Namespace)
                {
                    if (found != null) throw new InvalidDataException("Multiple chart metadata parts; preserved for recovery.");
                    found = part;
                }
            }
            return found;
        }

        private ChartDocument Read(PowerPoint.Shape shape)
        {
            var part = Part(shape);
            if (part == null) throw new InvalidDataException("Chart metadata missing; native chart preserved.");
            var document = MetadataCodec.Decode(part.XML);
            var tag = shape.Tags[IdTag];
            if (!String.IsNullOrEmpty(tag) && !String.Equals(tag, document.ChartId, StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("Chart tag and metadata disagree; native chart preserved.");
            return document;
        }

        private void Write(PowerPoint.Shape shape, ChartDocument document)
        {
            var xml = MetadataCodec.Encode(document);
            var old = Part(shape);
            var oldId = old == null ? null : old.Id;
            var otherOwners = new List<KeyValuePair<PowerPoint.Shape, string>>();
            if (oldId != null)
            {
                var targetSlide = (PowerPoint.Slide)shape.Parent;
                foreach (var candidate in Shapes())
                {
                    if (candidate.Id == shape.Id && ((PowerPoint.Slide)candidate.Parent).SlideID == targetSlide.SlideID) continue;
                    var existing = Part(candidate);
                    if (existing != null && existing.Id == oldId)
                    {
                        var existingXml = existing.XML;
                        MetadataCodec.Decode(existingXml);
                        otherOwners.Add(new KeyValuePair<PowerPoint.Shape, string>(candidate, existingXml));
                    }
                }
            }
            // Add and verify before removing the previous part. No ZIP edits behind Office.
            var replacement = shape.CustomerData.Add();
            try
            {
                if (!replacement.LoadXML(xml)) throw new InvalidDataException("PowerPoint rejected chart metadata.");
                MetadataCodec.Decode(replacement.XML);
            }
            catch { shape.CustomerData.Delete(replacement.Id); throw; }
            // Remove through the owning shape's collection, not the document-wide
            // XML store. Resolve the ID before adding/mutating collection entries.
            if (oldId != null) shape.CustomerData.Delete(oldId);
            // Slide duplication can leave multiple owners referring to the same
            // data-store part. Preserve each validated owner's XML independently
            // when replacing that shared part; never reconstruct it from a tag.
            foreach (var owner in otherOwners)
            {
                var existing = Part(owner.Key);
                if (existing == null)
                {
                    var restored = owner.Key.CustomerData.Add();
                    if (!restored.LoadXML(owner.Value) || restored.XML != owner.Value)
                        throw new InvalidDataException("Shared metadata owner restoration failed.");
                }
                else if (existing.XML != owner.Value)
                    throw new InvalidDataException("Shared metadata owner changed during replacement.");
            }
            shape.Tags.Add(IdTag, document.ChartId);
        }

        public IReadOnlyList<string> InsertEditableChart(ChartDocument chart)
        {
            CheckThread(); MetadataCodec.Validate(chart);
            if (Shapes().Any(s => String.Equals(s.Tags[IdTag], chart.ChartId, StringComparison.OrdinalIgnoreCase)))
                throw new InvalidOperationException("Chart ID already exists.");
            application.StartNewUndoEntry();
            var shape = Slide.Shapes.AddChart2(-1, Office.XlChartType.xlColumnClustered, 60, 90, 600, 340, false);
            try
            {
                SetData(shape, chart);
                shape.AlternativeText = "Editable clustered column chart: " + String.Join(", ", chart.Categories);
                Write(shape, chart);
                shape.Select(Office.MsoTriState.msoTrue);
                Status("Inserted native editable chart " + chart.ChartId);
                return new List<string> { chart.ChartId };
            }
            catch { shape.Delete(); throw; }
        }

        private static void SetData(PowerPoint.Shape shape, ChartDocument model)
        {
            var chart = shape.Chart;
            chart.ChartData.Activate();
            dynamic workbook = chart.ChartData.Workbook;
            try
            {
                dynamic sheet = workbook.Worksheets[1];
                // Write labels as literal text (never Excel formulas), numbers as numbers.
                sheet.Cells.Clear();
                sheet.Cells[1, 1].NumberFormat = "@";
                sheet.Cells[1, 1].Value2 = "Category";
                for (int r = 0; r < model.Categories.Count; r++)
                {
                    sheet.Cells[r + 2, 1].NumberFormat = "@";
                    sheet.Cells[r + 2, 1].Value2 = model.Categories[r];
                }
                for (int c = 0; c < model.Series.Count; c++)
                {
                    sheet.Cells[1, c + 2].NumberFormat = "@";
                    sheet.Cells[1, c + 2].Value2 = model.Series[c].Name;
                    for (int r = 0; r < model.Categories.Count; r++)
                        sheet.Cells[r + 2, c + 2].Value2 = model.Series[c].Values[r];
                }
                dynamic range = sheet.Range[sheet.Cells[1, 1], sheet.Cells[model.Categories.Count + 1, model.Series.Count + 1]];
                chart.SetSourceData(range.Address[true, true, 1, true], PowerPoint.XlRowCol.xlColumns);
                chart.ChartType = Office.XlChartType.xlColumnClustered;
                chart.HasLegend = true;
            }
            finally { RetryRejectedCall(() => workbook.Close(true)); }
            // Finish the embedded Excel transaction before refreshing its host chart.
            // Avoid refreshing the host while still relying on the workbook proxy.
            chart.Refresh();
        }

        private PowerPoint.Shape Find(string id)
        {
            var matches = Shapes().Where(s => String.Equals(s.Tags[IdTag], id, StringComparison.OrdinalIgnoreCase)).ToList();
            if (matches.Count != 1) throw new InvalidDataException("Chart absent or duplicated. Run Inspect / Recover first.");
            return matches[0];
        }
        public ChartDocument ReadChartMetadata(string chartId) { return Read(Find(chartId)); }

        // Explicit adapter transfer: PowerPoint's native clipboard can omit CustomerData
        // across presentations. Read the source before copying; never infer missing XML
        // from an ID tag or another open document during ordinary recovery.
        public string CopyEditableChartToActiveSlide(PowerPoint.Shape source)
        {
            CheckThread();
            var destination = Slide;
            var destinationWindow = application.ActiveWindow;
            if (source.HasChart != Office.MsoTriState.msoTrue)
                throw new InvalidDataException("Source is not a native chart.");
            ChartDocument model;
            try
            {
                // CustomerData access through an inactive presentation's shape can
                // fail even when its XML was readable before switching documents.
                var sourceSlide = (PowerPoint.Slide)source.Parent;
                var sourcePresentation = (PowerPoint.Presentation)sourceSlide.Parent;
                sourcePresentation.Windows[1].Activate();
                sourceSlide.Select();
                model = Read(source);
                string operation = "copy.capture";
                Capture(source, model, ref operation);
                source.Copy();
            }
            finally
            {
                destinationWindow.Activate();
                destination.Select();
            }
            if (Shapes().Any(s => String.Equals(s.Tags[IdTag], model.ChartId, StringComparison.OrdinalIgnoreCase)))
            {
                model.ChartId = Guid.NewGuid().ToString("D");
                model.Revision = 0;
            }
            application.StartNewUndoEntry();
            var pasted = destination.Shapes.Paste();
            try
            {
                if (pasted.Count != 1 || pasted[1].HasChart != Office.MsoTriState.msoTrue)
                    throw new InvalidDataException("Clipboard did not produce one native chart.");
                Write(pasted[1], model);
                if (MetadataCodec.Encode(Read(pasted[1])) != MetadataCodec.Encode(model))
                    throw new InvalidDataException("Copied chart metadata verification failed.");
                return model.ChartId;
            }
            catch { pasted.Delete(); throw; }
        }
        public void SaveChartMetadata(ChartDocument chart)
        {
            MetadataCodec.Validate(chart);
            var shape = Find(chart.ChartId);
            Read(shape); // Never overwrite corrupt or unknown metadata.
            Write(shape, chart);
        }
        public void UpdateEditableChart(string chartId, ChartDocument chart)
        {
            MetadataCodec.Validate(chart);
            if (!String.Equals(chartId, chart.ChartId, StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("Chart identity cannot change during update.");
            var shape = Find(chartId);
            var before = Read(shape);
            application.StartNewUndoEntry();
            try { SetData(shape, chart); Write(shape, chart); }
            catch { SetData(shape, before); throw; }
        }

        // Native chart edits are read into metadata, never projected back on open/save.
        // Unsupported structural edits fail safely; existing native data/formatting remain intact.
        private static ChartDocument Capture(PowerPoint.Shape shape, ChartDocument model, ref string operation)
        {
            operation = "capture.chart-type";
            PowerPoint.Chart chart = shape.Chart;
            if ((int)chart.ChartType != (int)Office.XlChartType.xlColumnClustered)
                throw new InvalidDataException("Native chart type changed; S0 metadata left intact.");
            operation = "capture.series-count";
            var series = (PowerPoint.SeriesCollection)chart.SeriesCollection();
            if (series.Count != model.Series.Count)
                throw new InvalidDataException("Native series count changed; S0 metadata left intact.");
            for (int i = 1; i <= model.Series.Count; i++)
            {
                operation = "capture.series-item";
                PowerPoint.Series item = series.Item(i);
                operation = "capture.series-values";
                var values = SeriesVector((object)item.Values);
                operation = "capture.series-categories";
                var categories = SeriesVector((object)item.XValues).Select(x => Convert.ToString(x, CultureInfo.InvariantCulture)).ToList();
                operation = "capture.category-validation";
                if (values.Length != model.Categories.Count || categories.Count != model.Categories.Count)
                    throw new InvalidDataException("Native categories changed; S0 metadata left intact.");
                if (i == 1) model.Categories = categories;
                else if (!model.Categories.SequenceEqual(categories)) throw new InvalidDataException("Series categories disagree.");
                operation = "capture.series-name";
                model.Series[i - 1].Name = Convert.ToString(item.Name, CultureInfo.InvariantCulture);
                operation = "capture.value-conversion";
                model.Series[i - 1].Values = values.Select(x => x == null ? (double?)null : Convert.ToDouble(x, CultureInfo.InvariantCulture)).ToList();
            }
            operation = "capture.metadata-validation";
            MetadataCodec.Validate(model);
            return model;
        }

        // COM VARIANTs can contain typed SAFEARRAYs with nonzero lower bounds,
        // or a scalar for a single point. Do not dynamically cast a COM result.
        internal static object[] NativeSeriesValues(PowerPoint.SeriesCollection collection,int index)
        {
            PowerPoint.Series item=collection.Item(index);object raw=item.Values;return SeriesVector(raw);
        }
        internal static object[] SeriesVector(object value)
        {
            var array = value as Array;
            if (array == null) return new object[] { value };
            if (array.Rank != 1) throw new InvalidDataException("Native series must be one-dimensional.");
            return array.Cast<object>().ToArray();
        }

        // Only replay a call which COM explicitly rejected before execution.
        // Never replay mutations after disconnection or an unknown failure.
        internal static void RetryRejectedCall(Action action, Action pause = null, int maxAttempts = 51)
        {
            if (maxAttempts < 1) throw new ArgumentOutOfRangeException("maxAttempts");
            for (int attempt = 1; ; attempt++)
            {
                try { action(); return; }
                catch (System.Runtime.InteropServices.COMException ex)
                    when (ex.ErrorCode == unchecked((int)0x80010001) && attempt < maxAttempts)
                {
                    if (pause != null) pause();
                    else Thread.Sleep(100);
                }
            }
        }

        public void InspectAndRecover(bool captureNativeEdits)
        {
            CheckThread();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int healthy = 0, problems = 0, copies = 0;
            foreach (var shape in Shapes())
            {
                string operation = "metadata.locate";
                try
                {
                    var part = Part(shape);
                    if (part == null && String.IsNullOrEmpty(shape.Tags[IdTag])) continue;
                    if (shape.HasChart != Office.MsoTriState.msoTrue) throw new InvalidDataException("Tagged object is no longer a chart.");
                    operation = "metadata.read";
                    var model = Read(shape);
                    bool duplicate = !seen.Add(model.ChartId);
                    if (duplicate) { model.ChartId = Guid.NewGuid().ToString("D"); model.Revision = 0; }
                    if (captureNativeEdits)
                    {
                        var prior = MetadataCodec.Encode(model);
                        Capture(shape, model, ref operation);
                        operation = "metadata.write";
                        if (prior != MetadataCodec.Encode(model)) model.Revision = checked(model.Revision + 1);
                        if (duplicate || prior != MetadataCodec.Encode(model)) Write(shape, model);
                    }
                    else if (duplicate) Write(shape, model);
                    if (String.IsNullOrEmpty(shape.Tags[IdTag])) shape.Tags.Add(IdTag, model.ChartId);
                    seen.Add(model.ChartId);
                    if (duplicate) copies++;
                    healthy++;
                }
                catch (Exception ex) when (ex is InvalidDataException || ex is InvalidOperationException || ex is System.Runtime.InteropServices.COMException || ex is FormatException || ex is OverflowException || ex is InvalidCastException)
                {
                    problems++;
                    Diagnostic(String.Format(CultureInfo.InvariantCulture,
                        "Chart recovery failed: operation={0}; exception={1}; HRESULT=0x{2:X8}",
                        operation, ex.GetType().Name, ex.HResult));
                } // Do not log metadata/workbook contents or rewrite damaged objects.
            }
            Status(String.Format(CultureInfo.InvariantCulture, "Charts: {0}; copies assigned new IDs: {1}; diagnostics: {2}. Damaged/unsupported objects preserved.", healthy, copies, problems));
        }
    }
}
