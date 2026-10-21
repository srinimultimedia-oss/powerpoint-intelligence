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
    public sealed class PowerPointHost : IPowerPointHost
    {
        public const string IdTag = "PPI_CHART_ID";
        private readonly PowerPoint.Application application;
        private readonly int threadId = Thread.CurrentThread.ManagedThreadId;
        public Action<string> Status = delegate { };

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
            // Add and verify before removing the previous part. No ZIP edits behind Office.
            var replacement = shape.CustomerData.Add();
            try
            {
                if (!replacement.LoadXML(xml)) throw new InvalidDataException("PowerPoint rejected chart metadata.");
                MetadataCodec.Decode(replacement.XML);
            }
            catch { replacement.Delete(); throw; }
            if (old != null) old.Delete();
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
                chart.Refresh();
            }
            finally { workbook.Close(true); }
        }

        private PowerPoint.Shape Find(string id)
        {
            var matches = Shapes().Where(s => String.Equals(s.Tags[IdTag], id, StringComparison.OrdinalIgnoreCase)).ToList();
            if (matches.Count != 1) throw new InvalidDataException("Chart absent or duplicated. Run Inspect / Recover first.");
            return matches[0];
        }
        public ChartDocument ReadChartMetadata(string chartId) { return Read(Find(chartId)); }
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
        private static ChartDocument Capture(PowerPoint.Shape shape, ChartDocument model)
        {
            dynamic chart = shape.Chart;
            if ((int)chart.ChartType != (int)Office.XlChartType.xlColumnClustered)
                throw new InvalidDataException("Native chart type changed; S0 metadata left intact.");
            dynamic series = chart.SeriesCollection();
            if ((int)series.Count != model.Series.Count)
                throw new InvalidDataException("Native series count changed; S0 metadata left intact.");
            for (int i = 1; i <= model.Series.Count; i++)
            {
                dynamic item = series.Item(i);
                var values = ((Array)item.Values).Cast<object>().ToArray();
                var categories = ((Array)item.XValues).Cast<object>().Select(x => Convert.ToString(x, CultureInfo.InvariantCulture)).ToList();
                if (values.Length != model.Categories.Count || categories.Count != model.Categories.Count)
                    throw new InvalidDataException("Native categories changed; S0 metadata left intact.");
                if (i == 1) model.Categories = categories;
                else if (!model.Categories.SequenceEqual(categories)) throw new InvalidDataException("Series categories disagree.");
                model.Series[i - 1].Name = Convert.ToString(item.Name, CultureInfo.InvariantCulture);
                model.Series[i - 1].Values = values.Select(x => x == null ? (double?)null : Convert.ToDouble(x, CultureInfo.InvariantCulture)).ToList();
            }
            MetadataCodec.Validate(model);
            return model;
        }

        public void InspectAndRecover(bool captureNativeEdits)
        {
            CheckThread();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int healthy = 0, problems = 0, copies = 0;
            foreach (var shape in Shapes())
            {
                try
                {
                    var part = Part(shape);
                    if (part == null && String.IsNullOrEmpty(shape.Tags[IdTag])) continue;
                    if (shape.HasChart != Office.MsoTriState.msoTrue) throw new InvalidDataException("Tagged object is no longer a chart.");
                    var model = Read(shape);
                    bool duplicate = !seen.Add(model.ChartId);
                    if (duplicate) { model.ChartId = Guid.NewGuid().ToString("D"); model.Revision = 0; }
                    if (captureNativeEdits)
                    {
                        var prior = MetadataCodec.Encode(model);
                        Capture(shape, model);
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
                { problems++; } // Do not log metadata/workbook contents or rewrite damaged objects.
            }
            Status(String.Format(CultureInfo.InvariantCulture, "Charts: {0}; copies assigned new IDs: {1}; diagnostics: {2}. Damaged/unsupported objects preserved.", healthy, copies, problems));
        }
    }
}
