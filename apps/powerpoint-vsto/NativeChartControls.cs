using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Office=Microsoft.Office.Core;
using PowerPoint=Microsoft.Office.Interop.PowerPoint;
namespace PowerPointIntelligence.AddIn
{
    internal sealed class NativeAxisSettings
    {
        internal bool Automatic=true,Logarithmic,Reversed;
        internal double Minimum,Maximum=100,Major=10;
        internal string NumberFormat="General";
        internal void Validate()
        {
            if(NumberFormat==null||NumberFormat.Length==0||NumberFormat.Length>100||NumberFormat.Any(Char.IsControl))throw new InvalidDataException("Use a number format of 1–100 printable characters.");
            if(!Finite(Minimum)||!Finite(Maximum)||!Finite(Major))throw new InvalidDataException("Axis values must be finite.");
            if(!Automatic&&(Minimum>=Maximum||Major<=0||!Finite(Maximum-Minimum)))throw new InvalidDataException("Minimum must be smaller than maximum and major interval must be positive.");
            if(!Automatic&&Logarithmic&&Minimum<=0)throw new InvalidDataException("A logarithmic axis needs a positive minimum.");
        }
        internal static bool Finite(double n){return !Double.IsNaN(n)&&!Double.IsInfinity(n);}
        internal static void ValidateLogValues(System.Collections.Generic.IEnumerable<object> values)
        {
            int count=0,numeric=0;foreach(var v in values){if(++count>100000)throw new InvalidDataException("The logarithmic preflight exceeds 100000 values.");if(v==null||v==DBNull.Value||v is string&&String.IsNullOrWhiteSpace((string)v))continue;double number;try{number=Convert.ToDouble(v,CultureInfo.InvariantCulture);}catch(Exception ex) when(ex is FormatException||ex is InvalidCastException||ex is OverflowException){throw new InvalidDataException("Logarithmic scale requires numeric values.",ex);}numeric++;if(!Finite(number)||number<=0)throw new InvalidDataException("Logarithmic scale requires positive values in every primary series.");}
            if(numeric==0)throw new InvalidDataException("The chart has no values to validate.");
        }
    }
    internal static class NativeChartControls
    {
        internal static PowerPoint.Shape Selected(PowerPoint.Application app)
        {var selection=app.ActiveWindow.Selection;if(selection.Type!=PowerPoint.PpSelectionType.ppSelectionShapes||selection.ShapeRange.Count!=1||selection.ShapeRange[1].HasChart!=Office.MsoTriState.msoTrue)throw new InvalidOperationException("Select one native chart first.");return selection.ShapeRange[1];}
        private static PowerPoint.Axis Axis(PowerPoint.Chart chart)
        {if(!Convert.ToBoolean(chart.HasAxis[PowerPoint.XlAxisType.xlValue,PowerPoint.XlAxisGroup.xlPrimary],CultureInfo.InvariantCulture))throw new InvalidOperationException("This chart has no primary value axis.");return (PowerPoint.Axis)chart.Axes(PowerPoint.XlAxisType.xlValue,PowerPoint.XlAxisGroup.xlPrimary);}
        private sealed class Snapshot
        {
            internal NativeAxisSettings Settings;internal bool MinAuto,MaxAuto,MajorAuto;
            internal Snapshot(PowerPoint.Axis axis){MinAuto=axis.MinimumScaleIsAuto;MaxAuto=axis.MaximumScaleIsAuto;MajorAuto=axis.MajorUnitIsAuto;Settings=new NativeAxisSettings {Automatic=MinAuto&&MaxAuto&&MajorAuto,Minimum=axis.MinimumScale,Maximum=axis.MaximumScale,Major=axis.MajorUnit,Logarithmic=axis.ScaleType==PowerPoint.XlScaleType.xlScaleLogarithmic,Reversed=axis.ReversePlotOrder,NumberFormat=axis.TickLabels.NumberFormat};}
            internal void Restore(PowerPoint.Axis axis){Apply(axis,Settings);axis.MinimumScaleIsAuto=MinAuto;axis.MaximumScaleIsAuto=MaxAuto;axis.MajorUnitIsAuto=MajorAuto;}
        }
        private static void Apply(PowerPoint.Axis axis,NativeAxisSettings settings)
        {
            // Reset bounds before setting a new range, to avoid a stale opposite bound rejecting it.
            axis.MinimumScaleIsAuto=true;axis.MaximumScaleIsAuto=true;axis.MajorUnitIsAuto=true;
            axis.ScaleType=settings.Logarithmic?PowerPoint.XlScaleType.xlScaleLogarithmic:PowerPoint.XlScaleType.xlScaleLinear;
            if(!settings.Automatic){axis.MaximumScale=settings.Maximum;axis.MinimumScale=settings.Minimum;if(!settings.Logarithmic)axis.MajorUnit=settings.Major;}
            axis.ReversePlotOrder=settings.Reversed;axis.TickLabels.NumberFormat=settings.NumberFormat;
        }
        internal static void SetAxis(PowerPoint.Application app,NativeAxisSettings settings)
        {
            settings.Validate();var shape=Selected(app);var chart=shape.Chart;var axis=Axis(chart);var old=new Snapshot(axis);
            if(settings.Logarithmic){var series=(PowerPoint.SeriesCollection)chart.SeriesCollection();var values=new System.Collections.Generic.List<object>();for(int seriesIndex=1;seriesIndex<=series.Count;seriesIndex++){PowerPoint.Series item=series.Item(seriesIndex);if(item.AxisGroup!=PowerPoint.XlAxisGroup.xlPrimary)continue;object raw=item.Values;var array=raw as Array;if(array!=null&&array.Rank!=1)throw new InvalidDataException("Series values must be one-dimensional.");if(array==null)values.Add(raw);else values.AddRange(array.Cast<object>());}NativeAxisSettings.ValidateLogValues(values);}
            app.StartNewUndoEntry();try{Apply(axis,settings);chart.Refresh();}catch(Exception error){try{old.Restore(axis);}catch(Exception restore){throw new InvalidOperationException("Axis update failed and restoration was incomplete. Use PowerPoint Undo.",new AggregateException(error,restore));}throw new InvalidOperationException("PowerPoint rejected the axis settings; previous settings were restored.",error);}
        }
        internal static void ShowAxis(PowerPoint.Application app)
        {
            var current=new Snapshot(Axis(Selected(app).Chart)).Settings;
            using(var form=new Form {Text="Value axis · Intelligence",Width=430,Height=370,FormBorderStyle=FormBorderStyle.FixedDialog,StartPosition=FormStartPosition.CenterParent,MinimizeBox=false,MaximizeBox=false})
            {
                var auto=new CheckBox {Text="Automatic bounds and interval",Checked=current.Automatic,Left=20,Top=20,Width=350};var log=new CheckBox {Text="Logarithmic (base 10)",Checked=current.Logarithmic,Left=20,Top=52,Width=350};var reverse=new CheckBox {Text="Reverse values",Checked=current.Reversed,Left=20,Top=84,Width=350};
                var labels=new[]{"Minimum","Maximum","Major interval","Number format"};var boxes=new TextBox[4];var initial=new[]{current.Minimum.ToString("R",CultureInfo.InvariantCulture),current.Maximum.ToString("R",CultureInfo.InvariantCulture),current.Major.ToString("R",CultureInfo.InvariantCulture),current.NumberFormat};
                for(int i=0;i<4;i++){form.Controls.Add(new Label {Text=labels[i],Left=20,Top=120+i*32,Width=130});boxes[i]=new TextBox {Left=155,Top=117+i*32,Width=235,Text=initial[i],MaxLength=100};form.Controls.Add(boxes[i]);}
                Action enabled=()=>{boxes[0].Enabled=boxes[1].Enabled=!auto.Checked;boxes[2].Enabled=!auto.Checked&&!log.Checked;};auto.CheckedChanged+=(s,e)=>enabled();log.CheckedChanged+=(s,e)=>enabled();enabled();
                var apply=new Button {Text="Apply",Left=220,Top=270,Width=80};var cancel=new Button {Text="Cancel",Left=310,Top=270,Width=80,DialogResult=DialogResult.Cancel};form.Controls.AddRange(new Control[]{auto,log,reverse,apply,cancel});form.AcceptButton=apply;form.CancelButton=cancel;
                NativeAxisSettings choice=null;apply.Click+=(s,e)=>{try{double min=current.Minimum,max=current.Maximum,major=current.Major;if(!auto.Checked&&(!Double.TryParse(boxes[0].Text,NumberStyles.Float,CultureInfo.InvariantCulture,out min)||!Double.TryParse(boxes[1].Text,NumberStyles.Float,CultureInfo.InvariantCulture,out max)||!log.Checked&&!Double.TryParse(boxes[2].Text,NumberStyles.Float,CultureInfo.InvariantCulture,out major)))throw new InvalidDataException("Use numbers with a decimal point.");var settings=new NativeAxisSettings {Automatic=auto.Checked,Logarithmic=log.Checked,Reversed=reverse.Checked,Minimum=min,Maximum=max,Major=major,NumberFormat=boxes[3].Text};settings.Validate();choice=settings;form.DialogResult=DialogResult.OK;}catch(InvalidDataException error){MessageBox.Show(form,error.Message,"Axis settings");}};
                if(form.ShowDialog()==DialogResult.OK)SetAxis(app,choice);
            }
        }
        internal static void MatchScales(PowerPoint.Application app)
        {
            var selection=app.ActiveWindow.Selection;if(selection.Type!=PowerPoint.PpSelectionType.ppSelectionShapes||selection.ShapeRange.Count<2||selection.ShapeRange.Count>30)throw new InvalidOperationException("Select 2–30 native charts on this slide.");
            var charts=selection.ShapeRange.Cast<PowerPoint.Shape>().Select(s=>{if(s.HasChart!=Office.MsoTriState.msoTrue)throw new InvalidOperationException("Every selected object must be a native chart.");return s.Chart;}).ToArray();var axes=charts.Select(Axis).ToArray();var old=axes.Select(a=>new Snapshot(a)).ToArray();bool log=old[0].Settings.Logarithmic;if(old.Any(s=>s.Settings.Logarithmic!=log))throw new InvalidOperationException("Choose axes with the same linear/logarithmic scale type.");
            var target=new NativeAxisSettings {Automatic=false,Logarithmic=log,Minimum=old.Min(s=>s.Settings.Minimum),Maximum=old.Max(s=>s.Settings.Maximum),Major=old.Max(s=>s.Settings.Major)};target.Validate();app.StartNewUndoEntry();int touched=0;
            try{for(int i=0;i<axes.Length;i++){touched=i+1;var choice=new NativeAxisSettings {Automatic=false,Logarithmic=log,Minimum=target.Minimum,Maximum=target.Maximum,Major=target.Major,Reversed=old[i].Settings.Reversed,NumberFormat=old[i].Settings.NumberFormat};Apply(axes[i],choice);}}
            catch(Exception error){var failures=new System.Collections.Generic.List<Exception>{error};for(int i=0;i<touched;i++)try{old[i].Restore(axes[i]);}catch(Exception ex){failures.Add(ex);}throw new InvalidOperationException(failures.Count==1?"Scale matching failed; previous settings restored.":"Scale matching failed; restoration incomplete. Use PowerPoint Undo.",new AggregateException(failures));}
        }
    }
}
