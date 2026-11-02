using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using PowerPoint=Microsoft.Office.Interop.PowerPoint;
using Office=Microsoft.Office.Core;
namespace PowerPointIntelligence.AddIn {
 internal static class ContextChartFontPlan {
  internal static string Family(object value){var text=value as string;return String.IsNullOrWhiteSpace(text)||text.Length>128?null:text;}
  internal static float? Size(object value){if(value==null||value==DBNull.Value||value is bool)return null;try{float size=Convert.ToSingle(value,CultureInfo.InvariantCulture);return Single.IsNaN(size)||Single.IsInfinity(size)||size<1||size>4000?null:(float?)size;}catch(FormatException){return null;}catch(InvalidCastException){return null;}catch(OverflowException){return null;}}
  internal static bool? Flag(object value){if(value==null||value==DBNull.Value)return null;if(value is bool)return (bool)value;try{double raw=Convert.ToDouble(value,CultureInfo.InvariantCulture);if(raw!=Math.Truncate(raw)||Double.IsNaN(raw)||Double.IsInfinity(raw))return null;int flag=Convert.ToInt32(raw,CultureInfo.InvariantCulture);if(flag==0)return false;if(flag==-1||flag==1)return true;return null;}catch(FormatException){return null;}catch(InvalidCastException){return null;}catch(OverflowException){return null;}}
  internal static int? Color(object value){if(value==null||value==DBNull.Value||value is bool)return null;try{double n=Convert.ToDouble(value,CultureInfo.InvariantCulture);return Double.IsNaN(n)||Double.IsInfinity(n)||n<0||n>0xFFFFFF||n!=Math.Truncate(n)?null:(int?)n;}catch(FormatException){return null;}catch(InvalidCastException){return null;}catch(OverflowException){return null;}}
 }
 internal static class ContextChartFormatting {
  static PowerPoint.ChartFont Font(PowerPoint.Application app){if(app==null||app.ActiveWindow==null||app.ActiveWindow.ViewType!=PowerPoint.PpViewType.ppViewNormal)throw new InvalidOperationException("Select one native chart in Normal view.");var selection=app.ActiveWindow.Selection;if(selection.Type!=PowerPoint.PpSelectionType.ppSelectionShapes||selection.ShapeRange.Count!=1||selection.ShapeRange[1].HasChart!=Office.MsoTriState.msoTrue)throw new InvalidOperationException("Select one native chart. Font changes affect its chart-wide labels.");return selection.ShapeRange[1].Chart.ChartArea.Font;}
  internal static FontSelectionState Read(PowerPoint.Application app){var font=Font(app);return new FontSelectionState{Family=ContextChartFontPlan.Family(font.Name),Size=ContextChartFontPlan.Size(font.Size),Bold=ContextChartFontPlan.Flag(font.Bold),Italic=ContextChartFontPlan.Flag(font.Italic),TargetCount=1,CharacterCount=0};}
  internal static void ApplyFont(PowerPoint.Application app,string family,float? size,bool? bold,bool? italic){ContextFontPlan.Validate(family,size,bold,italic);var font=Font(app);var changes=new List<NativeFormattingChange>();if(family!=null){object old=font.Name;if(ContextChartFontPlan.Family(old)==null)throw Mixed();changes.Add(new NativeFormattingChange(()=>font.Name=family,()=>font.Name=old));}if(size.HasValue){object old=font.Size;if(!ContextChartFontPlan.Size(old).HasValue)throw Mixed();changes.Add(new NativeFormattingChange(()=>font.Size=size.Value,()=>font.Size=old));}if(bold.HasValue){object old=font.Bold;if(!ContextChartFontPlan.Flag(old).HasValue)throw Mixed();changes.Add(new NativeFormattingChange(()=>font.Bold=bold.Value,()=>font.Bold=old));}if(italic.HasValue){object old=font.Italic;if(!ContextChartFontPlan.Flag(old).HasValue)throw Mixed();changes.Add(new NativeFormattingChange(()=>font.Italic=italic.Value,()=>font.Italic=old));}app.StartNewUndoEntry();NativeFormattingTransaction.Run(changes);}
  static InvalidOperationException Mixed(){return new InvalidOperationException("The requested chart-wide font property is mixed or unsupported. Select individual labels with PowerPoint formatting to preserve their distinct styles.");}
  internal static int? ReadTextColor(PowerPoint.Application app){return ContextChartFontPlan.Color(Font(app).Color);}
  internal static void ApplyColor(PowerPoint.Application app,int rgb){NativeFormattingPlan.ColorValue(rgb);var font=Font(app);object old=font.Color;if(!ContextChartFontPlan.Color(old).HasValue)throw Mixed();app.StartNewUndoEntry();NativeFormattingTransaction.Run(new[]{new NativeFormattingChange(()=>font.Color=rgb,()=>font.Color=old)});}
 }
}

