using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
namespace PowerPointIntelligence.AddIn
{
    internal enum ContextKind {None,Text,Shape,Table,Chart,Connector,Indicator,ProcessGrid,Template,Agenda,Image,Mixed}
    internal sealed class ContextDescriptor
    {
        internal int ShapeId;
        internal bool IsTextBox,IsShape,IsTable,IsChart,IsConnector,IsImage;
        internal bool CanText,CanFill,CanOutline,CanFont;
        internal string ManagedType;
        internal string[] ActionIds;
    }
    internal sealed class ContextSnapshot
    {
        internal readonly ContextKind Kind;
        internal readonly string Title,SelectionKey;
        internal readonly int SelectedCount;
        internal readonly bool CanText,CanFill,CanOutline,CanFont,CanElementEdit;
        internal readonly IReadOnlyList<string> ActionIds;
        internal ContextSnapshot(ContextKind kind,string title,string selectionKey,int count,bool text,bool fill,bool outline,bool font,IEnumerable<string> actions)
        {
            Kind=kind;Title=title;SelectionKey=selectionKey;SelectedCount=count;CanText=text;CanFill=fill;CanOutline=outline;CanFont=font;var ids=actions.ToArray();ActionIds=Array.AsReadOnly(ids);CanElementEdit=ids.Length>0;
        }
    }
    internal enum MixedValueKind {Empty,Uniform,Mixed}
    internal sealed class MixedValue<T>
    {
        internal readonly MixedValueKind Kind;
        private readonly T value;
        internal bool HasValue {get{return Kind==MixedValueKind.Uniform;}}
        internal T Value {get{if(!HasValue)throw new InvalidOperationException("No single common property value exists.");return value;}}
        private MixedValue(MixedValueKind kind,T value){Kind=kind;this.value=value;}
        internal static MixedValue<T> From(IEnumerable<T> values,IEqualityComparer<T> comparer=null)
        {
            if(values==null)throw new ArgumentNullException("values");comparer=comparer??EqualityComparer<T>.Default;bool found=false,mixed=false;T first=default(T);int count=0;foreach(var item in values){if(++count>10000)throw new InvalidDataException("Mixed property inspection is limited to 10000 values.");if(!found){first=item;found=true;}else if(!comparer.Equals(first,item))mixed=true;}return new MixedValue<T>(!found?MixedValueKind.Empty:mixed?MixedValueKind.Mixed:MixedValueKind.Uniform,first);
        }
    }
    internal static class ContextPanelModel
    {
        private static readonly string[] TextActions={"edit","fit","convert-text","convert-pentagon","convert-chevron","convert-rounded","duplicate-left","duplicate-right","duplicate-up","duplicate-down"};
        private static readonly string[] GridActions={"edit","grid-row","grid-column","grid-remove-row","grid-remove-column","grid-left","grid-right","grid-stripe-rows","grid-stripe-columns","grid-resize"};
        private static readonly string[] IndicatorActions={"edit","indicator-next","indicator-previous","indicator-keys","convert-harvey","convert-checkbox"};
        private static string Managed(ContextDescriptor d){return (d.ManagedType??"").Trim().ToLowerInvariant();}
        internal static ContextKind Classify(ContextDescriptor d)
        {
            if(d==null||d.ShapeId<1)throw new InvalidDataException("Context selection contains an invalid native shape identity.");
            if(new[]{d.IsTextBox,d.IsTable,d.IsChart,d.IsConnector,d.IsImage}.Count(v=>v)>1)throw new InvalidDataException("Context selection has conflicting native classifications.");
            switch(Managed(d)){case "indicator":return ContextKind.Indicator;case "grid":case "process-grid":return ContextKind.ProcessGrid;case "template":return ContextKind.Template;case "agenda":return ContextKind.Agenda;case "connector":return ContextKind.Connector;case "scene":case "annotation":return ContextKind.Shape;case "text":return ContextKind.Text;}
            return d.IsTable?ContextKind.Table:d.IsChart?ContextKind.Chart:d.IsConnector?ContextKind.Connector:d.IsImage?ContextKind.Image:d.IsTextBox?ContextKind.Text:ContextKind.Shape;
        }
        private static string[] Actions(ContextDescriptor d)
        {
            IEnumerable<string> defaults=Enumerable.Empty<string>();switch(Managed(d)){case "text":case "text-element":case "text-shape":defaults=TextActions;break;case "indicator":defaults=IndicatorActions;break;case "process-grid":case "grid":defaults=GridActions;break;case "connector":defaults=new[]{"edit"};break;}
            var custom=d.ActionIds??new string[0];if(custom.Length>64||custom.Any(a=>String.IsNullOrEmpty(a)||a.Length>64||!Regex.IsMatch(a,@"\A[a-z][a-z0-9:.\-]*\z",RegexOptions.CultureInvariant)))throw new InvalidDataException("Invalid or excessive contextual action identifiers.");return defaults.Concat(custom).Distinct(StringComparer.Ordinal).ToArray();
        }
        internal static ContextSnapshot Build(IEnumerable<ContextDescriptor> descriptors,string selectionKey)
        {
            if(descriptors==null)throw new ArgumentNullException("descriptors");var all=descriptors.Take(501).ToArray();if(all.Length>500)throw new InvalidDataException("Context panels support at most 500 selected objects.");
            if(all.Length==0)return new ContextSnapshot(ContextKind.None,"",selectionKey??"",0,false,false,false,false,new string[0]);
            if(String.IsNullOrEmpty(selectionKey)||selectionKey.Length>8192||all.Any(d=>d==null)||all.Select(d=>d.ShapeId).Distinct().Count()!=all.Length)throw new InvalidDataException("Invalid or duplicate contextual selection identity.");
            // The native bridge owns the opaque identity envelope, including document/table/chart sub-selection details.
            var kinds=all.Select(Classify).ToArray();ContextKind kind=kinds.All(k=>k==kinds[0])?kinds[0]:ContextKind.Mixed;var actions=all.Select(Actions).ToArray();IEnumerable<string> common=actions[0];for(int i=1;i<actions.Length;i++)common=common.Intersect(actions[i],StringComparer.Ordinal);
            // Native structural commands require a single object. Multi-selection fitting is the only shared element command currently implemented.
            if(all.Length>1)common=common.Where(id=>id=="arrange"||id=="fit"&&all.All(d=>new[]{"text","text-element","text-shape"}.Contains(Managed(d))));
            string title=kind==ContextKind.Mixed?"Mixed selection":kind==ContextKind.ProcessGrid?"Process grid":kind.ToString();if(all.Length>1)title+=" · "+all.Length.ToString(CultureInfo.InvariantCulture)+" objects";
            return new ContextSnapshot(kind,title,selectionKey,all.Length,all.All(d=>d.CanText),all.All(d=>d.CanFill),all.All(d=>d.CanOutline),all.All(d=>d.CanFont),common);
        }
        internal static string SelectionKey(long windowId,int slideId,IEnumerable<int> shapeIds,int? textStart=null,int? textLength=null)
        {
            if(windowId==0||slideId<1||shapeIds==null)throw new InvalidDataException("Selection requires a live window and slide identity.");var ids=shapeIds.Take(501).ToArray();if(ids.Length>500||ids.Any(i=>i<1)||ids.Distinct().Count()!=ids.Length||textStart.HasValue!=textLength.HasValue||textStart.HasValue&&(ids.Length!=1||textStart.Value<0||textLength.Value<0||(long)textStart.Value+textLength.Value>Int32.MaxValue))throw new InvalidDataException("Invalid contextual shape/text selection identity.");
            return windowId.ToString(CultureInfo.InvariantCulture)+":"+slideId.ToString(CultureInfo.InvariantCulture)+":"+(ids.Length==0?"none":String.Join(",",ids.OrderBy(i=>i).Select(i=>i.ToString(CultureInfo.InvariantCulture))))+(textStart.HasValue?":text:"+textStart.Value.ToString(CultureInfo.InvariantCulture)+":"+textLength.Value.ToString(CultureInfo.InvariantCulture):"");
        }
        internal static bool IsCurrent(ContextSnapshot snapshot,string liveKey){return snapshot!=null&&snapshot.SelectedCount>0&&!String.IsNullOrEmpty(liveKey)&&String.Equals(snapshot.SelectionKey,liveKey,StringComparison.Ordinal);}
        internal static bool CanExecute(ContextSnapshot snapshot,string liveKey,string action){return IsCurrent(snapshot,liveKey)&&snapshot.ActionIds.Contains(action,StringComparer.Ordinal);}
        private static void RectangleBounds(Rectangle r,bool selection)
        {
            if(r.Width<(selection?0:1)||r.Height<(selection?0:1)||r.Width>1000000||r.Height>1000000||(long)r.X+r.Width>Int32.MaxValue||(long)r.X+r.Width<Int32.MinValue||(long)r.Y+r.Height>Int32.MaxValue||(long)r.Y+r.Height<Int32.MinValue)throw new InvalidDataException("Invalid contextual screen bounds.");
        }
        private static int Clamp(long value,int minimum,int maximum){return (int)Math.Max(minimum,Math.Min(maximum,value));}
        private static long Area(Rectangle a,Rectangle b){var intersection=Rectangle.Intersect(a,b);return intersection.Width<1||intersection.Height<1?0:(long)intersection.Width*intersection.Height;}
        internal static Rectangle Place(Rectangle selectedScreen,Rectangle windowScreen,Rectangle monitorWorkArea,Size panel,int gap=12)
        {
            RectangleBounds(selectedScreen,true);RectangleBounds(windowScreen,false);RectangleBounds(monitorWorkArea,false);if(panel.Width<1||panel.Height<1||panel.Width>1000000||panel.Height>1000000||gap<0||gap>128)throw new InvalidDataException("Invalid contextual panel size or gap.");var area=Rectangle.Intersect(windowScreen,monitorWorkArea);if(area.Width<panel.Width||area.Height<panel.Height)throw new InvalidOperationException("The property panel cannot fit in the available window work area.");
            var target=Rectangle.Intersect(selectedScreen,area);if(target.Width<=0||target.Height<=0)target=new Rectangle(Clamp(selectedScreen.X,area.Left,area.Right-1),Clamp(selectedScreen.Y,area.Top,area.Bottom-1),1,1);
            Func<long,long,Rectangle> candidate=(x,y)=>new Rectangle(Clamp(x,area.Left,area.Right-panel.Width),Clamp(y,area.Top,area.Bottom-panel.Height),panel.Width,panel.Height);
            var choices=new[]{candidate((long)target.Right+gap,target.Top),candidate((long)target.Left-panel.Width-gap,target.Top),candidate(target.Left,(long)target.Bottom+gap),candidate(target.Left,(long)target.Top-panel.Height-gap)};
            foreach(var choice in choices)if(Area(choice,target)==0)return choice;
            return choices.OrderBy(c=>Area(c,target)).First();
        }
    }
}


