using System;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using PowerPoint=Microsoft.Office.Interop.PowerPoint;
namespace PowerPointIntelligence.AddIn
{
    internal sealed class ContextPropertiesController:IDisposable
    {
        private readonly PowerPoint.Application app;
        private readonly Func<bool> available;
        private readonly Action<Action> execute;
        private readonly Action<string,string> specific;
        private readonly Timer timer;
        private ContextPropertiesView view;
        private string lastKey,dismissedKey,observedKey,error,errorKey;
        private IntPtr lastOwner;
        private DateTime lastFull=DateTime.MinValue;
        private bool dirty=true,acting,disposed,refreshing;
        private string reason="not-refreshed",stage="startup",failureType="",failureHresult="";
        private long refreshCount;
        private const string EmptyObservation="<properties visible=\"false\" kind=\"None\" font=\"false\" fill=\"false\" outline=\"false\" selectionMatches=\"false\" foreground=\"false\" reason=\"not-refreshed\" stage=\"startup\" refreshes=\"0\" />";
        private string observation=EmptyObservation;
        internal bool IsVisible {get{return view!=null&&!view.IsDisposed&&view.Visible;}}
        internal string CurrentKey {get{return view==null?null:view.SelectionKey;}}
        private sealed class Owner:IWin32Window {internal IntPtr Value;public IntPtr Handle{get{return Value;}}}
        internal ContextPropertiesController(PowerPoint.Application app,Func<bool> available,Action<Action> execute,Action<string,string> specific)
        {this.app=app;this.available=available;this.execute=execute;this.specific=specific;timer=new Timer {Interval=300};timer.Tick+=(s,e)=>Refresh(false);timer.Start();}
        internal void Request(){dirty=true;}
        internal void Hide(string why="hidden"){reason=why;if(view!=null&&!view.IsDisposed)view.Hide();lastKey=null;dirty=true;PublishObservation(null);}
        internal void ShowCurrent(){dismissedKey=null;dirty=true;Refresh(true);}
        private void Refresh(bool force)
        {
            if(refreshing){dirty=true;return;}refreshing=true;
            string liveKey=null;refreshCount++;stage="availability";failureType=failureHresult="";
            try{
                if(view!=null&&view.IsDisposed)view=null;
                if(disposed){Hide("disposed");return;}
                if(acting){reason="action-in-progress";return;}
                if(!available()){Hide("commands-unavailable");return;}
                stage="foreground-before-selection";
                if(view!=null&&view.Visible&&!view.ContainsFocus&&lastOwner!=IntPtr.Zero){var fg=ContextSelectionBridge.GetForegroundWindow();if(fg!=lastOwner&&fg!=view.Handle&&!force){Hide("powerpoint-not-foreground");return;}}
                stage="selection";string key=liveKey=ContextSelectionBridge.Key(app);if(key==null){Hide("no-supported-selection");dismissedKey=null;observedKey=null;return;}if(key!=observedKey){Hide("selection-changing");dismissedKey=null;observedKey=key;error=null;errorKey=null;}
                if(view!=null&&view.ContainsFocus&&!force){reason="editing-properties";return;}
                if(key==dismissedKey&&!force){reason="dismissed";return;}
                if(!dirty&&key==lastKey&&(DateTime.UtcNow-lastFull).TotalMilliseconds<1500){reason="visible";return;}
                stage="capture";ContextNativeState state;try{state=ContextSelectionBridge.Capture(app);}catch(Exception detail) when(detail is InvalidOperationException||detail is System.IO.InvalidDataException){Failure(detail);state=ContextSelectionBridge.Capture(app,true);if(state!=null)state.Note=detail.Message;}if(state==null){Hide("selection-lost");return;}
                stage="capture-validation";if(!String.Equals(state.Snapshot.SelectionKey,ContextSelectionBridge.Key(app),StringComparison.Ordinal)){Hide("selection-changed-during-capture");return;}stage="foreground-after-capture";var foreground=ContextSelectionBridge.GetForegroundWindow();if(!force&&foreground!=state.Owner&&(view==null||foreground!=view.Handle)){Hide("powerpoint-not-foreground");return;}
                stage="create-view";if(view==null){view=new ContextPropertiesView();view.Command+=Apply;view.Dismissed+=keyClosed=>{dismissedKey=keyClosed;};}
                if(lastOwner!=state.Owner)view.Hide();
                stage="placement";Rectangle placement=ContextPanelModel.Place(state.SelectionScreen,state.WindowScreen,state.WorkArea,view.Size);
                stage="render";view.Render(state);if(errorKey==key&&error!=null)view.Error(error);view.Bounds=placement;stage="show";if(!view.Visible)view.Show(new Owner {Value=state.Owner});lastOwner=state.Owner;lastKey=key;lastFull=DateTime.UtcNow;dirty=false;reason="visible";stage="ready";
            }catch(Exception e) when(IsInspectionFailure(e)){Failure(e);Hide("inspection-failed");}
            finally{try{PublishObservation(liveKey);}finally{refreshing=false;}}
        }
        internal static bool IsInspectionFailure(Exception error){return error is COMException||error is InvalidComObjectException||error is InvalidCastException||error is NotImplementedException||error is InvalidOperationException||error is System.IO.InvalidDataException;}
        private void Failure(Exception exception){failureType=exception.GetType().Name;failureHresult="0x"+exception.HResult.ToString("X8",System.Globalization.CultureInfo.InvariantCulture);if(stage=="capture"||stage=="selection")stage+=":"+ContextSelectionBridge.InspectionStage;}
        private void Apply(string expected,string command,object value)
        {
            if(acting||disposed)return;acting=true;
            try{
                ContextSelectionBridge.Ensure(app,expected);
                if(command=="action"){string id=(string)value;var state=ContextSelectionBridge.Capture(app);if(!ContextPanelModel.CanExecute(state.Snapshot,expected,id))throw new InvalidOperationException("This action is unavailable for the current selection.");specific(expected,id);return;}
                execute(()=>{
                    try{ContextSelectionBridge.Ensure(app,expected);var state=ContextSelectionBridge.Capture(app);
                        if(command.StartsWith("font-",StringComparison.Ordinal)){
                            if(!state.Snapshot.CanFont)throw new InvalidOperationException("Font editing is unavailable for this selection.");string family=command=="font-family"?(string)value:null;float? size=command=="font-size"?(float?)Convert.ToSingle(value):null;bool? bold=command=="font-bold"?(bool?)(bool)value:null,italic=command=="font-italic"?(bool?)(bool)value:null;
                            if(state.Snapshot.Kind==ContextKind.Chart)ContextChartFormatting.ApplyFont(app,family,size,bold,italic);else ContextFormatting.ApplyFont(app,family,size,bold,italic);
                        }
                        else if(command=="text-mode"){if(!state.Snapshot.CanText)throw new InvalidOperationException("Text alignment is unavailable for this selection.");NativeFormatting.Text(app,(string)value);}
                        else if(command.StartsWith("color-",StringComparison.Ordinal)){
                            string kind=command.Substring(6);bool allowed=kind=="fill"?state.Snapshot.CanFill:kind=="outline"?state.Snapshot.CanOutline:state.Snapshot.CanText||state.Snapshot.Kind==ContextKind.Chart;if(!allowed)throw new InvalidOperationException("This color property is unavailable for the selection.");
                            using(var chooser=new ColorDialog {FullOpen=true,AnyColor=true}){int? initial=kind=="fill"?state.FillColor:kind=="outline"?state.OutlineColor:state.TextColor;if(initial.HasValue)chooser.Color=Color.FromArgb(initial.Value&255,(initial.Value>>8)&255,(initial.Value>>16)&255);if(chooser.ShowDialog(view)!=DialogResult.OK)return;ContextSelectionBridge.Ensure(app,expected);ContextSelectionBridge.ApplyColor(app,expected,kind,NativeFormattingPlan.Rgb(chooser.Color));}
                        }
                        else throw new InvalidOperationException("Unknown contextual property command.");
                        this.error=null;errorKey=null;
                    }catch(Exception error){this.error=error is COMException?"Office could not apply the property. Select the element again and retry.":error.Message;errorKey=expected;view.Error(this.error);throw;}
                });
            }catch(Exception error){this.error=error is COMException?"Office is busy or the selection is no longer available.":error.Message;errorKey=expected;if(view!=null)view.Error(this.error);}
            finally{acting=false;dirty=true;lastKey=null;}
        }
        // Called only by UI-thread refresh. External COM callers read this published string.
        private void PublishObservation(string liveKey)
        {
            var doc=new System.Xml.XmlDocument();var root=doc.CreateElement("properties");doc.AppendChild(root);var state=view==null||view.IsDisposed?null:view.State;bool visible=IsVisible;root.SetAttribute("visible",visible?"true":"false");root.SetAttribute("kind",state==null?"None":state.Snapshot.Kind.ToString());root.SetAttribute("font",state!=null&&state.Snapshot.CanFont?"true":"false");root.SetAttribute("fill",state!=null&&state.Snapshot.CanFill?"true":"false");root.SetAttribute("outline",state!=null&&state.Snapshot.CanOutline?"true":"false");root.SetAttribute("selectionMatches",visible&&state!=null&&ContextPanelModel.IsCurrent(state.Snapshot,liveKey)?"true":"false");root.SetAttribute("foreground",visible&&state!=null&&(ContextSelectionBridge.GetForegroundWindow()==state.Owner||view.ContainsFocus)?"true":"false");root.SetAttribute("nativeStage",ContextSelectionBridge.InspectionStage??"");root.SetAttribute("reason",reason);root.SetAttribute("stage",stage);root.SetAttribute("exception",failureType);root.SetAttribute("hresult",failureHresult);root.SetAttribute("refreshes",refreshCount.ToString(System.Globalization.CultureInfo.InvariantCulture));root.SetAttribute("uiSta",System.Threading.Thread.CurrentThread.GetApartmentState()==System.Threading.ApartmentState.STA?"true":"false");System.Threading.Volatile.Write(ref observation,doc.OuterXml);
        }
        internal string AutomationState(){return System.Threading.Volatile.Read(ref observation);}
        public void Dispose(){if(disposed)return;disposed=true;timer.Stop();timer.Dispose();if(view!=null){view.Dispose();view=null;}Hide("disposed");}
    }
}
