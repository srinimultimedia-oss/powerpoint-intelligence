using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Office=Microsoft.Office.Core;
using PowerPoint=Microsoft.Office.Interop.PowerPoint;
namespace PowerPointIntelligence.AddIn
{
    internal static class ElementCommands
    {
        // Kept independent of COM so the ribbon's command set can be checked without Office activation.
        internal static readonly string[] Ids={"pentagon","chevron","connector","edit","fit","convert-text","convert-pentagon","convert-chevron","convert-rounded","duplicate-left","duplicate-right","duplicate-up","duplicate-down","indicator-next","indicator-previous","indicator-keys","convert-harvey","convert-checkbox","grid-create","grid-row","grid-column","grid-remove-row","grid-remove-column","grid-left","grid-right","grid-stripe-rows","grid-stripe-columns","grid-resize","refresh"};
        internal static bool Enabled(PowerPoint.Application app,string id)
        {
            if(!Ids.Contains(id)||app==null)return false;
            try{if(app.Presentations.Count==0||app.ActiveWindow.ViewType!=PowerPoint.PpViewType.ppViewNormal)return false;if(new[]{"pentagon","chevron","connector","grid-create","refresh"}.Contains(id))return true;var s=app.ActiveWindow.Selection;if(s.Type!=PowerPoint.PpSelectionType.ppSelectionShapes&&s.Type!=PowerPoint.PpSelectionType.ppSelectionText)return false;if(id=="fit")return s.Type==PowerPoint.PpSelectionType.ppSelectionShapes&&s.ShapeRange.Count>0&&s.ShapeRange.Count<=100&&s.ShapeRange.Cast<PowerPoint.Shape>().All(x=>TextShapeElements.IsManaged(x)&&String.IsNullOrEmpty(x.Tags[ProcessFlowTools.GridTag]));if(s.ShapeRange.Count!=1)return false;var shape=s.ShapeRange[1];if(id.StartsWith("grid-",StringComparison.Ordinal))return !String.IsNullOrEmpty(shape.Tags[ProcessFlowTools.GridTag]);if(id=="edit")return TextShapeElements.IsManaged(shape)||UniversalConnectors.IsManaged(shape)||IndicatorElements.IsIndicator(s);if(id.StartsWith("indicator-",StringComparison.Ordinal)||id=="convert-harvey"||id=="convert-checkbox")return IndicatorElements.IsIndicator(s);return TextShapeElements.IsManaged(shape)&&String.IsNullOrEmpty(shape.Tags[ProcessFlowTools.GridTag]);}catch(System.Runtime.InteropServices.COMException){return false;}
        }
        internal static void Run(PowerPoint.Application app,string id)
        {
            if(!Enabled(app,id))throw new InvalidOperationException("Select the relevant managed element in Normal view first.");
            if(id.StartsWith("duplicate-",StringComparison.Ordinal)){TextShapeElements.DuplicateSelected(app,id.Substring(10));return;}
            if(id.StartsWith("convert-",StringComparison.Ordinal)){var kind=id.Substring(8);if(kind=="harvey"||kind=="checkbox")IndicatorElements.ConvertSelected(app,kind);else TextShapeElements.ConvertSelected(app,kind=="rounded"?"roundedRectangle":kind);return;}
            switch(id){
                case "pentagon":TextShapeElements.Insert(app,"rightPentagon");break;
                case "chevron":TextShapeElements.Insert(app,"chevron");break;
                case "connector":UniversalConnectors.Insert(app);break;
                case "edit":var s=app.ActiveWindow.Selection;if(!String.IsNullOrEmpty(s.ShapeRange[1].Tags[ProcessFlowTools.GridTag]))Resize(app);else if(IndicatorElements.IsIndicator(s))IndicatorElements.Edit(app);else if(UniversalConnectors.IsManaged(s.ShapeRange[1]))UniversalConnectors.Edit(app);else TextShapeElements.Edit(app);break;
                case "fit":TextShapeElements.FitSelected(app);break;
                case "indicator-next":IndicatorElements.Step(app,1);break;
                case "indicator-previous":IndicatorElements.Step(app,-1);break;
                case "indicator-keys":IndicatorKeys(app);break;
                case "grid-create":ProcessFlowTools.Create(app);break;
                case "grid-row":ProcessFlowTools.AddRow(app);break;
                case "grid-column":ProcessFlowTools.AddColumn(app);break;
                case "grid-remove-row":if(MessageBox.Show("Hide the selected row? Its cells are retained as inactive history. PowerPoint Undo can restore this action.","Process grid",MessageBoxButtons.OKCancel)==DialogResult.OK)ProcessFlowTools.DeleteSelectedRowOrColumn(app,false);break;
                case "grid-remove-column":if(MessageBox.Show("Hide the selected column? Its cells are retained as inactive history. PowerPoint Undo can restore this action.","Process grid",MessageBoxButtons.OKCancel)==DialogResult.OK)ProcessFlowTools.DeleteSelectedRowOrColumn(app,true);break;
                case "grid-left":ProcessFlowTools.ReorderColumn(app,false);break;
                case "grid-right":ProcessFlowTools.ReorderColumn(app,true);break;
                case "grid-stripe-rows":Stripe(app,"row");break;
                case "grid-stripe-columns":Stripe(app,"column");break;
                case "grid-resize":Resize(app);break;
                case "refresh":TextShapeElements.AutoReflow(app);UniversalConnectors.Reflow(app);break;
                default:throw new InvalidOperationException("Unknown element command.");
            }
        }
        private static void Stripe(PowerPoint.Application app,string direction)
        {using(var first=new ColorDialog {FullOpen=true,Color=Color.White})using(var second=new ColorDialog {FullOpen=true,Color=Color.FromArgb(232,240,255)}){if(first.ShowDialog()!=DialogResult.OK||second.ShowDialog()!=DialogResult.OK)return;Func<Color,int> rgb=c=>c.R|(c.G<<8)|(c.B<<16);ProcessFlowTools.Stripe(app,direction,rgb(first.Color),rgb(second.Color));}}
        private static void Resize(PowerPoint.Application app)
        {using(var f=new Form {Text="Resize process grid",Width=440,Height=285,StartPosition=FormStartPosition.CenterParent,FormBorderStyle=FormBorderStyle.FixedDialog,MaximizeBox=false,MinimizeBox=false}){var fields=new TextBox[4];string[] names={"Left","Top","Width","Height"},defaults={"36 pt","90 pt","80%w","55%h"};for(int i=0;i<4;i++){fields[i]=new TextBox {Left=130,Top=20+38*i,Width=250,Text=defaults[i]};f.Controls.Add(new Label {Left=20,Top=24+38*i,Text=names[i]});f.Controls.Add(fields[i]);}var apply=new Button {Text="Apply",Left=210,Top=182,Width=80};var cancel=new Button {Text="Cancel",Left=300,Top=182,Width=80,DialogResult=DialogResult.Cancel};f.Controls.AddRange(new Control[]{apply,cancel});f.CancelButton=cancel;apply.Click+=(s,e)=>{try{float w=app.ActivePresentation.PageSetup.SlideWidth,h=app.ActivePresentation.PageSetup.SlideHeight;var v=fields.Select(t=>TextShapePlan.ParseLength(t.Text,w,h)).ToArray();ProcessFlowTools.Resize(app,v[0],v[1],v[2],v[3]);f.DialogResult=DialogResult.OK;}catch(Exception error){MessageBox.Show(f,error.Message,"Resize process grid");}};f.ShowDialog();}}
        private static void IndicatorKeys(PowerPoint.Application app)
        {
            var window=app.ActiveWindow;var slide=(PowerPoint.Slide)window.View.Slide;var selected=window.Selection.ShapeRange[1];int id=selected.Id;
            using(var f=new Form {Text="Indicator keyboard controls",Width=460,Height=205,StartPosition=FormStartPosition.CenterParent,KeyPreview=true})
            {f.Controls.Add(new Label {Left=20,Top=20,Width=410,Height=90,Text="Keep this panel focused.\r\nHarvey: + / −, 0–8. Checkbox: Space / 0, V / 1, X / 2.\r\nEnter cycles. Escape closes. Values clamp at endpoints.\r\nFor more than 8 segments, digit k selects nearest k/8 fill."});var close=new Button {Text="Close",Left=335,Top=120,DialogResult=DialogResult.Cancel};f.Controls.Add(close);f.CancelButton=close;
                f.KeyDown+=(sender,e)=>{try{if(app.ActiveWindow!=window||((PowerPoint.Slide)window.View.Slide).SlideID!=slide.SlideID)throw new InvalidOperationException("Return to the original slide.");var matches=slide.Shapes.Cast<PowerPoint.Shape>().Where(s=>s.Id==id).ToArray();if(matches.Length!=1)throw new InvalidOperationException("The original indicator is unavailable.");matches[0].Select();bool done=true;if(e.KeyCode==Keys.Add||e.KeyCode==Keys.Oemplus)IndicatorElements.Step(app,1);else if(e.KeyCode==Keys.Subtract||e.KeyCode==Keys.OemMinus)IndicatorElements.Step(app,-1);else if(e.KeyCode>=Keys.D0&&e.KeyCode<=Keys.D8)IndicatorElements.SetShortcut(app,(int)e.KeyCode-(int)Keys.D0);else if(e.KeyCode>=Keys.NumPad0&&e.KeyCode<=Keys.NumPad8)IndicatorElements.SetShortcut(app,(int)e.KeyCode-(int)Keys.NumPad0);else if(e.KeyCode==Keys.Space)IndicatorElements.SetShortcut(app,0);else if(e.KeyCode==Keys.V)IndicatorElements.SetShortcut(app,1);else if(e.KeyCode==Keys.X)IndicatorElements.SetShortcut(app,2);else if(e.KeyCode==Keys.Enter)IndicatorElements.CycleSelected(app);else done=false;if(done){id=app.ActiveWindow.Selection.ShapeRange[1].Id;e.Handled=e.SuppressKeyPress=true;}}catch(Exception error){MessageBox.Show(f,error.Message,"Indicator controls");}};f.ShowDialog();}
        }
    }
}
