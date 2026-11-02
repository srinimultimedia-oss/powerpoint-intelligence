using System;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
namespace PowerPointIntelligence.AddIn
{
    internal sealed class PropertySwatch:Button
    {
        internal int? Rgb;
        protected override void OnPaint(PaintEventArgs e){base.OnPaint(e);var box=new Rectangle(6,Height/2-5,10,10);if(Rgb.HasValue){using(var brush=new SolidBrush(Color.FromArgb(Rgb.Value&255,(Rgb.Value>>8)&255,(Rgb.Value>>16)&255)))e.Graphics.FillRectangle(brush,box);}else{e.Graphics.DrawLine(Pens.Gray,box.Left,box.Bottom,box.Right,box.Top);}e.Graphics.DrawRectangle(Pens.Gray,box);}
    }
    internal sealed class ContextPropertiesView:Form
    {
        private readonly Label title,note;
        private readonly ComboBox family,size,horizontal,vertical,direction;
        private readonly CheckBox bold,italic;
        private readonly PropertySwatch textColor,fillColor,outlineColor;
        private readonly FlowLayoutPanel specific;
        private readonly ToolTip tips=new ToolTip();
        private bool rendering;
        private ContextNativeState current;
        internal event Action<string,string,object> Command;
        internal event Action<string> Dismissed;
        internal ContextNativeState State {get{return current;}}
        internal string SelectionKey {get{return current==null?null:current.Snapshot.SelectionKey;}}
        protected override bool ShowWithoutActivation {get{return true;}}
        internal ContextPropertiesView()
        {
            Text="Element properties · Intelligence";FormBorderStyle=FormBorderStyle.FixedToolWindow;ShowInTaskbar=false;MaximizeBox=false;MinimizeBox=false;StartPosition=FormStartPosition.Manual;AutoScaleMode=AutoScaleMode.Dpi;ClientSize=new Size(316,300);BackColor=Color.White;Font=new Font("Segoe UI",9);KeyPreview=true;
            var layout=new TableLayoutPanel {Dock=DockStyle.Fill,Padding=new Padding(10),ColumnCount=1,RowCount=7};layout.RowStyles.Add(new RowStyle(SizeType.Absolute,26));layout.RowStyles.Add(new RowStyle(SizeType.Absolute,34));layout.RowStyles.Add(new RowStyle(SizeType.Absolute,32));layout.RowStyles.Add(new RowStyle(SizeType.Absolute,34));layout.RowStyles.Add(new RowStyle(SizeType.Absolute,34));layout.RowStyles.Add(new RowStyle(SizeType.Percent,100));layout.RowStyles.Add(new RowStyle(SizeType.Absolute,36));
            title=new Label {Dock=DockStyle.Fill,Text="Element properties",AutoEllipsis=true,Font=new Font(Font,FontStyle.Bold)};layout.Controls.Add(title);
            var fontRow=new FlowLayoutPanel {Dock=DockStyle.Fill,WrapContents=false};family=new ComboBox {Width=176,AccessibleName="Font family",AutoCompleteMode=AutoCompleteMode.SuggestAppend,AutoCompleteSource=AutoCompleteSource.ListItems};using(var fonts=new System.Drawing.Text.InstalledFontCollection())family.Items.AddRange(fonts.Families.Select(f=>f.Name).Distinct().OrderBy(x=>x).Take(1000).Cast<object>().ToArray());size=new ComboBox {Width=76,AccessibleName="Font size in points"};size.Items.AddRange(new object[]{"8","10","12","14","16","18","20","24","28","32","36","44","48","60","72"});fontRow.Controls.AddRange(new Control[]{family,size});layout.Controls.Add(fontRow);
            family.SelectionChangeCommitted+=(s,e)=>Send("font-family",family.Text,family.Tag as string);family.KeyDown+=(s,e)=>{if(e.KeyCode==Keys.Enter){Send("font-family",family.Text,family.Tag as string);e.SuppressKeyPress=true;}};size.SelectionChangeCommitted+=(s,e)=>SizeCommand();size.KeyDown+=(s,e)=>{if(e.KeyCode==Keys.Enter){SizeCommand();e.SuppressKeyPress=true;}};
            var colorRow=new FlowLayoutPanel {Dock=DockStyle.Fill,WrapContents=false};bold=new CheckBox {Text="B",Appearance=Appearance.Button,AutoCheck=false,ThreeState=true,TextAlign=ContentAlignment.MiddleCenter,Width=30,Height=26,AccessibleName="Bold"};italic=new CheckBox {Text="I",Appearance=Appearance.Button,AutoCheck=false,ThreeState=true,TextAlign=ContentAlignment.MiddleCenter,Width=30,Height=26,AccessibleName="Italic"};textColor=new PropertySwatch {Text="    Font color",Width=110,Height=26,AccessibleName="Font color"};colorRow.Controls.AddRange(new Control[]{bold,italic,textColor});layout.Controls.Add(colorRow);bold.Click+=(s,e)=>Send("font-bold",current.Font==null||current.Font.Bold!=true,bold.Tag as string);italic.Click+=(s,e)=>Send("font-italic",current.Font==null||current.Font.Italic!=true,italic.Tag as string);textColor.Click+=(s,e)=>Send("color-text",null,textColor.Tag as string);
            var alignRow=new FlowLayoutPanel {Dock=DockStyle.Fill,WrapContents=false};horizontal=Choices(146,"Horizontal text alignment",new[]{"Mixed / unavailable","Left","Center","Right","Justify"});vertical=Choices(132,"Vertical text alignment",new[]{"Mixed / unavailable","Top","Middle","Bottom"});alignRow.Controls.AddRange(new Control[]{horizontal,vertical});layout.Controls.Add(alignRow);horizontal.SelectionChangeCommitted+=(s,e)=>{if(horizontal.SelectedIndex>0)Send("text-mode",horizontal.Text.ToLowerInvariant(),horizontal.Tag as string);};vertical.SelectionChangeCommitted+=(s,e)=>{if(vertical.SelectedIndex>0)Send("text-mode",vertical.Text.ToLowerInvariant(),vertical.Tag as string);};
            var surfaceRow=new FlowLayoutPanel {Dock=DockStyle.Fill,WrapContents=false};direction=Choices(105,"Text direction",new[]{"Mixed / unavailable","Horizontal","Upward","Downward"});fillColor=new PropertySwatch {Text="    Fill",Width=78,Height=26,AccessibleName="Fill color"};outlineColor=new PropertySwatch {Text="    Outline",Width=93,Height=26,AccessibleName="Outline color"};surfaceRow.Controls.AddRange(new Control[]{direction,fillColor,outlineColor});layout.Controls.Add(surfaceRow);direction.SelectionChangeCommitted+=(s,e)=>{if(direction.SelectedIndex>0)Send("text-mode",direction.Text.ToLowerInvariant(),direction.Tag as string);};fillColor.Click+=(s,e)=>Send("color-fill",null,fillColor.Tag as string);outlineColor.Click+=(s,e)=>Send("color-outline",null,outlineColor.Tag as string);
            specific=new FlowLayoutPanel {Dock=DockStyle.Fill,AutoScroll=true,WrapContents=true,AccessibleName="Element specific properties"};layout.Controls.Add(specific);note=new Label {Dock=DockStyle.Fill,AutoEllipsis=true,ForeColor=Color.DimGray,AccessibleName="Property box status"};layout.Controls.Add(note);Controls.Add(layout);
            foreach(var editor in new Control[]{family,size,bold,italic,textColor,fillColor,outlineColor,horizontal,vertical,direction}){var bound=editor;bound.Enter+=(s,e)=>bound.Tag=SelectionKey;}
            tips.SetToolTip(family,"Choose a font or type its name, then press Enter. Mixed selection is shown blank.");tips.SetToolTip(size,"Font size in points. Enter applies only size; other formatting is retained.");tips.SetToolTip(textColor,"Font color. A diagonal swatch means mixed, absent or unavailable.");tips.SetToolTip(fillColor,"Fill color. For a chart this means its chart-area background.");tips.SetToolTip(outlineColor,"Outline color. Unsupported controls are disabled.");KeyDown+=(s,e)=>{if(e.KeyCode==Keys.Escape){Dismiss();e.Handled=e.SuppressKeyPress=true;}};FormClosing+=(s,e)=>{if(e.CloseReason==CloseReason.UserClosing){e.Cancel=true;Dismiss();}};
        }
        private static ComboBox Choices(int width,string name,string[] choices){var c=new ComboBox {Width=width,DropDownStyle=ComboBoxStyle.DropDownList,AccessibleName=name};c.Items.AddRange(choices);return c;}
        private void SizeCommand(){float value;if(!Single.TryParse(size.Text,NumberStyles.Float,CultureInfo.InvariantCulture,out value)||Single.IsNaN(value)||Single.IsInfinity(value)||value<1||value>4000){Error("Use a font size between1 and4,000 points.");return;}Send("font-size",value,size.Tag as string);}
        private void Send(string id,object value,string expected=null){if(!rendering&&current!=null&&Command!=null)Command(expected??current.Snapshot.SelectionKey,id,value);}
        private void Dismiss(){if(Dismissed!=null)Dismissed(SelectionKey);Hide();}
        private static void Select(ComboBox combo,string value){combo.SelectedIndex=0;if(value!=null)for(int i=1;i<combo.Items.Count;i++)if(String.Equals(combo.Items[i].ToString(),value,StringComparison.OrdinalIgnoreCase)){combo.SelectedIndex=i;break;}}
        internal void Render(ContextNativeState state)
        {
            current=state;rendering=true;try{title.Text=state.Snapshot.Title+" properties";var f=state.Font;family.Text=f==null?"":f.Family??"";size.Text=f==null||!f.Size.HasValue?"":f.Size.Value.ToString("0.##",CultureInfo.InvariantCulture);family.Enabled=size.Enabled=bold.Enabled=italic.Enabled=state.Snapshot.CanFont&&f!=null;bold.CheckState=f==null||!f.Bold.HasValue?CheckState.Indeterminate:f.Bold.Value?CheckState.Checked:CheckState.Unchecked;italic.CheckState=f==null||!f.Italic.HasValue?CheckState.Indeterminate:f.Italic.Value?CheckState.Checked:CheckState.Unchecked;
                horizontal.Enabled=vertical.Enabled=direction.Enabled=state.Snapshot.CanText;Select(horizontal,state.Horizontal);Select(vertical,state.Vertical);Select(direction,state.Direction);textColor.Enabled=state.Snapshot.CanText||state.Snapshot.Kind==ContextKind.Chart&&!state.ReadOnly;fillColor.Enabled=state.Snapshot.CanFill;outlineColor.Enabled=state.Snapshot.CanOutline;textColor.Rgb=state.TextColor;fillColor.Rgb=state.FillColor;outlineColor.Rgb=state.OutlineColor;textColor.Invalidate();fillColor.Invalidate();outlineColor.Invalidate();
                foreach(Control old in specific.Controls.Cast<Control>().ToArray())old.Dispose();specific.Controls.Clear();foreach(var id in state.Snapshot.ActionIds){string action=id;var b=new Button {Text=Caption(id),AutoSize=true,Height=28,AccessibleName=Caption(id),Tag=id};string expected=state.Snapshot.SelectionKey;b.Click+=(s,e)=>{if(!rendering&&Command!=null)Command(expected,"action",action);};specific.Controls.Add(b);}note.ForeColor=Color.DimGray;note.Text=state.Note??"Changes apply to this selection. Esc closes until you select again.";
            }finally{rendering=false;}
        }
        internal Bitmap Preview()
        {
            var content=Controls[0];using(var host=new Panel {Size=ClientSize,Font=Font,BackColor=BackColor}){
                content.Parent=host;try{host.CreateControl();host.PerformLayout();foreach(var c in Descendants(host)){var handle=c.Handle;c.PerformLayout();}var image=new Bitmap(host.Width,host.Height);host.DrawToBitmap(image,new Rectangle(Point.Empty,host.Size));
                    // WinForms WM_PRINT omits native combo edit text on hidden controls.
                    // Preview-only overlay uses actual control values; the live form uses native rendering.
                    using(var g=Graphics.FromImage(image))foreach(var combo in Descendants(host).OfType<ComboBox>()){int x=combo.Left,y=combo.Top;for(Control parent=combo.Parent;parent!=null&&parent!=host;parent=parent.Parent){x+=parent.Left;y+=parent.Top;}var box=new Rectangle(x+2,y+2,Math.Max(1,combo.Width-20),Math.Max(1,combo.Height-4));g.FillRectangle(SystemBrushes.Window,box);using(var brush=new SolidBrush(combo.Enabled?SystemColors.WindowText:SystemColors.GrayText))using(var format=new StringFormat {LineAlignment=StringAlignment.Center,Trimming=StringTrimming.EllipsisCharacter,FormatFlags=StringFormatFlags.NoWrap})g.DrawString(combo.Text,combo.Font,brush,box,format);}
                    return image;}finally{content.Parent=this;}
            }
        }
        private static System.Collections.Generic.IEnumerable<Control> Descendants(Control root){foreach(Control c in root.Controls){yield return c;foreach(var child in Descendants(c))yield return child;}}
        internal void Error(string message){note.ForeColor=Color.Firebrick;note.Text=message;}
        internal static string Caption(string id){switch(id){case "arrange":return "Arrange objects";case "edit":return "Element settings";case "fit":return "Fit text";case "indicator-next":return "Increase / next";case "indicator-previous":return "Decrease / previous";case "indicator-keys":return "Keyboard controls";case "convert-text":return "Text Box";case "convert-pentagon":return "Pentagon";case "convert-chevron":return "Chevron";case "convert-rounded":return "Rounded rectangle";case "convert-harvey":return "Harvey ball";case "convert-checkbox":return "Checkbox";case "grid-row":return "Add row";case "grid-column":return "Add column";case "grid-remove-row":return "Hide row";case "grid-remove-column":return "Hide column";case "grid-left":return "Column left";case "grid-right":return "Column right";case "grid-stripe-rows":return "Row colors";case "grid-stripe-columns":return "Column colors";case "grid-resize":return "Resize grid";case "chart-data":return "Edit chart data";case "chart-series":return "Series and axes";case "chart-axes":return "Value axis";case "chart-style":return "Chart style";case "chart-legend":return "Legend";case "chart-labels":return "Value labels";case "chart-data-table":return "Data table";case "data-links":return "Data links";case "table-data":return "Edit table data";case "template-palette":return "Template colors";case "template-alignment":return "Template alignment";case "agenda-settings":return "Manage agenda";case "agenda-refresh":return "Update agenda";case "annotation-refresh":return "Refresh annotation";}if(id.StartsWith("duplicate-",StringComparison.Ordinal))return "Duplicate "+id.Substring(10);return id;}
        protected override void Dispose(bool disposing){if(disposing){tips.Dispose();}base.Dispose(disposing);}
    }
}
