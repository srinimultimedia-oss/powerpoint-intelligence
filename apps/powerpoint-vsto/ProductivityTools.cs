using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Office=Microsoft.Office.Core;
using PowerPoint=Microsoft.Office.Interop.PowerPoint;
namespace PowerPointIntelligence.AddIn
{
    internal sealed class ProductivitySymbol
    {
        internal readonly string Glyph,Name;
        internal ProductivitySymbol(string glyph,string name){Glyph=glyph;Name=name;}
        public override string ToString(){return Glyph+"   "+Name+"   U+"+Char.ConvertToUtf32(Glyph,0).ToString("X4",CultureInfo.InvariantCulture);}
    }
    internal sealed class DecimalCharacterChange
    {
        internal int Index; internal char Before,After;
    }
    internal static class ProductivityTools
    {
        internal static readonly string[] Ids={"symbol-picker","decimal-safe","raster-copy","chart-workbooks"};
        internal const int MaxCharacters=20000,MaxTargets=100;
        internal static readonly ProductivitySymbol[] Symbols={
            new ProductivitySymbol("→","right arrow"),new ProductivitySymbol("←","left arrow"),new ProductivitySymbol("↑","up arrow"),new ProductivitySymbol("↓","down arrow"),new ProductivitySymbol("↔","left right arrow"),new ProductivitySymbol("⇒","right double arrow"),new ProductivitySymbol("✓","check mark"),new ProductivitySymbol("✗","cross mark"),new ProductivitySymbol("☐","empty ballot box"),new ProductivitySymbol("☑","checked ballot box"),new ProductivitySymbol("●","black circle"),new ProductivitySymbol("○","white circle"),new ProductivitySymbol("◐","half circle"),new ProductivitySymbol("■","black square"),new ProductivitySymbol("□","white square"),new ProductivitySymbol("▲","up triangle"),new ProductivitySymbol("▼","down triangle"),new ProductivitySymbol("★","star"),new ProductivitySymbol("☆","outline star"),new ProductivitySymbol("±","plus minus"),new ProductivitySymbol("×","multiplication"),new ProductivitySymbol("÷","division"),new ProductivitySymbol("≤","less than or equal"),new ProductivitySymbol("≥","greater than or equal"),new ProductivitySymbol("≠","not equal"),new ProductivitySymbol("≈","approximately equal"),new ProductivitySymbol("∞","infinity"),new ProductivitySymbol("∑","sum sigma"),new ProductivitySymbol("Δ","delta change"),new ProductivitySymbol("α","alpha"),new ProductivitySymbol("β","beta"),new ProductivitySymbol("π","pi"),new ProductivitySymbol("μ","micro mu"),new ProductivitySymbol("°","degree"),new ProductivitySymbol("€","euro"),new ProductivitySymbol("£","pound"),new ProductivitySymbol("¥","yen"),new ProductivitySymbol("₹","rupee"),new ProductivitySymbol("©","copyright"),new ProductivitySymbol("®","registered"),new ProductivitySymbol("™","trademark"),new ProductivitySymbol("•","bullet"),new ProductivitySymbol("…","ellipsis"),new ProductivitySymbol("–","en dash"),new ProductivitySymbol("—","em dash")};
        internal static ProductivitySymbol[] SearchSymbols(string query)
        {query=(query??"").Trim();if(query.Length>100)throw new ArgumentException("Search is limited to 100 characters.");return Symbols.Where(s=>s.ToString().IndexOf(query,StringComparison.OrdinalIgnoreCase)>=0).ToArray();}
        internal static DecimalCharacterChange[] DecimalChanges(string text)
        {
            if(text==null)throw new ArgumentNullException("text");if(text.Length>MaxCharacters)throw new InvalidDataException("Selected text exceeds 20,000 characters.");
            var changes=new List<DecimalCharacterChange>();for(int i=1;i+1<text.Length;i++)if((text[i]=='.'||text[i]==',')&&Char.IsDigit(text[i-1])&&Char.IsDigit(text[i+1]))changes.Add(new DecimalCharacterChange{Index=i,Before=text[i],After=text[i]=='.'?',':'.'});return changes.ToArray();
        }
        internal static void Run(PowerPoint.Application app,string id)
        {
            if(!Ids.Contains(id))throw new ArgumentException("Unknown local productivity tool.");ActiveSlide(app);
            switch(id){case "symbol-picker":ShowSymbols(app);break;case "decimal-safe":SwitchDecimals(app);break;case "raster-copy":RasterCopy(app);break;case "chart-workbooks":ShowWorkbooks(app);break;}
        }
        private static PowerPoint.Slide ActiveSlide(PowerPoint.Application app)
        {
            if(app==null||app.Presentations.Count==0||app.ActiveWindow.ViewType!=PowerPoint.PpViewType.ppViewNormal)throw new InvalidOperationException("Open a presentation and select a slide in Normal view.");
            var slide=app.ActiveWindow.View.Slide as PowerPoint.Slide;if(slide==null)throw new InvalidOperationException("Select a slide in Normal view.");return slide;
        }
        internal static void InsertSymbol(PowerPoint.Application app,string symbol)
        {
            var slide=ActiveSlide(app);if(!Symbols.Any(s=>s.Glyph==symbol))throw new ArgumentException("Choose a catalog symbol.");app.StartNewUndoEntry();
            if(app.ActiveWindow.Selection.Type==PowerPoint.PpSelectionType.ppSelectionText){var range=app.ActiveWindow.Selection.TextRange;if(range.Length==0)range.InsertAfter(symbol);else range.Text=symbol;return;}
            PowerPoint.Shape added=null;try{added=slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,60,80,120,60);added.TextFrame.TextRange.Text=symbol;added.TextFrame.TextRange.Font.Size=28;added.Select();}catch{if(added!=null)added.Delete();throw;}
        }
        private static void ShowSymbols(PowerPoint.Application app)
        {
            using(var form=new Form{Text="Insert Unicode symbol",Width=600,Height=550,StartPosition=FormStartPosition.CenterScreen})
            {
                var search=new TextBox{Dock=DockStyle.Top,AccessibleName="Search symbol name or Unicode code"};var list=new ListBox{Dock=DockStyle.Fill};var insert=new Button{Dock=DockStyle.Bottom,Text="Insert selected symbol",DialogResult=DialogResult.OK,Height=36};var note=new Label{Dock=DockStyle.Bottom,Text="Insert at selected text caret, replace selected text, or create a new text box. Glyph appearance depends on your font.",Height=38};
                Action update=()=>{list.Items.Clear();foreach(var entry in SearchSymbols(search.Text))list.Items.Add(entry);if(list.Items.Count>0)list.SelectedIndex=0;insert.Enabled=list.Items.Count>0;};search.MaxLength=100;search.TextChanged+=(s,e)=>update();form.Controls.Add(list);form.Controls.Add(search);form.Controls.Add(note);form.Controls.Add(insert);form.AcceptButton=insert;update();
                if(form.ShowDialog()==DialogResult.OK&&list.SelectedItem is ProductivitySymbol)InsertSymbol(app,((ProductivitySymbol)list.SelectedItem).Glyph);
            }
        }
        private sealed class TextEdit {internal PowerPoint.TextRange Range;internal string Original;internal DecimalCharacterChange[] Changes;}
        private static void CollectText(PowerPoint.Shape shape,List<PowerPoint.TextRange> ranges,int depth,ref int visited)
        {
            if(++visited>MaxTargets||depth>8)throw new InvalidOperationException("Select at most 100 shapes including nested group members (depth 8).");
            if(shape.Type==Office.MsoShapeType.msoGroup){for(int j=1;j<=shape.GroupItems.Count;j++)CollectText(shape.GroupItems[j],ranges,depth+1,ref visited);return;}
            if(shape.HasTextFrame==Office.MsoTriState.msoTrue&&shape.TextFrame.HasText==Office.MsoTriState.msoTrue)ranges.Add(shape.TextFrame.TextRange);
        }
        internal static int SwitchDecimals(PowerPoint.Application app,bool confirm=true)
        {
            ActiveSlide(app);var selection=app.ActiveWindow.Selection;var ranges=new List<PowerPoint.TextRange>();
            if(selection.Type==PowerPoint.PpSelectionType.ppSelectionText)ranges.Add(selection.TextRange);
            else if(selection.Type==PowerPoint.PpSelectionType.ppSelectionShapes){int visited=0;for(int i=1;i<=selection.ShapeRange.Count;i++)CollectText(selection.ShapeRange[i],ranges,0,ref visited);}
            else throw new InvalidOperationException("Select text or text-bearing shapes first.");
            var plans=new List<TextEdit>();int chars=0;foreach(var range in ranges){if(range.Length>MaxCharacters||chars+range.Length>MaxCharacters)throw new InvalidOperationException("Selection exceeds 20,000 text characters.");string original=range.Text;chars+=original.Length;plans.Add(new TextEdit{Range=range,Original=original,Changes=DecimalChanges(original)});}
            int total=plans.Sum(p=>p.Changes.Length);if(total==0){if(confirm)MessageBox.Show("No digit-adjacent dots or commas were found. Table cells and chart labels are not included.","Switch decimal symbols");return 0;}
            if(confirm&&MessageBox.Show("Swap "+total+" digit-adjacent dots/commas in "+plans.Count+" selected text ranges? This also affects digit-separated dates and thousands separators. Tables and chart data are excluded. Character formatting is retained; use Undo to revert.","Review text substitution",MessageBoxButtons.OKCancel)!=DialogResult.OK)return 0;
            app.StartNewUndoEntry();var applied=new List<Tuple<TextEdit,DecimalCharacterChange>>();
            try{foreach(var plan in plans){if(plan.Range.Text!=plan.Original)throw new InvalidOperationException("Selected text changed; run again.");foreach(var change in plan.Changes){plan.Range.Characters(change.Index+1,1).Text=change.After.ToString();applied.Add(Tuple.Create(plan,change));}}}
            catch(Exception original){bool restored=true;for(int i=applied.Count-1;i>=0;i--)try{applied[i].Item1.Range.Characters(applied[i].Item2.Index+1,1).Text=applied[i].Item2.Before.ToString();}catch{restored=false;}if(!restored)throw new InvalidOperationException("Substitution failed; rollback incomplete. Use PowerPoint Undo.",original);throw;}return total;
        }
        internal static PowerPoint.Shape RasterCopy(PowerPoint.Application app)
        {
            var slide=ActiveSlide(app);var selection=app.ActiveWindow.Selection;if(selection.Type!=PowerPoint.PpSelectionType.ppSelectionShapes||selection.ShapeRange.Count!=1)throw new InvalidOperationException("Select one native object or one group. The original will be retained.");
            var source=selection.ShapeRange[1];if(Math.Abs(source.Rotation%360)>.01f)throw new InvalidOperationException("Raster copy currently supports unrotated objects only; rotation bounds are not approximated.");
            if(source.Width<=0||source.Height<=0||source.Width>4000||source.Height>4000)throw new InvalidOperationException("Choose an object between 0 and 4,000 points in each dimension.");
            int width=Math.Max(1,(int)Math.Round(source.Width*2)),height=Math.Max(1,(int)Math.Round(source.Height*2));if(width>4096||height>4096)throw new InvalidOperationException("Raster copy is limited to 4096 pixels per side at 2 pixels per point.");
            string directory=Path.Combine(Path.GetTempPath(),"ppi-raster-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(directory);string file=Path.Combine(directory,"object.png");PowerPoint.Shape picture=null;
            try{source.Export(file,PowerPoint.PpShapeFormat.ppShapeFormatPNG,(int)Math.Round(app.ActivePresentation.PageSetup.SlideWidth*1.5),(int)Math.Round(app.ActivePresentation.PageSetup.SlideHeight*1.5),PowerPoint.PpExportMode.ppRelativeToSlide);if(!File.Exists(file)||new FileInfo(file).Length>64L*1024*1024)throw new InvalidDataException("Raster export was missing or exceeded 64 MB.");using(var bitmap=System.Drawing.Image.FromFile(file)){if(bitmap.Width>4096||bitmap.Height>4096)throw new InvalidDataException("Exported image exceeded 4096 pixels per side.");}app.StartNewUndoEntry();picture=slide.Shapes.AddPicture(file,Office.MsoTriState.msoFalse,Office.MsoTriState.msoTrue,source.Left,source.Top,source.Width,source.Height);picture.Name="Intelligence raster copy "+picture.Id;picture.AlternativeText="Raster copy; original editable object retained.";picture.Select();return picture;}
            catch{if(picture!=null)picture.Delete();throw;}
            finally{try{if(File.Exists(file))File.Delete(file);if(Directory.Exists(directory))Directory.Delete(directory);}catch(Exception cleanup) when(cleanup is IOException||cleanup is UnauthorizedAccessException){MessageBox.Show("The temporary raster export could not be removed: "+directory,"Raster export cleanup");}}
        }
        private sealed class WorkbookItem
        {
            internal int SlideId,SlideIndex,ShapeId;internal string Label;
            public override string ToString(){return Label;}
        }
        private static void ShowWorkbooks(PowerPoint.Application app)
        {
            var deck=app.ActivePresentation;
            using(var form=new Form{Text="Native chart workbooks — separate embedded workbooks",Width=700,Height=500,StartPosition=FormStartPosition.CenterScreen})
            {
                var list=new ListBox{Dock=DockStyle.Fill};var open=new Button{Text="Open selected chart's Edit Data workbook",Dock=DockStyle.Bottom,Height=40};var refresh=new Button{Text="Refresh chart list",Dock=DockStyle.Top,Height=32};var note=new Label{Dock=DockStyle.Bottom,Height=42,Text="Top-level native charts only; maximum 500 slides / 2,000 shapes. Each chart keeps its own workbook. No shared slide workbook is created."};
                Action reload=()=>{list.Items.Clear();int objects=0;bool limited=deck.Slides.Count>500;for(int i=1;i<=Math.Min(500,deck.Slides.Count);i++){var slide=deck.Slides[i];for(int j=1;j<=slide.Shapes.Count;j++){if(++objects>2000){limited=true;break;}var shape=slide.Shapes[j];if(shape.HasChart==Office.MsoTriState.msoTrue)list.Items.Add(new WorkbookItem{SlideId=slide.SlideID,SlideIndex=i,ShapeId=shape.Id,Label="Slide "+i+" · chart "+shape.Id+" · "+shape.Name});}if(objects>2000)break;}if(list.Items.Count>0)list.SelectedIndex=0;open.Enabled=list.Items.Count>0;note.Text=(limited?"INCOMPLETE LIST. ":"")+"Top-level native charts; separate embedded workbooks. No shared slide workbook. Chart list is bounded to 500 slides / 2,000 shapes.";};
                refresh.Click+=(s,e)=>{try{reload();}catch(Exception ex){note.Text="Refresh unavailable ("+ex.GetType().Name+"). Presentation may have closed.";}};
                open.Click+=(s,e)=>{var item=list.SelectedItem as WorkbookItem;if(item==null)return;try{deck.Windows[1].Activate();var slide=deck.Slides.FindBySlideID(item.SlideId);slide.Select();PowerPoint.Shape target=null;for(int j=1;j<=slide.Shapes.Count;j++)if(slide.Shapes[j].Id==item.ShapeId){target=slide.Shapes[j];break;}if(target==null||target.HasChart!=Office.MsoTriState.msoTrue)throw new InvalidOperationException("Chart changed; refresh the list.");target.Select();target.Chart.ChartData.Activate();}catch(Exception ex){note.Text="Workbook unavailable ("+ex.GetType().Name+"). Refresh after document changes.";}};
                form.Controls.Add(list);form.Controls.Add(refresh);form.Controls.Add(note);form.Controls.Add(open);reload();form.ShowDialog();
            }
        }
    }
}

