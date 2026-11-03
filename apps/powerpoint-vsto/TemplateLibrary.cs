using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml;
using Office=Microsoft.Office.Core;
using PowerPoint=Microsoft.Office.Interop.PowerPoint;
namespace PowerPointIntelligence.AddIn
{
    internal static class TemplatePreview
    {
        internal static Color ColorOf(string value){return ColorTranslator.FromHtml(value??"#FFFFFF");}
        internal static Bitmap Render(SlideTemplate template,int width,int height,IDictionary<string,string> fields=null,string paletteId=null,string horizontal="Center",string vertical="Middle",string highlight=null)
        {
            TemplateAlignment.Validate(horizontal,vertical);
            var image=new Bitmap(width,height,PixelFormat.Format32bppArgb);
            using(var g=Graphics.FromImage(image))
            {
                g.Clear(ColorOf(TemplatePalettes.Get(paletteId).Colors[9]));g.SmoothingMode=SmoothingMode.AntiAlias;
                float scale=Math.Min(width/960f,height/540f);g.TranslateTransform((width-960*scale)/2,(height-540*scale)/2);g.ScaleTransform(scale,scale);
                foreach(var node in template.Build(paletteId))
                {
                    if(node.Kind=="line"){using(var pen=new Pen(ColorOf(node.Stroke),1.5f))g.DrawLine(pen,node.Points[0],node.Points[1],node.Points[2],node.Points[3]);continue;}
                    if(node.Kind=="text")
                    {
                        string text=node.Text;if(fields!=null&&fields.ContainsKey(node.Key))text=fields[node.Key];
                        if(!String.IsNullOrWhiteSpace(highlight)&&Regex.Split(highlight.Trim(),@"\s+").Any(term=>text.IndexOf(term,StringComparison.OrdinalIgnoreCase)>=0))using(var mark=new SolidBrush(Color.FromArgb(90,255,213,79)))g.FillRectangle(mark,node.X,node.Y,node.W,node.H);
                        using(var font=new Font("Segoe UI",node.Font,node.Bold?FontStyle.Bold:FontStyle.Regular,GraphicsUnit.Pixel))
                        using(var brush=new SolidBrush(ColorOf(node.Fill)))
                        using(var format=new StringFormat {Trimming=StringTrimming.EllipsisWord,Alignment=horizontal=="Left"?StringAlignment.Near:horizontal=="Right"?StringAlignment.Far:StringAlignment.Center,LineAlignment=vertical=="Top"?StringAlignment.Near:vertical=="Bottom"?StringAlignment.Far:StringAlignment.Center})g.DrawString(text,font,brush,new RectangleF(node.X,node.Y,node.W,node.H),format);
                        continue;
                    }
                    using(var brush=new SolidBrush(ColorOf(node.Fill)))using(var pen=new Pen(ColorOf(node.Stroke),1))
                    {
                        var rect=new RectangleF(node.X,node.Y,node.W,node.H);
                        if(node.Kind=="ellipse"){g.FillEllipse(brush,rect);g.DrawEllipse(pen,rect);}
                        else if(node.Kind=="polygon"||node.Kind=="chevron"||node.Kind=="diamond")
                        {
                            float[] p=node.Points;
                            if(node.Kind=="chevron")p=new[]{node.X,node.Y,node.X+node.W*.8f,node.Y,node.X+node.W,node.Y+node.H/2,node.X+node.W*.8f,node.Y+node.H,node.X,node.Y+node.H,node.X+node.W*.2f,node.Y+node.H/2};
                            if(node.Kind=="diamond")p=new[]{node.X+node.W/2,node.Y,node.X+node.W,node.Y+node.H/2,node.X+node.W/2,node.Y+node.H,node.X,node.Y+node.H/2};
                            var points=Enumerable.Range(0,p.Length/2).Select(i=>new PointF(p[i*2],p[i*2+1])).ToArray();g.FillPolygon(brush,points);g.DrawPolygon(pen,points);
                        }
                        else {g.FillRectangle(brush,rect);g.DrawRectangle(pen,node.X,node.Y,node.W,node.H);}
                    }
                }
            }
            return image;
        }
    }
    internal static class TemplateInserter
    {
        internal const string TemplateTag="PPI_TEMPLATE_ID",FieldTag="PPI_TEMPLATE_FIELD";
        internal static PowerPoint.Slide Insert(PowerPoint.Application app,SlideTemplate template,IDictionary<string,string> fields,bool builds,string transition,string buildStyle="Fade",string paletteId=null,string horizontal="Center",string vertical="Middle",bool objectsOnly=false,bool destinationTheme=false)
        {
            if(app==null||app.Presentations.Count==0)throw new InvalidOperationException("Open a presentation before inserting a template.");
            if(objectsOnly&&builds)throw new InvalidOperationException("Object-only insertion currently supports static groups. Use new-slide insertion for staged builds.");
            if(!TemplateCatalog.Items.Any(t=>t.Id==template.Id))throw new InvalidDataException("Unknown template.");
            if(transition!="None"&&transition!="Fade"&&transition!="Wipe")throw new InvalidDataException("Unknown transition.");
            if(buildStyle!="Fade"&&buildStyle!="Wipe"&&buildStyle!="Fly in")throw new InvalidDataException("Unknown animation build style.");
            TemplateAlignment.Validate(horizontal,vertical);
            var palette=TemplatePalettes.Get(paletteId);var nodes=template.Build(palette.Id);var keys=new HashSet<string>(nodes.Where(n=>n.Kind=="text").Select(n=>n.Key));
            if(fields!=null&&fields.Any(f=>!keys.Contains(f.Key)||f.Value==null||f.Value.Length>2000))throw new InvalidDataException("Editable fields must be known and no longer than 2000 characters.");
            var deck=app.ActivePresentation;var active=app.ActiveWindow.View.Slide as PowerPoint.Slide;
            if(active==null)throw new InvalidOperationException("Select a slide in Normal view first.");
            float scale=Math.Min(deck.PageSetup.SlideWidth/960f,deck.PageSetup.SlideHeight/540f),ox=(deck.PageSetup.SlideWidth-960*scale)/2,oy=(deck.PageSetup.SlideHeight-540*scale)/2;
            app.StartNewUndoEntry();PowerPoint.Slide slide=null;var created=new List<PowerPoint.Shape>();PowerPoint.Shape instanceGroup=null;string instance=Guid.NewGuid().ToString("N");
            try
            {
                slide=objectsOnly?active:deck.Slides.Add(active.SlideIndex+1,PowerPoint.PpSlideLayout.ppLayoutBlank);
                if(!objectsOnly){slide.Design=active.Design;slide.FollowMasterBackground=Office.MsoTriState.msoFalse;slide.Background.Fill.ForeColor.RGB=Rgb(palette.Colors[9]);slide.Tags.Add("PPI_TEMPLATE_PALETTE",palette.Id);slide.Tags.Add(TemplateTag,template.Id);slide.Tags.Add("PPI_TEMPLATE_VERSION","1");if(destinationTheme)slide.FollowMasterBackground=Office.MsoTriState.msoTrue;}
                var stages=new SortedDictionary<int,List<string>>();int index=0;
                foreach(var node in nodes)
                {
                    PowerPoint.Shape shape;
                    if(node.Kind=="line")shape=slide.Shapes.AddLine(ox+node.Points[0]*scale,oy+node.Points[1]*scale,ox+node.Points[2]*scale,oy+node.Points[3]*scale);
                    else if(node.Kind=="polygon")
                    {
                        var path=slide.Shapes.BuildFreeform(Office.MsoEditingType.msoEditingCorner,ox+node.Points[0]*scale,oy+node.Points[1]*scale);
                        for(int i=2;i<node.Points.Length;i+=2)path.AddNodes(Office.MsoSegmentType.msoSegmentLine,Office.MsoEditingType.msoEditingCorner,ox+node.Points[i]*scale,oy+node.Points[i+1]*scale);
                        path.AddNodes(Office.MsoSegmentType.msoSegmentLine,Office.MsoEditingType.msoEditingCorner,ox+node.Points[0]*scale,oy+node.Points[1]*scale);shape=path.ConvertToShape();
                    }
                    else if(node.Kind=="text")shape=slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,ox+node.X*scale,oy+node.Y*scale,node.W*scale,node.H*scale);
                    else {var kind=node.Kind=="ellipse"?Office.MsoAutoShapeType.msoShapeOval:node.Kind=="round"?Office.MsoAutoShapeType.msoShapeRoundedRectangle:node.Kind=="chevron"?Office.MsoAutoShapeType.msoShapeChevron:node.Kind=="diamond"?Office.MsoAutoShapeType.msoShapeDiamond:Office.MsoAutoShapeType.msoShapeRectangle;shape=slide.Shapes.AddShape(kind,ox+node.X*scale,oy+node.Y*scale,node.W*scale,node.H*scale);}
                    created.Add(shape);shape.Name="ppi_template_"+instance+"_"+(++index);shape.Tags.Add("PPI_TEMPLATE_INSTANCE",instance);shape.Tags.Add("PPI_TEMPLATE_NODE",index.ToString(CultureInfo.InvariantCulture));shape.Tags.Add(TemplateTag,template.Id);
                    if(node.Kind=="text")
                    {
                        shape.Tags.Add(FieldTag,node.Key);shape.Fill.Visible=Office.MsoTriState.msoFalse;shape.Line.Visible=Office.MsoTriState.msoFalse;
                        var frame=shape.TextFrame;frame.MarginLeft=frame.MarginRight=frame.MarginTop=frame.MarginBottom=0;frame.WordWrap=Office.MsoTriState.msoTrue;frame.VerticalAnchor=Anchor(vertical);
                        var range=frame.TextRange;range.ParagraphFormat.Alignment=Paragraph(horizontal);range.Text=fields!=null&&fields.ContainsKey(node.Key)?fields[node.Key]:node.Text;range.Font.Name="Segoe UI";range.Font.Size=node.Font*scale;range.Font.Bold=node.Bold?Office.MsoTriState.msoTrue:Office.MsoTriState.msoFalse;range.Font.Color.RGB=Rgb(node.Fill);
                        shape.TextFrame2.AutoSize=Office.MsoAutoSize.msoAutoSizeTextToFitShape;frame.VerticalAnchor=Anchor(vertical);range.ParagraphFormat.Alignment=Paragraph(horizontal);shape.AlternativeText="Editable field: "+node.Key;
                    }
                    else {if(node.Kind!="line"){shape.Fill.Solid();shape.Fill.ForeColor.RGB=Rgb(node.Fill);}shape.Line.ForeColor.RGB=Rgb(node.Stroke??"#FFFFFF");shape.Line.Weight=scale;}
                    if(destinationTheme)ApplyTheme(shape,node,palette);
                    if(node.Stage>0){if(!stages.ContainsKey(node.Stage))stages[node.Stage]=new List<string>();stages[node.Stage].Add(shape.Name);}
                }
                if(builds)foreach(var stage in stages)
                {
                    var shape=stage.Value.Count==1?slide.Shapes[stage.Value[0]]:slide.Shapes.Range(stage.Value.ToArray()).Group();
                    shape.Tags.Add("PPI_TEMPLATE_INSTANCE",instance);shape.Tags.Add(TemplateTag,template.Id);shape.Tags.Add("PPI_TEMPLATE_STAGE",stage.Key.ToString(CultureInfo.InvariantCulture));
                    var effect=slide.TimeLine.MainSequence.AddEffect(shape,(buildStyle=="Wipe"?PowerPoint.MsoAnimEffect.msoAnimEffectWipe:buildStyle=="Fly in"?PowerPoint.MsoAnimEffect.msoAnimEffectFly:PowerPoint.MsoAnimEffect.msoAnimEffectFade),PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone,PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);effect.Timing.Duration=.35f;
                }
                if(!objectsOnly){slide.SlideShowTransition.EntryEffect=transition=="Fade"?PowerPoint.PpEntryEffect.ppEffectFadeSmoothly:transition=="Wipe"?PowerPoint.PpEntryEffect.ppEffectWipeRight:PowerPoint.PpEntryEffect.ppEffectNone;
                slide.SlideShowTransition.AdvanceOnClick=Office.MsoTriState.msoTrue;slide.SlideShowTransition.AdvanceOnTime=Office.MsoTriState.msoFalse;}
                var xml=new XmlDocument {XmlResolver=null};var root=xml.CreateElement("Template","urn:powerpoint-intelligence:template:v1");xml.AppendChild(root);root.SetAttribute("id",template.Id);root.SetAttribute("version","1");root.SetAttribute("builds",builds?"true":"false");root.SetAttribute("transition",transition);root.SetAttribute("buildStyle",buildStyle);root.SetAttribute("palette",palette.Id);root.SetAttribute("horizontal",horizontal);root.SetAttribute("vertical",vertical);root.SetAttribute("instance",instance);root.SetAttribute("objectsOnly",objectsOnly?"true":"false");root.SetAttribute("formatting",destinationTheme?"destination":"palette");
                foreach(var node in nodes.Where(n=>n.Kind=="text")){var field=xml.CreateElement("Field",root.NamespaceURI);field.SetAttribute("key",node.Key);field.InnerText=fields!=null&&fields.ContainsKey(node.Key)?fields[node.Key]:node.Text;root.AppendChild(field);}
                // Attach metadata to this owned slide, so copying slides preserves it.
                var anchor=created[0];if(!anchor.CustomerData.Add().LoadXML(xml.OuterXml))throw new InvalidDataException("Template metadata was rejected.");
                if(objectsOnly){var top=new List<string>();foreach(PowerPoint.Shape shape in slide.Shapes)if(String.Equals(shape.Tags["PPI_TEMPLATE_INSTANCE"],instance,StringComparison.OrdinalIgnoreCase))top.Add(shape.Name);instanceGroup=slide.Shapes.Range(top.ToArray()).Group();instanceGroup.Tags.Add(TemplateTag,template.Id);instanceGroup.Tags.Add("PPI_TEMPLATE_INSTANCE",instance);instanceGroup.Tags.Add("PPI_TEMPLATE_PALETTE",palette.Id);instanceGroup.Tags.Add("PPI_TEMPLATE_OBJECTS","1");instanceGroup.Select();}else slide.Select();return slide;
            }
            catch(Exception original) {bool restored=true;if(slide!=null&&!objectsOnly){try{slide.Delete();}catch{restored=false;}}else if(slide!=null){try{if(instanceGroup!=null)instanceGroup.Delete();else {var remove=new List<PowerPoint.Shape>();foreach(PowerPoint.Shape shape in slide.Shapes)if(String.Equals(shape.Tags["PPI_TEMPLATE_INSTANCE"],instance,StringComparison.OrdinalIgnoreCase))remove.Add(shape);foreach(var shape in remove)shape.Delete();}}catch{restored=false;}}if(!restored)throw new InvalidOperationException("Template insertion failed and cleanup was incomplete. Use PowerPoint Undo.",original);throw;}
        }
        private static void ApplyTheme(PowerPoint.Shape shape,TemplateNode node,TemplatePalette palette)
        {
            if(node.Kind=="text"){shape.TextFrame.TextRange.Font.Name=node.Key=="title"?"+mj-lt":"+mn-lt";ThemeColor(shape.TextFrame.TextRange.Font.Color,node.Fill,palette,true);}
            else {if(node.Kind!="line")ThemeColor(shape.Fill.ForeColor,node.Fill,palette,false);ThemeColor(shape.Line.ForeColor,node.Stroke,palette,false);}
        }
        private static void ThemeColor(PowerPoint.ColorFormat color,string value,TemplatePalette palette,bool text)
        {
            var binding=TemplateTheme.Binding(value,palette,text);color.ObjectThemeColor=(Office.MsoThemeColorIndex)binding.Item1;color.TintAndShade=binding.Item2;
        }
        private static PowerPoint.PpParagraphAlignment Paragraph(string value){return value=="Left"?PowerPoint.PpParagraphAlignment.ppAlignLeft:value=="Right"?PowerPoint.PpParagraphAlignment.ppAlignRight:value=="Justify"?PowerPoint.PpParagraphAlignment.ppAlignJustify:PowerPoint.PpParagraphAlignment.ppAlignCenter;}
        private static Office.MsoVerticalAnchor Anchor(string value){return value=="Top"?Office.MsoVerticalAnchor.msoAnchorTop:value=="Bottom"?Office.MsoVerticalAnchor.msoAnchorBottom:Office.MsoVerticalAnchor.msoAnchorMiddle;}
        internal static void ChangeAlignment(PowerPoint.Application app,string horizontal,string vertical)
        {
            TemplateAlignment.Validate(horizontal,vertical);var slide=app.ActiveWindow.View.Slide as PowerPoint.Slide;
            if(slide==null)throw new InvalidOperationException("Select a library slide first.");
            var scope=SelectedInstance(app);string templateId=scope==null?slide.Tags[TemplateTag]:scope.Tags[TemplateTag];
            var template=TemplateCatalog.Items.SingleOrDefault(t=>String.Equals(t.Id,templateId,StringComparison.OrdinalIgnoreCase));
            if(template==null)throw new InvalidOperationException("Select an inserted library template slide first.");
            var targets=new List<Tuple<PowerPoint.Shape,TemplateNode>>();
            var selection=app.ActiveWindow.Selection;
            if(selection.Type==PowerPoint.PpSelectionType.ppSelectionShapes){foreach(PowerPoint.Shape shape in selection.ShapeRange)Collect(shape,template.Id,template.Build(),targets);}else CollectAll(slide,template,targets);
            if(!targets.Any(t=>t.Item2.Kind=="text"))throw new InvalidOperationException("Select library text or a template group, or deselect shapes to align the whole slide.");
            var restore=new List<Action>();app.StartNewUndoEntry();
            try {foreach(var target in targets.Where(t=>t.Item2.Kind=="text")){
                var frame=target.Item1.TextFrame;var oldAnchor=frame.VerticalAnchor;restore.Add(()=>frame.VerticalAnchor=oldAnchor);
                var range=frame.TextRange;for(int i=1;i<=range.Paragraphs().Count;i++){var paragraph=range.Paragraphs(i,1);var old=paragraph.ParagraphFormat.Alignment;restore.Add(()=>paragraph.ParagraphFormat.Alignment=old);}
                range.ParagraphFormat.Alignment=Paragraph(horizontal);frame.VerticalAnchor=Anchor(vertical);
            }}catch(Exception original){bool restored=true;for(int i=restore.Count-1;i>=0;i--)try{restore[i]();}catch{restored=false;}if(!restored)throw new InvalidOperationException("Alignment was interrupted and rollback was incomplete. Use PowerPoint Undo.",original);throw;}

        }
        private static void CollectAll(PowerPoint.Slide slide,SlideTemplate template,List<Tuple<PowerPoint.Shape,TemplateNode>> targets){var nodes=template.Build();foreach(PowerPoint.Shape shape in slide.Shapes)Collect(shape,template.Id,nodes,targets);}
        internal static void ChangePalette(PowerPoint.Application app,string paletteId)
        {
            var palette=TemplatePalettes.Get(paletteId);var slide=app.ActiveWindow.View.Slide as PowerPoint.Slide;
            if(slide==null)throw new InvalidOperationException("Select an inserted template slide first.");
            var scope=SelectedInstance(app);string templateId=scope==null?slide.Tags[TemplateTag]:scope.Tags[TemplateTag];
            var template=TemplateCatalog.Items.SingleOrDefault(t=>String.Equals(t.Id,templateId,StringComparison.OrdinalIgnoreCase));
            if(template==null)throw new InvalidOperationException("This slide is not a library template. Select a library slide first.");
            var nodes=template.Build(palette.Id);var targets=new List<Tuple<PowerPoint.Shape,TemplateNode>>();
            if(scope!=null)Collect(scope,template.Id,nodes,targets);else foreach(PowerPoint.Shape shape in slide.Shapes)Collect(shape,template.Id,nodes,targets);
            if(targets.Count==0)throw new InvalidOperationException("No retained library objects were found on this slide.");
            var rollback=new List<Action>();int oldBackground=slide.Background.Fill.ForeColor.RGB;string oldPalette=slide.Tags["PPI_TEMPLATE_PALETTE"];
            app.StartNewUndoEntry();
            try
            {
                foreach(var target in targets)
                {
                    var shape=target.Item1;var node=target.Item2;
                    if(node.Kind=="text")
                    {
                        // Retain individual rich-text run colors if a later operation fails.
                        var range=shape.TextFrame.TextRange;for(int i=1;i<=range.Length;i++){var run=range.Characters(i,1);int original=run.Font.Color.RGB;rollback.Add(()=>run.Font.Color.RGB=original);}
                        range.Font.Color.RGB=Rgb(node.Fill);
                    }
                    else {if(node.Kind!="line"){int original=shape.Fill.ForeColor.RGB;rollback.Add(()=>shape.Fill.ForeColor.RGB=original);shape.Fill.ForeColor.RGB=Rgb(node.Fill);}int lineColor=shape.Line.ForeColor.RGB;rollback.Add(()=>shape.Line.ForeColor.RGB=lineColor);shape.Line.ForeColor.RGB=Rgb(node.Stroke??palette.Colors[9]);}
                }
                if(scope==null){rollback.Add(()=>slide.Background.Fill.ForeColor.RGB=oldBackground);slide.Background.Fill.ForeColor.RGB=Rgb(palette.Colors[9]);
                rollback.Add(()=>{if(String.IsNullOrEmpty(oldPalette))slide.Tags.Delete("PPI_TEMPLATE_PALETTE");else slide.Tags.Add("PPI_TEMPLATE_PALETTE",oldPalette);});slide.Tags.Add("PPI_TEMPLATE_PALETTE",palette.Id);}else {string previous=scope.Tags["PPI_TEMPLATE_PALETTE"];rollback.Add(()=>scope.Tags.Add("PPI_TEMPLATE_PALETTE",previous));scope.Tags.Add("PPI_TEMPLATE_PALETTE",palette.Id);}
                // Update only palette metadata; never replace insertion-time text snapshots.
                foreach(var target in targets)foreach(Office.CustomXMLPart part in target.Item1.CustomerData)if(part.NamespaceURI=="urn:powerpoint-intelligence:template:v1")
                {
                    var attribute=part.DocumentElement.SelectSingleNode("@palette");
                    if(attribute!=null){string previous=attribute.NodeValue;rollback.Add(()=>attribute.NodeValue=previous);attribute.NodeValue=palette.Id;}
                    else {part.DocumentElement.AppendChildNode("palette","",Office.MsoCustomXMLNodeType.msoCustomXMLNodeAttribute,palette.Id);var created=part.DocumentElement.SelectSingleNode("@palette");if(created==null)throw new InvalidDataException("Palette attribute was not created.");rollback.Add(()=>created.Delete());}
                }
            }
            catch(Exception original)
            {
                bool restored=true;for(int i=rollback.Count-1;i>=0;i--)try{rollback[i]();}catch{restored=false;}
                if(!restored)throw new InvalidOperationException("Palette change was interrupted and rollback was incomplete. Use PowerPoint Undo and check this slide.",original);
                throw;
            }
        }
        private static PowerPoint.Shape SelectedInstance(PowerPoint.Application app)
        {
            var selection=app.ActiveWindow.Selection;
            if(selection.Type!=PowerPoint.PpSelectionType.ppSelectionShapes)return null;
            if(selection.ShapeRange.Count==1){var shape=selection.ShapeRange[1];if(!String.IsNullOrEmpty(shape.Tags[TemplateTag]))return shape;}
            // Multiple separately inserted instances require an explicit group selection.
            if(selection.ShapeRange.Cast<PowerPoint.Shape>().Any(s=>s.Tags["PPI_TEMPLATE_OBJECTS"]=="1"))throw new InvalidOperationException("Select one library template group at a time.");return null;
        }
        private static void Collect(PowerPoint.Shape shape,string id,List<TemplateNode> nodes,List<Tuple<PowerPoint.Shape,TemplateNode>> targets)
        {
            if(shape.Type==Office.MsoShapeType.msoGroup){foreach(PowerPoint.Shape child in shape.GroupItems)Collect(child,id,nodes,targets);return;}
            if(!String.Equals(shape.Tags[TemplateTag],id,StringComparison.OrdinalIgnoreCase))return;
            int index;string token=shape.Tags["PPI_TEMPLATE_NODE"];
            if(String.IsNullOrEmpty(token)&&shape.Name.StartsWith("ppi_template_",StringComparison.Ordinal))token=shape.Name.Substring("ppi_template_".Length);
            if(Int32.TryParse(token,NumberStyles.None,CultureInfo.InvariantCulture,out index)&&index>0&&index<=nodes.Count)targets.Add(Tuple.Create(shape,nodes[index-1]));
        }
        private static int Rgb(string color){var c=TemplatePreview.ColorOf(color);return c.R|(c.G<<8)|(c.B<<16);}
    }
    internal static class AlignmentPicker
    {
        internal static string[] Choose()
        {
            using(var form=new Form {Text="Template text alignment",Width=410,Height=190,StartPosition=FormStartPosition.CenterScreen})
            using(var layout=new FlowLayoutPanel {Dock=DockStyle.Fill,Padding=new Padding(16)})
            using(var horizontal=new ComboBox {DropDownStyle=ComboBoxStyle.DropDownList,Width=160})
            using(var vertical=new ComboBox {DropDownStyle=ComboBoxStyle.DropDownList,Width=160})
            using(var apply=new Button {Text="Apply",DialogResult=DialogResult.OK})
            using(var cancel=new Button {Text="Cancel",DialogResult=DialogResult.Cancel})
            {horizontal.Items.AddRange(TemplateAlignment.Horizontal);horizontal.SelectedItem="Center";vertical.Items.AddRange(TemplateAlignment.Vertical);vertical.SelectedItem="Middle";layout.Controls.Add(new Label {Text="Horizontal / vertical alignment",AutoSize=true,Width=330});layout.Controls.Add(horizontal);layout.Controls.Add(vertical);layout.Controls.Add(apply);layout.Controls.Add(cancel);form.Controls.Add(layout);form.AcceptButton=apply;form.CancelButton=cancel;return form.ShowDialog()==DialogResult.OK?new[]{(string)horizontal.SelectedItem,(string)vertical.SelectedItem}:null;}
        }
    }
    internal static class PalettePicker
    {
        internal static void Configure(ComboBox combo)
        {
            combo.DropDownStyle=ComboBoxStyle.DropDownList;combo.DrawMode=DrawMode.OwnerDrawFixed;combo.ItemHeight=28;combo.DropDownWidth=440;
            foreach(var item in TemplatePalettes.Items)combo.Items.Add(item);
            combo.DrawItem+=(sender,args)=>{args.DrawBackground();if(args.Index<0)return;var palette=(TemplatePalette)combo.Items[args.Index];using(var brush=new SolidBrush(args.ForeColor))args.Graphics.DrawString(palette.Name,combo.Font,brush,args.Bounds.X+4,args.Bounds.Y+5);if(args.Bounds.Width>240)for(int c=0;c<10;c++)using(var brush=new SolidBrush(TemplatePreview.ColorOf(palette.Colors[c])))args.Graphics.FillRectangle(brush,args.Bounds.Right-190+c*18,args.Bounds.Y+5,17,18);args.DrawFocusRectangle();};
        }
        internal static string Choose(string current)
        {
            using(var form=new Form {Text="Change template color palette",Width=510,Height=170,StartPosition=FormStartPosition.CenterScreen})
            using(var combo=new ComboBox {Dock=DockStyle.Top})
            using(var label=new Label {Text="Apply to library objects on the current slide. Text, layout and animations are retained.",Dock=DockStyle.Top,Height=40})
            using(var buttons=new FlowLayoutPanel {Dock=DockStyle.Bottom,Height=40,FlowDirection=FlowDirection.RightToLeft})
            using(var apply=new Button {Text="Apply palette",DialogResult=DialogResult.OK,AutoSize=true})
            using(var cancel=new Button {Text="Cancel",DialogResult=DialogResult.Cancel})
            {Configure(combo);combo.SelectedIndex=Math.Max(0,TemplatePalettes.Items.ToList().FindIndex(p=>p.Id==current));buttons.Controls.Add(cancel);buttons.Controls.Add(apply);form.Controls.Add(combo);form.Controls.Add(label);form.Controls.Add(buttons);form.AcceptButton=apply;form.CancelButton=cancel;return form.ShowDialog()==DialogResult.OK?((TemplatePalette)combo.SelectedItem).Id:null;}
        }
    }
    internal sealed class TemplateLibrary : Form
    {
        private readonly ListBox categories=new ListBox {Dock=DockStyle.Fill};
        private readonly TextBox search=new TextBox {Dock=DockStyle.Top,AccessibleName="Search template library"};
        private readonly ListView results=new ListView {Dock=DockStyle.Fill,View=View.LargeIcon,MultiSelect=false,HideSelection=false};
        private readonly ImageList thumbnails=new ImageList {ImageSize=new Size(192,108),ColorDepth=ColorDepth.Depth32Bit};
        private readonly PictureBox preview=new PictureBox {Dock=DockStyle.Top,Height=240,SizeMode=PictureBoxSizeMode.Zoom,BackColor=Color.White};
        private readonly DataGridView fields=new DataGridView {Dock=DockStyle.Fill,AllowUserToAddRows=false,AllowUserToDeleteRows=false,RowHeadersVisible=false,AutoSizeColumnsMode=DataGridViewAutoSizeColumnsMode.Fill};
        private readonly Label status=new Label {Dock=DockStyle.Bottom,Height=48};
        private readonly Label pageLabel=new Label {AutoSize=true,Padding=new Padding(8)};
        private readonly CheckBox builds=new CheckBox {Text="Click to reveal each stage",AutoSize=true};
        private readonly ComboBox buildStyle=new ComboBox {DropDownStyle=ComboBoxStyle.DropDownList,Width=100};
        private readonly ComboBox palette=new ComboBox {DropDownStyle=ComboBoxStyle.DropDownList,Width=190,AccessibleName="Template color palette"};
        private readonly ComboBox horizontal=new ComboBox {DropDownStyle=ComboBoxStyle.DropDownList,Width=100,AccessibleName="Horizontal text alignment"};
        private readonly ComboBox vertical=new ComboBox {DropDownStyle=ComboBoxStyle.DropDownList,Width=100,AccessibleName="Vertical text alignment"};
        private readonly ComboBox transition=new ComboBox {DropDownStyle=ComboBoxStyle.DropDownList,Width=100};
        private readonly ComboBox insertMode=new ComboBox {DropDownStyle=ComboBoxStyle.DropDownList,Width=180,AccessibleName="Template insertion target"};
        private readonly CheckBox destinationTheme=new CheckBox {Text="Use destination theme",AutoSize=true};
        private readonly PowerPoint.Application app;
        private SlideTemplate selected;
        private List<SlideTemplate> filtered=new List<SlideTemplate>();
        private int page;
        private readonly LibraryDrafts drafts=new LibraryDrafts();
        private readonly TrackBar thumbnailSize=new TrackBar {Minimum=128,Maximum=256,Value=192,SmallChange=32,LargeChange=32,TickFrequency=32,Width=180,AccessibleName="Thumbnail size"};
        private readonly ComboBox sort=new ComboBox {DropDownStyle=ComboBoxStyle.DropDownList,Width=150,AccessibleName="Sort templates"};
        private readonly Timer previewTimer=new Timer {Interval=250};
        private bool updatingCategories,loadingFields;
        internal TemplateLibrary(PowerPoint.Application application)
        {
            app=application;Text="NexWave Library · 520 original layouts";Width=1480;Height=880;MinimumSize=new Size(1080,700);StartPosition=FormStartPosition.CenterScreen;Font=new Font("Segoe UI",10);
            var split=new SplitContainer {Dock=DockStyle.Fill,Width=1440,SplitterDistance=240};var right=new SplitContainer {Dock=DockStyle.Fill,Width=1200,Orientation=Orientation.Vertical,SplitterDistance=680};
            var leftTitle=new Label {Text="LIBRARY\n520 layouts · 26 families",Dock=DockStyle.Top,Height=62,Padding=new Padding(8)};
            categories.Items.Add("All templates");foreach(var category in TemplateCatalog.Categories)categories.Items.Add(category);split.Panel1.Controls.Add(categories);split.Panel1.Controls.Add(leftTitle);
            results.LargeImageList=thumbnails;right.Panel1.Controls.Add(results);right.Panel1.Controls.Add(search);
            var browseControls=new FlowLayoutPanel {Dock=DockStyle.Top,Height=54};var localFiles=new Button {Text="My PowerPoint files…",AutoSize=true};localFiles.Click+=(s,e)=>{using(var local=new LocalLibraryForm(app))local.ShowDialog(this);};browseControls.Controls.Add(localFiles);browseControls.Controls.Add(new Label {Text="Thumbnail size",AutoSize=true,Padding=new Padding(4,12,0,0)});browseControls.Controls.Add(thumbnailSize);sort.Items.AddRange(new object[]{"Catalog order","Name A–Z"});sort.SelectedIndex=0;browseControls.Controls.Add(sort);right.Panel1.Controls.Add(browseControls);results.ShowItemToolTips=true;
            var pager=new FlowLayoutPanel {Dock=DockStyle.Bottom,AutoSize=true};var previous=new Button {Text="Previous",AutoSize=true};var next=new Button {Text="Next",AutoSize=true};pager.Controls.Add(previous);pager.Controls.Add(pageLabel);pager.Controls.Add(next);right.Panel1.Controls.Add(pager);
            fields.Columns.Add("field","Editable field");fields.Columns[0].ReadOnly=true;fields.Columns[0].FillWeight=40;fields.Columns.Add("value","Your text");fields.Columns[1].FillWeight=60;((DataGridViewTextBoxColumn)fields.Columns[1]).MaxInputLength=2000;
            var controls=new FlowLayoutPanel {Dock=DockStyle.Bottom,Height=180,Padding=new Padding(6)};transition.Items.AddRange(new object[]{"None","Fade","Wipe"});transition.SelectedIndex=0;var insert=new Button {Text="Insert into current presentation",AutoSize=true};var refresh=new Button {Text="Update preview",AutoSize=true};buildStyle.Items.AddRange(new object[]{"Fade","Wipe","Fly in"});buildStyle.SelectedIndex=0;PalettePicker.Configure(palette);palette.SelectedIndex=0;controls.Controls.Add(new Label {Text="Color palette",AutoSize=true});controls.Controls.Add(palette);horizontal.Items.AddRange(TemplateAlignment.Horizontal);horizontal.SelectedItem="Center";vertical.Items.AddRange(TemplateAlignment.Vertical);vertical.SelectedItem="Middle";controls.Controls.Add(new Label {Text="Text alignment",AutoSize=true});controls.Controls.Add(horizontal);controls.Controls.Add(vertical);controls.Controls.Add(builds);controls.Controls.Add(buildStyle);controls.Controls.Add(new Label {Text="Transition",AutoSize=true,Padding=new Padding(5)});controls.Controls.Add(transition);insertMode.Items.AddRange(new object[]{"New slide after current","Objects on current slide"});insertMode.SelectedIndex=0;controls.Controls.Add(insertMode);controls.Controls.Add(destinationTheme);controls.Controls.Add(refresh);controls.Controls.Add(insert);
            right.Panel2.Controls.Add(fields);right.Panel2.Controls.Add(preview);right.Panel2.Controls.Add(status);right.Panel2.Controls.Add(controls);split.Panel2.Controls.Add(right);Controls.Add(split);
            categories.SelectedIndexChanged+=(s,e)=>{if(!updatingCategories)Filter();};search.TextChanged+=(s,e)=>Filter();sort.SelectedIndexChanged+=(s,e)=>Filter();
            thumbnailSize.ValueChanged+=(s,e)=>ShowPage();results.MouseWheel+=(s,e)=>{if((ModifierKeys&Keys.Control)==Keys.Control)thumbnailSize.Value=Math.Max(128,Math.Min(256,thumbnailSize.Value+(e.Delta>0?32:-32)));};
            previous.Click+=(s,e)=>{if(page>0){page--;ShowPage();}};next.Click+=(s,e)=>{if((page+1)*40<filtered.Count){page++;ShowPage();}};
            horizontal.SelectedIndexChanged+=(s,e)=>{if(selected!=null)UpdatePreview();};vertical.SelectedIndexChanged+=(s,e)=>{if(selected!=null)UpdatePreview();};
            results.SelectedIndexChanged+=(s,e)=>SelectTemplate();refresh.Click+=(s,e)=>{if(selected!=null)UpdatePreview();};insert.Click+=(s,e)=>Insert();palette.SelectedIndexChanged+=(s,e)=>ShowPage();
            fields.CellValueChanged+=(s,e)=>{if(!loadingFields&&selected!=null){previewTimer.Stop();previewTimer.Start();}};
            previewTimer.Tick+=(s,e)=>{previewTimer.Stop();if(selected!=null)UpdatePreview();};categories.SelectedIndex=0;
        }
        private void SaveDraft(){if(selected!=null)drafts.Save(selected,Values());}
        private void Filter()
        {
            SaveDraft();int categoryIndex=Math.Max(0,categories.SelectedIndex);string category=categoryIndex==0?null:TemplateCatalog.Categories[categoryIndex-1];
            filtered=TemplateCatalog.Search(category,search.Text).ToList();if(sort.SelectedIndex==1)filtered=filtered.OrderBy(t=>t.Name,StringComparer.CurrentCultureIgnoreCase).ToList();
            updatingCategories=true;
            try {categories.BeginUpdate();categories.Items.Clear();categories.Items.Add("All templates ("+TemplateCatalog.Search(null,search.Text).Count()+")");foreach(var name in TemplateCatalog.Categories)categories.Items.Add(name+" ("+TemplateCatalog.Search(name,search.Text).Count()+")");categories.SelectedIndex=categoryIndex;}
            finally {categories.EndUpdate();updatingCategories=false;}
            page=0;ShowPage();
        }
        private void ShowPage()
        {
            SaveDraft();string retain=selected==null?null:selected.Id;loadingFields=true;previewTimer.Stop();
            results.BeginUpdate();results.Items.Clear();thumbnails.Images.Clear();int thumbWidth=thumbnailSize.Value;thumbnails.ImageSize=new Size(thumbWidth,thumbWidth*9/16);selected=null;fields.Rows.Clear();ReplacePreview(null);
            foreach(var template in filtered.Skip(page*40).Take(40)){using(var image=TemplatePreview.Render(template,thumbWidth,thumbWidth*9/16,drafts.Load(template),((TemplatePalette)palette.SelectedItem).Id,highlight:search.Text))thumbnails.Images.Add(template.Id,image);results.Items.Add(new ListViewItem(template.Name,template.Id){Tag=template,ToolTipText=template.Category+" · "+template.Purpose+" · "+template.Id});}
            pageLabel.Text=filtered.Count==0?"No matching templates":(page+1)+" / "+Math.Ceiling(filtered.Count/40.0)+" · "+filtered.Count+" matches";
            results.EndUpdate();loadingFields=false;
            if(results.Items.Count>0){var item=results.Items.Cast<ListViewItem>().FirstOrDefault(i=>((SlideTemplate)i.Tag).Id==retain)??results.Items[0];item.Selected=true;ShowTemplate((SlideTemplate)item.Tag);}
        }
        private void SelectTemplate()
        {if(loadingFields||results.SelectedItems.Count==0)return;SaveDraft();ShowTemplate((SlideTemplate)results.SelectedItems[0].Tag);}
        private void ShowTemplate(SlideTemplate template)
        {
            loadingFields=true;selected=template;fields.Rows.Clear();foreach(var pair in drafts.Load(template))fields.Rows.Add(pair.Key,pair.Value);loadingFields=false;
            status.Text=template.Category+" · "+template.Purpose+". Preview shows the selected palette; destination theme is applied at insertion.";UpdatePreview();
        }
        private Dictionary<string,string> Values(){fields.EndEdit();return fields.Rows.Cast<DataGridViewRow>().ToDictionary(r=>(string)r.Cells[0].Value,r=>Convert.ToString(r.Cells[1].Value,CultureInfo.InvariantCulture));}
        private void ReplacePreview(Image image){var previous=preview.Image;preview.Image=image;if(previous!=null)previous.Dispose();}
        private void UpdatePreview(){SaveDraft();ReplacePreview(TemplatePreview.Render(selected,768,432,Values(),((TemplatePalette)palette.SelectedItem).Id,(string)horizontal.SelectedItem,(string)vertical.SelectedItem,search.Text));}
        private void Insert()
        {
            if(selected==null)return;
            try {TemplateInserter.Insert(app,selected,Values(),builds.Checked,(string)transition.SelectedItem,(string)buildStyle.SelectedItem,((TemplatePalette)palette.SelectedItem).Id,(string)horizontal.SelectedItem,(string)vertical.SelectedItem,insertMode.SelectedIndex==1,destinationTheme.Checked);status.Text="Inserted "+selected.Name+". Edit text/shapes, Animation Pane and Transitions in PowerPoint.";}
            catch(Exception ex) when(ex is InvalidOperationException||ex is InvalidDataException||ex is System.Runtime.InteropServices.COMException){status.Text="Insertion failed. "+(ex is System.Runtime.InteropServices.COMException?"PowerPoint rejected the operation (0x"+ex.HResult.ToString("X8")+").":ex.Message);}
        }
        protected override void Dispose(bool disposing){if(disposing){previewTimer.Dispose();thumbnailSize.Dispose();sort.Dispose();ReplacePreview(null);thumbnails.Dispose();}base.Dispose(disposing);}
    }
}


namespace PowerPointIntelligence.AddIn
{
    internal sealed class LocalSlideAsset
    {
        internal string File,Title,Text,Author,Layout,Master,Signature;
        internal PowerPointIntelligence.Library.LibrarySlideReference IndexedReference;
        internal int Slide;
        internal float Width,Height;
        internal DateTime Modified;
        public override string ToString(){return Path.GetFileName(File)+" · "+Slide+" · "+Title;}
    }
    internal static class LocalSlideIndex
    {
        private static XmlDocument Read(ZipArchive zip,string part)
        {
            var entry=zip.GetEntry(part);if(entry==null)return null;
            if(entry.Length>2*1024*1024)throw new InvalidDataException("An XML part exceeds the 2 MB library limit.");
            var doc=new XmlDocument {XmlResolver=null};using(var stream=entry.Open())using(var reader=XmlReader.Create(stream,new XmlReaderSettings {DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null,MaxCharactersInDocument=2*1024*1024}))doc.Load(reader);return doc;
        }
        private static string SlideText(ZipArchive zip,string part,out string digest)
        {
            var entry=zip.GetEntry(part);if(entry==null)throw new InvalidDataException("Missing slide part.");if(entry.Length>40L*1024*1024)throw new InvalidDataException("Slide XML exceeds the 40 MB streaming limit.");var text=new StringBuilder();
            using(var stream=entry.Open())using(var reader=XmlReader.Create(stream,new XmlReaderSettings {DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null,MaxCharactersInDocument=40L*1024*1024})){while(reader.Read())if(reader.NodeType==XmlNodeType.Element&&reader.LocalName=="t"){string value=reader.ReadString();if(text.Length+value.Length+1>64000)throw new InvalidDataException("Slide text exceeds 64000 characters.");if(text.Length>0)text.Append(' ');text.Append(value);}}
            using(var stream=entry.Open())using(var hash=SHA256.Create())digest=BitConverter.ToString(hash.ComputeHash(stream));return text.ToString();
        }
        private static string Value(XmlDocument doc,string local){var node=doc==null?null:doc.SelectSingleNode("//*[local-name()='"+local+"']");return node==null?"":node.InnerText;}
        private static string Resolve(string source,string target)
        {
            if(String.IsNullOrWhiteSpace(target)||target.Contains(":")||target.Contains("\\"))throw new InvalidDataException("Invalid package relationship.");
            var uri=new Uri(new Uri("https://package/"+source),target);if(uri.Host!="package")throw new InvalidDataException("External package part.");return Uri.UnescapeDataString(uri.AbsolutePath.TrimStart('/'));
        }
        private static string Relation(XmlDocument rels,string id,string suffix=null)
        {
            if(rels==null)return null;foreach(XmlElement node in rels.DocumentElement.ChildNodes.OfType<XmlElement>())if((id==null||node.GetAttribute("Id")==id)&&(suffix==null||node.GetAttribute("Type").EndsWith(suffix,StringComparison.Ordinal))){if(node.GetAttribute("TargetMode")=="External")throw new InvalidDataException("External slide relationships are not indexed.");return node.GetAttribute("Target");}return null;
        }
        internal static bool IsWithin(string root,string file){return Path.GetFullPath(file).StartsWith(Path.GetFullPath(root).TrimEnd(Path.DirectorySeparatorChar)+Path.DirectorySeparatorChar,StringComparison.OrdinalIgnoreCase);}
        internal static string[] Discover(string root,bool recursive)
        {
            var found=new List<string>();var pending=new Stack<string>();pending.Push(Path.GetFullPath(root));int directories=0,files=0;
            while(pending.Count>0){string directory=pending.Pop();if(++directories>500)throw new InvalidDataException("Library exceeds 500 folders.");foreach(string file in Directory.EnumerateFiles(directory)){if(++files>5000)throw new InvalidDataException("Library exceeds 5000 files.");if(new[]{".pptx",".potx"}.Contains(Path.GetExtension(file).ToLowerInvariant())){if(found.Count>=200)throw new InvalidDataException("Library exceeds 200 presentation files.");found.Add(file);}}if(recursive)foreach(string child in Directory.EnumerateDirectories(directory))if((File.GetAttributes(child)&FileAttributes.ReparsePoint)==0)pending.Push(child);}
            return found.OrderBy(f=>f,StringComparer.OrdinalIgnoreCase).ToArray();
        }
        internal static List<LocalSlideAsset> ReadFile(string file)
        {
            var info=new FileInfo(file);if(!new[]{".pptx",".potx"}.Contains(info.Extension.ToLowerInvariant()))throw new InvalidDataException("Choose a .pptx or .potx file.");if(info.Length>100*1024*1024)throw new InvalidDataException("Library files are limited to 100 MB.");
            string packageHash;using(var stream=info.OpenRead())using(var hash=SHA256.Create())packageHash=BitConverter.ToString(hash.ComputeHash(stream));
            using(var zip=ZipFile.OpenRead(file)){
                if(zip.Entries.Count>10000||zip.Entries.Sum(e=>e.Length)>200L*1024*1024)throw new InvalidDataException("Package exceeds library limits.");
                var presentation=Read(zip,"ppt/presentation.xml");var rels=Read(zip,"ppt/_rels/presentation.xml.rels");if(presentation==null||rels==null)throw new InvalidDataException("Missing presentation parts.");
                var size=(XmlElement)presentation.SelectSingleNode("//*[local-name()='sldSz']");float width=0,height=0;if(size==null||!Single.TryParse(size.GetAttribute("cx"),NumberStyles.None,CultureInfo.InvariantCulture,out width)||!Single.TryParse(size.GetAttribute("cy"),NumberStyles.None,CultureInfo.InvariantCulture,out height)||width<=0||height<=0)throw new InvalidDataException("Invalid slide size.");
                string author=Value(Read(zip,"docProps/core.xml"),"creator");int textBudget=0;var result=new List<LocalSlideAsset>();var slides=presentation.SelectNodes("//*[local-name()='sldIdLst']/*[local-name()='sldId']");if(slides.Count>1000)throw new InvalidDataException("Library files are limited to 1000 slides.");
                foreach(XmlElement item in slides){string id=item.GetAttribute("id","http://schemas.openxmlformats.org/officeDocument/2006/relationships");string target=Relation(rels,id);if(target==null)throw new InvalidDataException("Missing slide relationship.");string part=Resolve("ppt/presentation.xml",target);string slideDigest;string text=SlideText(zip,part,out slideDigest);if(text.Length>64000||(textBudget+=text.Length)>2000000)throw new InvalidDataException("Slide text exceeds library limits.");string folder=part.Substring(0,part.LastIndexOf('/')+1);var slideRels=Read(zip,folder+"_rels/"+Path.GetFileName(part)+".rels");string layoutTarget=Relation(slideRels,null,"/slideLayout"),layoutName="",master="";
                    if(layoutTarget!=null){string layoutPart=Resolve(part,layoutTarget);var layout=Read(zip,layoutPart);var common=layout==null?null:layout.SelectSingleNode("//*[local-name()='cSld']") as XmlElement;layoutName=common==null?Path.GetFileNameWithoutExtension(layoutPart):common.GetAttribute("name");string dir=layoutPart.Substring(0,layoutPart.LastIndexOf('/')+1);var layoutRels=Read(zip,dir+"_rels/"+Path.GetFileName(layoutPart)+".rels");string masterTarget=Relation(layoutRels,null,"/slideMaster");if(masterTarget!=null)master=Resolve(layoutPart,masterTarget);}
                    string signature;using(var hash=SHA256.Create())signature=BitConverter.ToString(hash.ComputeHash(Encoding.UTF8.GetBytes(packageHash+slideDigest+(slideRels==null?"":slideRels.OuterXml))));
                    result.Add(new LocalSlideAsset {File=info.FullName,Slide=result.Count+1,Text=text,Title=text.Length>100?text.Substring(0,100):text,Author=author,Layout=layoutName,Master=master,Width=width/12700,Height=height/12700,Modified=info.LastWriteTimeUtc,Signature=signature});
                }return result;
            }
        }
        internal static IEnumerable<LocalSlideAsset> Filter(IEnumerable<LocalSlideAsset> items,string query,string author,string layout,string aspect,DateTime? since,bool duplicates)
        {
            var terms=Regex.Split((query??"").Trim(),@"\s+").Where(t=>t.Length>0);var result=items.Where(a=>terms.All(t=>(Path.GetFileName(a.File)+" "+a.Text).IndexOf(t,StringComparison.OrdinalIgnoreCase)>=0)&&(String.IsNullOrEmpty(author)||a.Author.IndexOf(author,StringComparison.OrdinalIgnoreCase)>=0)&&(String.IsNullOrEmpty(layout)||(a.Layout+" "+a.Master).IndexOf(layout,StringComparison.OrdinalIgnoreCase)>=0)&&(!since.HasValue||a.Modified>=since.Value.ToUniversalTime())&&(aspect=="All"||aspect=="16:9"&&Math.Abs(a.Width/a.Height-16f/9)<.02||aspect=="4:3"&&Math.Abs(a.Width/a.Height-4f/3)<.02));
            return duplicates?result.GroupBy(a=>a.Signature).Select(g=>g.First()):result;
        }
    }

}
namespace PowerPointIntelligence.AddIn
{
    internal sealed class LayoutRule
    {
        internal string Leader,LeaderEdge,OwnEdge;
        internal float Offset;
        internal string Encode(){Validate();return "v1|"+Leader+"|"+LeaderEdge+"|"+OwnEdge+"|"+Offset.ToString("R",CultureInfo.InvariantCulture);}
        internal static LayoutRule Decode(string value){var parts=(value??"").Split('|');float offset;if(parts.Length!=5||!String.Equals(parts[0],"v1",StringComparison.OrdinalIgnoreCase)||!Single.TryParse(parts[4],NumberStyles.Float,CultureInfo.InvariantCulture,out offset))throw new InvalidDataException("Invalid layout relationship.");var rule=new LayoutRule {Leader=parts[1],LeaderEdge=parts[2].ToUpperInvariant(),OwnEdge=parts[3].ToUpperInvariant(),Offset=offset};rule.Validate();return rule;}
        private void Validate(){Guid id;if(Leader!="SLIDE"&&!Guid.TryParse(Leader,out id))throw new InvalidDataException("Invalid layout anchor.");if(!new[]{"L","R","C","T","B","M"}.Contains(LeaderEdge)||!new[]{"L","R","C","T","B","M"}.Contains(OwnEdge)||Horizontal(LeaderEdge)!=Horizontal(OwnEdge)||Single.IsNaN(Offset)||Single.IsInfinity(Offset)||Math.Abs(Offset)>100000)throw new InvalidDataException("Invalid layout relationship geometry.");}
        internal static bool Horizontal(string edge){return edge=="L"||edge=="R"||edge=="C";}
        internal static float Edge(RectangleF rect,string edge){return edge=="L"?rect.Left:edge=="R"?rect.Right:edge=="C"?rect.Left+rect.Width/2:edge=="T"?rect.Top:edge=="B"?rect.Bottom:rect.Top+rect.Height/2;}
        internal float Position(RectangleF leader,SizeF own){Validate();return Edge(leader,LeaderEdge)+Offset-(OwnEdge=="R"?own.Width:OwnEdge=="C"?own.Width/2:OwnEdge=="B"?own.Height:OwnEdge=="M"?own.Height/2:0);}
    }
    internal static class LibraryLayoutLinks
    {
        private const string Identity="PPI_LAYOUT_ANCHOR",X="PPI_LAYOUT_X",Y="PPI_LAYOUT_Y";
        private static RectangleF Rect(PowerPoint.Shape s){return new RectangleF(s.Left,s.Top,s.Width,s.Height);}
        internal static void Set(PowerPoint.Application app,string mode)
        {
            var selection=app.ActiveWindow.Selection;if(selection.Type!=PowerPoint.PpSelectionType.ppSelectionShapes)throw new InvalidOperationException("Select native objects first.");var shapes=selection.ShapeRange.Cast<PowerPoint.Shape>().ToArray();if(shapes.Length==0)throw new InvalidOperationException("Select native objects first.");var slide=app.ActiveWindow.View.Slide as PowerPoint.Slide;if(slide==null)throw new InvalidOperationException("Use Normal view.");
            if(mode=="Clear links"){app.StartNewUndoEntry();var identities=new HashSet<string>(shapes.Select(s=>s.Tags[Identity]).Where(id=>!String.IsNullOrEmpty(id)),StringComparer.OrdinalIgnoreCase);foreach(PowerPoint.Shape other in slide.Shapes)foreach(string tag in new[]{X,Y}){string value=other.Tags[tag];if(!String.IsNullOrEmpty(value)&&identities.Contains(LayoutRule.Decode(value).Leader))other.Tags.Delete(tag);}foreach(var shape in shapes){shape.Tags.Delete(X);shape.Tags.Delete(Y);shape.Tags.Delete(Identity);}return;}
            if(mode=="Fit shape to text"){app.StartNewUndoEntry();foreach(var shape in shapes)if(shape.HasTextFrame==Office.MsoTriState.msoTrue)shape.TextFrame2.AutoSize=Office.MsoAutoSize.msoAutoSizeShapeToFitText;return;}
            bool slideLock=mode.StartsWith("Lock ",StringComparison.Ordinal);string edge=mode.EndsWith("left")?"L":mode.EndsWith("right")?"R":mode.EndsWith("top")?"T":mode.EndsWith("bottom")?"B":mode.EndsWith("middle")?"M":"C";bool horizontal=LayoutRule.Horizontal(edge);string leaderId="SLIDE";var leaderRect=new RectangleF(0,0,app.ActivePresentation.PageSetup.SlideWidth,app.ActivePresentation.PageSetup.SlideHeight);var followers=shapes;
            if(!slideLock){if(shapes.Length<2)throw new InvalidOperationException("Select an anchor first, then follower objects. The first shape in PowerPoint's selection order is the anchor.");var leader=shapes[0];if(!String.IsNullOrEmpty(leader.Tags[X])||!String.IsNullOrEmpty(leader.Tags[Y]))throw new InvalidOperationException("Clear the anchor's own links first. Chained or cyclic relationships are not supported.");leaderId=leader.Tags[Identity];Guid parsed;if(!Guid.TryParse(leaderId,out parsed))leaderId=Guid.NewGuid().ToString("D");int count=0;foreach(PowerPoint.Shape s in slide.Shapes)if(String.Equals(s.Tags[Identity],leaderId,StringComparison.OrdinalIgnoreCase))count++;if(count>1)throw new InvalidOperationException("Copied anchor identity is ambiguous. Clear anchor identity before linking again.");leaderRect=Rect(leader);followers=shapes.Skip(1).ToArray();}
            bool gap=mode=="Keep horizontal gap"||mode=="Keep vertical gap";if(gap){horizontal=mode=="Keep horizontal gap";edge=horizontal?"L":"T";}
            var plans=followers.Select(s=>new {Shape=s,Rule=new LayoutRule {Leader=leaderId,LeaderEdge=gap?(horizontal?"R":"B"):edge,OwnEdge=edge,Offset=slideLock||gap?LayoutRule.Edge(Rect(s),edge)-LayoutRule.Edge(leaderRect,gap?(horizontal?"R":"B"):edge):0}}).ToArray();foreach(var plan in plans)plan.Rule.Encode();
            // An anchor already referenced by other objects cannot become a follower.
            foreach(var plan in plans){string id=plan.Shape.Tags[Identity];if(String.IsNullOrEmpty(id))continue;foreach(PowerPoint.Shape other in slide.Shapes)foreach(string tag in new[]{X,Y}){string value=other.Tags[tag];if(!String.IsNullOrEmpty(value)&&String.Equals(LayoutRule.Decode(value).Leader,id,StringComparison.OrdinalIgnoreCase))throw new InvalidOperationException("A referenced anchor cannot become a follower. Clear its dependents first.");}}
            app.StartNewUndoEntry();var undo=new List<Action>();try{if(!slideLock){string old=shapes[0].Tags[Identity];undo.Add(()=>{if(String.IsNullOrEmpty(old))shapes[0].Tags.Delete(Identity);else shapes[0].Tags.Add(Identity,old);});shapes[0].Tags.Add(Identity,leaderId);}foreach(var plan in plans){string tag=horizontal?X:Y,old=plan.Shape.Tags[tag];var target=plan.Shape;undo.Add(()=>{if(String.IsNullOrEmpty(old))target.Tags.Delete(tag);else target.Tags.Add(tag,old);});target.Tags.Add(tag,plan.Rule.Encode());}slide.Tags.Add("PPI_LAYOUT_ENABLED","1");Reflow(app);}catch{for(int i=undo.Count-1;i>=0;i--)try{undo[i]();}catch{}throw;}
        }
        internal static void Reflow(PowerPoint.Application app)
        {
            if(app.Presentations.Count==0||app.SlideShowWindows.Count>0||app.ActiveWindow.ViewType!=PowerPoint.PpViewType.ppViewNormal)return;var slide=app.ActiveWindow.View.Slide as PowerPoint.Slide;if(slide==null||slide.Tags["PPI_LAYOUT_ENABLED"]!="1")return;var shapes=slide.Shapes.Cast<PowerPoint.Shape>().ToArray();if(shapes.Length>2000)throw new InvalidOperationException("Layout links support at most 2000 top-level objects.");var linked=shapes.Where(s=>!String.IsNullOrEmpty(s.Tags[X])||!String.IsNullOrEmpty(s.Tags[Y])).ToArray();if(linked.Length==0)return;
            var canvas=new RectangleF(0,0,app.ActivePresentation.PageSetup.SlideWidth,app.ActivePresentation.PageSetup.SlideHeight);var anchors=shapes.Where(s=>!String.IsNullOrEmpty(s.Tags[Identity])).GroupBy(s=>s.Tags[Identity],StringComparer.OrdinalIgnoreCase).ToDictionary(g=>g.Key,g=>g.ToArray(),StringComparer.OrdinalIgnoreCase);var planned=new List<Tuple<PowerPoint.Shape,PointF,PointF>>();
            foreach(var shape in linked){var old=Rect(shape);float left=old.Left,top=old.Top;foreach(string tag in new[]{X,Y}){string value=shape.Tags[tag];if(String.IsNullOrEmpty(value))continue;var rule=LayoutRule.Decode(value);RectangleF leader=canvas;if(rule.Leader!="SLIDE"){PowerPoint.Shape[] found;if(!anchors.TryGetValue(rule.Leader,out found)||found.Length!=1)throw new InvalidOperationException("Missing or copied layout anchor. Clear and recreate the relationship.");if(!String.IsNullOrEmpty(found[0].Tags[X])||!String.IsNullOrEmpty(found[0].Tags[Y]))throw new InvalidOperationException("Chained or cyclic layout link rejected.");leader=Rect(found[0]);}float position=rule.Position(leader,old.Size);if((tag==X)!=LayoutRule.Horizontal(rule.OwnEdge))throw new InvalidDataException("Layout axis mismatch.");if(tag==X)left=position;else top=position;}if(left<0||top<0||left+old.Width>canvas.Width+.1f||top+old.Height>canvas.Height+.1f)throw new InvalidOperationException("Layout link would place an object outside the slide. No linked objects moved.");planned.Add(Tuple.Create(shape,new PointF(old.Left,old.Top),new PointF(left,top)));}
            var moved=new List<Tuple<PowerPoint.Shape,PointF,PointF>>();try{foreach(var plan in planned){if(Math.Abs(plan.Item2.X-plan.Item3.X)<.1f&&Math.Abs(plan.Item2.Y-plan.Item3.Y)<.1f)continue;moved.Add(plan);plan.Item1.Left=plan.Item3.X;plan.Item1.Top=plan.Item3.Y;}}catch(Exception original){bool restored=true;foreach(var plan in moved.AsEnumerable().Reverse())try{plan.Item1.Left=plan.Item2.X;plan.Item1.Top=plan.Item2.Y;}catch{restored=false;}if(!restored)throw new InvalidOperationException("Layout update failed and rollback was incomplete. Use PowerPoint Undo.",original);throw;}
        }
        internal static Form Controls(PowerPoint.Application app)
        {
            var form=new Form {Text="Library layout links",Width=410,Height=480,StartPosition=FormStartPosition.CenterScreen};var panel=new FlowLayoutPanel {Dock=DockStyle.Fill,Padding=new Padding(10),AutoScroll=true};var status=new Label {Text="Opt-in native relationships. Select the anchor first, then followers. Updates occur while the add-in is loaded. Shape-to-text fit may change geometry.",AutoSize=false,Width=360,Height=70};panel.Controls.Add(status);
            foreach(string mode in new[]{"Align left","Align right","Align top","Align bottom","Align center","Align middle","Keep horizontal gap","Keep vertical gap","Lock left","Lock right","Lock top","Lock bottom","Fit shape to text","Clear links"}){string action=mode;var button=new Button {Text=mode,Width=170,Height=32};button.Click+=(s,e)=>{try{Set(app,action);status.Text="Applied: "+action+". Select Clear links to release constraints.";}catch(Exception ex) when(ex is InvalidOperationException||ex is InvalidDataException||ex is System.Runtime.InteropServices.COMException){status.Text=ex.Message;}};panel.Controls.Add(button);}form.Controls.Add(panel);return form;
        }
    }
}
namespace PowerPointIntelligence.AddIn
{
    internal static class TemplateSources
    {
        internal static string ReadFolderSetting(string xml)
        {
            if(String.IsNullOrEmpty(xml)||xml.Length>1000000)throw new InvalidDataException("Template configuration is missing or excessive.");var doc=new XmlDocument {XmlResolver=null};using(var reader=XmlReader.Create(new StringReader(xml),new XmlReaderSettings {DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null,MaxCharactersInDocument=1000000}))doc.Load(reader);if(doc.DocumentElement==null||doc.DocumentElement.Name!="configuration")throw new InvalidDataException("Invalid add-in configuration.");var settings=doc.SelectNodes("/configuration/appSettings/add[@key='TemplateLibraryFolder']");if(settings.Count>1)throw new InvalidDataException("TemplateLibraryFolder is configured more than once.");return settings.Count==0?null:((XmlElement)settings[0]).GetAttribute("value");
        }
        internal static string ResolveFolder(string assemblyDirectory,string configured,string vendorFolder,Func<string,bool> exists)
        {
            if(!String.IsNullOrWhiteSpace(configured)){string explicitFolder=Path.GetFullPath(Path.IsPathRooted(configured)?configured:Path.Combine(assemblyDirectory,configured));if(!exists(explicitFolder))throw new DirectoryNotFoundException("Configured template folder is unavailable. Update TemplateLibraryFolder in the add-in configuration.");return explicitFolder;}
            string local=Path.GetFullPath(Path.Combine(assemblyDirectory,"templates"));if(exists(local))return local;return null;
        }
        internal static string InstalledOriginals
        {
            get {string assembly=typeof(TemplateSources).Assembly.Location;string config=assembly+".config";string setting=File.Exists(config)?ReadFolderSetting(File.ReadAllText(config)):null;return ResolveFolder(Path.GetDirectoryName(assembly),setting,null,Directory.Exists);}
        }
        internal static void RequireSameSize(LocalSlideAsset item,float width,float height)
        {if(Math.Abs(item.Width-width)>.05f||Math.Abs(item.Height-height)>.05f)throw new InvalidOperationException("Keep-original insertion requires the same slide dimensions as the template ("+item.Width.ToString("0.##",CultureInfo.InvariantCulture)+" × "+item.Height.ToString("0.##",CultureInfo.InvariantCulture)+" pt). Change the destination slide size explicitly in PowerPoint first; this command never resizes your presentation.");}
    }
    internal sealed class SourceTemplateSession:IDisposable
    {
        private readonly PowerPoint.Application app;
        internal readonly PowerPoint.Presentation Destination;
        internal readonly PowerPoint.Slide TargetSlide;
        private PowerPoint.Presentation source;
        private string sourceFile;
        private DateTime sourceWriteTime;
        private long sourceLength;
        private string sourceFingerprint;
        internal SourceTemplateSession(PowerPoint.Application app)
        {
            this.app=app;if(app.Presentations.Count==0)throw new InvalidOperationException("Open a destination presentation first.");Destination=app.ActivePresentation;TargetSlide=app.ActiveWindow.View.Slide as PowerPoint.Slide;if(TargetSlide==null)throw new InvalidOperationException("Select a slide in Normal view first.");
        }
        internal PowerPoint.Slide Source(LocalSlideAsset item)
        {
            if(String.Equals(Path.GetFullPath(Destination.FullName),item.File,StringComparison.OrdinalIgnoreCase))throw new InvalidOperationException("The source must differ from the destination presentation.");
            if(item.IndexedReference!=null&&PowerPointIntelligence.Library.PresentationSource.Fingerprint(item.File,System.Threading.CancellationToken.None)!=item.IndexedReference.SourceFingerprint)throw new InvalidOperationException("Indexed source changed; refresh before preview or insertion.");
            var info=new FileInfo(item.File);if(source!=null&&((item.IndexedReference!=null&&sourceFingerprint!=item.IndexedReference.SourceFingerprint)||!String.Equals(sourceFile,item.File,StringComparison.OrdinalIgnoreCase)||sourceWriteTime!=info.LastWriteTimeUtc||sourceLength!=info.Length))Dispose();
            if(source==null){foreach(PowerPoint.Presentation open in app.Presentations)if(String.Equals(Path.GetFullPath(open.FullName),item.File,StringComparison.OrdinalIgnoreCase))throw new InvalidOperationException("This source is already open in PowerPoint. Close it normally before previewing; the library never takes ownership of a user presentation.");source=app.Presentations.Open(item.File,Office.MsoTriState.msoTrue,Office.MsoTriState.msoFalse,Office.MsoTriState.msoFalse);sourceFile=item.File;sourceWriteTime=info.LastWriteTimeUtc;sourceLength=info.Length;sourceFingerprint=item.IndexedReference==null?null:item.IndexedReference.SourceFingerprint;}
            if(item.Slide<1||item.Slide>source.Slides.Count)throw new InvalidDataException("Source slide is missing. Refresh the library.");var selected=source.Slides[item.Slide];if(item.IndexedReference!=null&&selected.SlideID.ToString(CultureInfo.InvariantCulture)!=item.IndexedReference.NativeId)throw new InvalidDataException("Native slide identity changed; refresh before insertion.");return selected;
        }
        internal Bitmap Preview(LocalSlideAsset item)
        {
            var slide=Source(item);string file=Path.Combine(Path.GetTempPath(),"ppi-source-preview-"+Guid.NewGuid().ToString("N")+".png");
            try{slide.Export(file,"PNG",720,(int)Math.Round(720*item.Height/item.Width));using(var image=Image.FromFile(file))return new Bitmap(image);}finally{if(File.Exists(file))File.Delete(file);}
        }
        internal PowerPoint.Slide InsertOriginal(LocalSlideAsset item)
        {
            TemplateSources.RequireSameSize(item,Destination.PageSetup.SlideWidth,Destination.PageSetup.SlideHeight);
            if(item.IndexedReference==null){var current=LocalSlideIndex.ReadFile(item.File).Single(a=>a.Slide==item.Slide);if(current.Signature!=item.Signature)throw new InvalidOperationException("Source changed. Refresh before inserting the original.");}
            // Use the validated preview source instance; Source reloads if its file stamp changes.
            var original=Source(item);PowerPoint.Slide inserted=null;PowerPoint.Design cloned=null;app.StartNewUndoEntry();
            try{int index=TargetSlide.SlideIndex;int count=Destination.Slides.InsertFromFile(item.File,index,item.Slide,item.Slide);if(count!=1)throw new InvalidOperationException("PowerPoint did not insert exactly one source slide.");inserted=Destination.Slides[index+1];cloned=Destination.Designs.Clone(original.Design);inserted.Design=cloned;int layout=original.CustomLayout.Index;if(layout<1||layout>cloned.SlideMaster.CustomLayouts.Count)throw new InvalidDataException("Source custom layout could not be retained.");inserted.CustomLayout=cloned.SlideMaster.CustomLayouts[layout];inserted.FollowMasterBackground=original.FollowMasterBackground;
                inserted.Tags.Add("PPI_LIBRARY_ASSET_SCHEMA","1");inserted.Tags.Add("PPI_LIBRARY_ASSET_VERSION",item.Signature);inserted.Tags.Add("PPI_LIBRARY_SOURCE_SLIDE",item.Slide.ToString(CultureInfo.InvariantCulture));inserted.Tags.Add("PPI_LIBRARY_FORMATTING","KEEP_SOURCE");if(item.IndexedReference!=null){inserted.Tags.Add("PPI_LIBRARY_INDEX_SCHEMA","1");inserted.Tags.Add("PPI_LIBRARY_SLIDE_ID",item.IndexedReference.SlideId);}inserted.Select();return inserted;
            }catch(Exception originalError){bool cleaned=true;if(inserted!=null)try{inserted.Delete();}catch{cleaned=false;}if(cloned!=null)try{cloned.Delete();}catch{cleaned=false;}if(!cleaned)throw new InvalidOperationException("Original insertion failed and cleanup was incomplete. Use PowerPoint Undo.",originalError);throw;}
        }
        public void Dispose(){if(source==null)return;source.Close();source=null;sourceFile=null;}
    }
}
