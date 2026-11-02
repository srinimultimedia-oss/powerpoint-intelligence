using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using PowerPointIntelligence.AddIn;
using Office=Microsoft.Office.Core;
using PowerPoint=Microsoft.Office.Interop.PowerPoint;
internal static class TextShapeOfficeTests
{
    private static int count;
    private static void Check(bool ok,string name){if(!ok)throw new Exception("FAIL text element Office: "+name);count++;Console.WriteLine("PASS text element Office: "+name);}
    private static RectangleF Bounds(PowerPoint.Shape s){return new RectangleF(s.Left,s.Top,s.Width,s.Height);}
    internal static void Run(PowerPoint.Application app,string outputDirectory)
    {
        if(app==null)throw new ArgumentNullException("app");count=0;Directory.CreateDirectory(outputDirectory);PowerPoint.Presentation owned=null;string file=Path.Combine(Path.GetFullPath(outputDirectory),"text-shape-native-"+Guid.NewGuid().ToString("N")+".pptx");
        Console.WriteLine("TEXT ELEMENT Office="+app.Version+"; process="+(Environment.Is64BitProcess?"x64":"x86"));
        try
        {
            owned=app.Presentations.Add(Office.MsoTriState.msoTrue);owned.PageSetup.SlideWidth=960;owned.PageSetup.SlideHeight=540;owned.Windows[1].ViewType=PowerPoint.PpViewType.ppViewNormal;var slide=owned.Slides.Add(1,PowerPoint.PpSlideLayout.ppLayoutBlank);slide.Select();
            var box=TextShapeElements.Insert(app,"textbox");int boxId=box.Id;
            Check(TextShapeElements.IsManaged(box)&&box.HasTextFrame==Office.MsoTriState.msoTrue&&box.TextFrame.TextRange.Text=="","blank native editable element inserted without sample content");
            Check(box.TextFrame.TextRange.ParagraphFormat.Alignment==PowerPoint.PpParagraphAlignment.ppAlignCenter&&box.TextFrame.VerticalAnchor==Office.MsoVerticalAnchor.msoAnchorMiddle,"center alignment defaults");
            Check(box.TextFrame2.AutoSize==Office.MsoAutoSize.msoAutoSizeNone,"font shrinking disabled");
            box.TextFrame.TextRange.Text="First line\rSecond line\rThird line";box.Left=100;box.Top=100;box.Width=240;box.Height=30;float font=box.TextFrame.TextRange.Font.Size;box.Select();TextShapeElements.FitSelected(app);
            Check(box.Height>30&&Math.Abs(box.Left-100)<.1&&Math.Abs(box.Top-100)<.1&&Math.Abs(box.Width-240)<.1,"actual text bounds grow unlocked bottom with other edges fixed");
            Check(Math.Abs(box.TextFrame.TextRange.Font.Size-font)<.1,"fitting preserves native font size");
            float tall=box.Height;box.TextFrame.TextRange.Text="Short";box.Select();TextShapeElements.FitSelected(app);Check(box.Height<tall&&Math.Abs(box.Top-100)<.1,"unlocked edge shrinks when text shortens");
            var spec=TextShapeElements.Read(box);spec.Locks=TextEdgeLocks.Right|TextEdgeLocks.Bottom;box.Select();TextShapeElements.Configure(app,spec);float right=box.Left+box.Width,bottom=box.Top+box.Height;box.TextFrame.TextRange.Text="A wider native label";box.Select();TextShapeElements.FitSelected(app);Check(Math.Abs(box.Left+box.Width-right)<.2&&Math.Abs(box.Top+box.Height-bottom)<.2,"right and bottom locks preserve opposite edges");
            box.TextFrame.TextRange.Characters(1,1).Font.Bold=Office.MsoTriState.msoTrue;box.Select();var converted=TextShapeElements.ConvertSelected(app,"roundedRectangle");Check(converted.Id==boxId&&converted.AutoShapeType==Office.MsoAutoShapeType.msoShapeRoundedRectangle&&converted.TextFrame.TextRange.Text=="A wider native label"&&converted.TextFrame.TextRange.Characters(1,1).Font.Bold==Office.MsoTriState.msoTrue,"conversion preserves identity native text and mixed formatting");
            converted.Select();NativeFormatting.ApplyColor(app,"fill",0xB0D0E0);NativeFormatting.ApplyColor(app,"outline",0x304050);NativeFormatting.Text(app,"right");Check(converted.Fill.ForeColor.RGB==0xB0D0E0&&converted.Line.ForeColor.RGB==0x304050&&converted.TextFrame.TextRange.ParagraphFormat.Alignment==PowerPoint.PpParagraphAlignment.ppAlignRight,"existing native formatting controls apply (fill="+converted.Fill.ForeColor.RGB+", outline="+converted.Line.ForeColor.RGB+", alignment="+converted.TextFrame.TextRange.ParagraphFormat.Alignment+")");
            converted.Left=100;converted.Top=100;converted.Width=200;converted.Height=70;converted.TextFrame.TextRange.Text="Phase";converted.Select();converted=TextShapeElements.ConvertSelected(app,"rightPentagon");converted.Left=100;converted.Top=100;converted.Width=200;converted.Height=70;converted.Select();var copy=TextShapeElements.DuplicateSelected(app,"right");int copyId=copy.Id;
            Check(copy.AutoShapeType==Office.MsoAutoShapeType.msoShapeChevron&&TextShapeElements.Read(copy).Kind=="chevron"&&copy.TextFrame.TextRange.Text=="Phase"&&Math.Abs(copy.Left-312)<.1,"directional duplicate creates native chevron at bounded position");
            copy.Select();spec=TextShapeElements.Read(copy);spec.Adjustment=.4f;TextShapeElements.Configure(app,spec);Check(Math.Abs(TextShapeElements.Read(converted).Adjustment-.25)<.01&&Math.Abs(TextShapeElements.Read(copy).Adjustment-.4)<.01,"duplicate has independent persisted adjustment settings");
            // A late failing target must not resize an earlier target in a multi-selection.
            converted.Left=100;converted.Top=100;converted.Width=200;converted.Height=20;converted.TextFrame.TextRange.Text="One\rTwo\rThree";copy.Left=0;copy.Top=100;copy.Width=100;copy.Height=30;copy.TextFrame.TextRange.Text=new string('W',300);var first=Bounds(converted);var second=Bounds(copy);slide.Select();var indices=Enumerable.Range(1,slide.Shapes.Count).Where(i=>slide.Shapes[i].Id==converted.Id||slide.Shapes[i].Id==copy.Id).Select(i=>(object)i).ToArray();if(indices.Length!=2)throw new InvalidOperationException("Both native test targets must exist before multi-selection.");slide.Shapes.Range(indices).Select();bool refused=false;try{TextShapeElements.FitSelected(app);}catch(Exception error) when(error is InvalidOperationException||error is InvalidDataException&&error.Message=="Invalid fitting geometry."){refused=true;}
            Check(refused&&Bounds(converted)==first&&Bounds(copy)==second,"multi-selection overflow preflight preserves every original");
            copy.TextFrame.TextRange.Text="Editable chevron";copy.Left=500;copy.Top=200;copy.Width=250;copy.Height=70;converted.TextFrame.TextRange.Text="Editable pentagon";converted.Left=100;converted.Top=100;converted.Width=250;converted.Height=70;
            owned.SaveAs(file,PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation,Office.MsoTriState.msoFalse);owned.Close();owned=null;owned=app.Presentations.Open(file,Office.MsoTriState.msoFalse,Office.MsoTriState.msoFalse,Office.MsoTriState.msoTrue);slide=owned.Slides[1];converted=slide.Shapes.Cast<PowerPoint.Shape>().Single(s=>s.Id==boxId);copy=slide.Shapes.Cast<PowerPoint.Shape>().Single(s=>s.Id==copyId);
            Check(TextShapeElements.Read(converted).Kind=="rightPentagon"&&TextShapeElements.Read(copy).Kind=="chevron"&&Math.Abs(TextShapeElements.Read(copy).Adjustment-.4)<.01,"typed versioned settings survive save and reopen");
            copy.TextFrame.TextRange.Text="Still native editable";copy.Select();TextShapeElements.FitSelected(app);Check(copy.TextFrame.TextRange.Text=="Still native editable"&&copy.Type==Office.MsoShapeType.msoAutoShape,"editing and fitting work after reopen");
            owned.Save();Console.WriteLine("TEXT ELEMENT "+count+" native assertions passed. Evidence deck: "+file);
        }
        finally{if(owned!=null){owned.Saved=Office.MsoTriState.msoTrue;owned.Close();}}
    }
}
