using System;
using System.Globalization;
using System.IO;
using System.Linq;
using PowerPointIntelligence.AddIn;
using Office=Microsoft.Office.Core;
using PowerPoint=Microsoft.Office.Interop.PowerPoint;
internal static class ProcessFlowOfficeTests
{
    private static int count;
    private static void Check(bool value,string name){if(!value)throw new Exception("FAIL process flow Office: "+name);count++;Console.WriteLine("PASS process flow Office: "+name);}
    private static ProcessGridSpec State(PowerPoint.Presentation deck,string id){return deck.CustomXMLParts.SelectByNamespace(ProcessGridCodec.Namespace).Cast<Office.CustomXMLPart>().Select(p=>ProcessGridCodec.Decode(p.XML)).Single(g=>String.Equals(g.Id,id,StringComparison.OrdinalIgnoreCase));}
    private static PowerPoint.Shape Cell(PowerPoint.Slide slide,ProcessGridSpec grid,int row,int column){int id=grid.Cells.Single(c=>c.Active&&c.Row==row&&c.Column==column).ShapeId;return slide.Shapes.Cast<PowerPoint.Shape>().Single(s=>s.Id==id);}
    internal static void Run(PowerPoint.Application app,string outputDirectory)
    {
        if(app==null)throw new ArgumentNullException("app");Directory.CreateDirectory(outputDirectory);count=0;PowerPoint.Presentation owned=null;string file=Path.Combine(Path.GetFullPath(outputDirectory),"process-flow-native-"+Guid.NewGuid().ToString("N")+".pptx");
        Console.WriteLine("PROCESS FLOW Office="+app.Version+"; process="+(Environment.Is64BitProcess?"x64":"x86"));
        try
        {
            owned=app.Presentations.Add(Office.MsoTriState.msoTrue);owned.PageSetup.SlideWidth=960;owned.PageSetup.SlideHeight=540;owned.Windows[1].ViewType=PowerPoint.PpViewType.ppViewNormal;var slide=owned.Slides.Add(1,PowerPoint.PpSlideLayout.ppLayoutBlank);slide.Select();var outside=slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,40,480,500,30);outside.TextFrame.TextRange.Text="Unrelated user object";int outsideId=outside.Id;slide.Select();var source=TextShapeElements.Insert(app,"roundedRectangle");source.TextFrame.TextRange.Font.Name="Arial";source.TextFrame.TextRange.Font.Size=18;source.Fill.ForeColor.RGB=0xDDEEFF;string sourceSettings=TextShapeCodec.Encode(TextShapeElements.Read(source));int sourceId=source.Id;source.Select();var first=ProcessFlowTools.Create(app);string gridId=first.Tags[ProcessFlowTools.GridTag];var grid=State(owned,gridId);
            Check(TextShapeCodec.Encode(TextShapeElements.Read(source))==sourceSettings,"selected style source metadata survives all cell creation");
            var styled=grid.Cells.Select(c=>Cell(slide,grid,c.Row,c.Column)).ToArray();Check(styled.All(s=>TextShapeElements.Read(s).AutoFit==false&&s.TextFrame.TextRange.Font.Name=="Arial"&&s.TextFrame.TextRange.Font.Size==18&&s.Fill.ForeColor.RGB==0xDDEEFF),"native style transfer retains appearance and verified cell metadata");Check(styled.SelectMany(s=>s.CustomerData.Cast<Office.CustomXMLPart>().Where(p=>p.NamespaceURI==TextShapeCodec.Namespace).Select(p=>p.Id)).Distinct().Count()==6,"each styled cell owns an independent text metadata part");
            Check(grid.Rows==2&&grid.Columns==3&&grid.Cells.Count==6&&grid.Cells.All(c=>Cell(slide,grid,c.Row,c.Column).TextFrame.TextRange.Text==""),"default empty two by three grid contains native editable fields");
            Check(Cell(slide,grid,0,0).AutoShapeType==Office.MsoAutoShapeType.msoShapePentagon&&Cell(slide,grid,0,1).AutoShapeType==Office.MsoAutoShapeType.msoShapeChevron&&TextShapeElements.Read(Cell(slide,grid,1,0)).Kind=="textbox","native header and body geometry");
            var body=Cell(slide,grid,1,1);body.TextFrame.TextRange.Text="Editable body";int bodyId=body.Id;body.Select();ProcessFlowTools.AddRow(app);grid=State(owned,gridId);Check(grid.Rows==3&&grid.Cells.Single(c=>c.ShapeId==bodyId).Row==1,"add row retains native content and identities");
            Cell(slide,grid,0,0).Select();ProcessFlowTools.AddColumn(app);grid=State(owned,gridId);Check(grid.Columns==4&&grid.Cells.Single(c=>c.ShapeId==bodyId).Column==2&&slide.Shapes.Cast<PowerPoint.Shape>().Single(s=>s.Id==bodyId).TextFrame.TextRange.Text=="Editable body","add column shifts mapping without losing body text");
            int oldFirst=Cell(slide,grid,0,0).Id;Cell(slide,grid,0,0).Select();ProcessFlowTools.ReorderColumn(app,true);grid=State(owned,gridId);Check(grid.Cells.Single(c=>c.ShapeId==oldFirst).Column==1&&Cell(slide,grid,0,0).AutoShapeType==Office.MsoAutoShapeType.msoShapePentagon&&Cell(slide,grid,0,1).AutoShapeType==Office.MsoAutoShapeType.msoShapeChevron,"whole-column reorder updates header kinds in place");
            Cell(slide,grid,0,0).Select();ProcessFlowTools.Stripe(app,"row",0xDDEEFF,0xCCDDEE);Check(Cell(slide,grid,0,0).Fill.ForeColor.RGB==0xDDEEFF&&Cell(slide,grid,1,0).Fill.ForeColor.RGB==0xCCDDEE,"alternating row fills apply to native cells");
            Cell(slide,grid,0,0).Select();ProcessFlowTools.Stripe(app,"column",0xABCDEF,0xFEDCBA);Check(Cell(slide,grid,0,0).Fill.ForeColor.RGB==0xABCDEF&&Cell(slide,grid,0,1).Fill.ForeColor.RGB==0xFEDCBA,"alternating column fills apply to native cells");
            int removed=Cell(slide,grid,2,0).Id;Cell(slide,grid,2,0).Select();ProcessFlowTools.DeleteSelectedRowOrColumn(app,false);grid=State(owned,gridId);Check(grid.Rows==2&&!grid.Cells.Single(c=>c.ShapeId==removed).Active&&slide.Shapes.Cast<PowerPoint.Shape>().Single(s=>s.Id==removed).Visible==Office.MsoTriState.msoFalse,"delete row safely retains hidden native history");
            Cell(slide,grid,0,0).Select();ProcessFlowTools.DeleteSelectedRowOrColumn(app,true);grid=State(owned,gridId);Check(grid.Columns==3&&grid.Cells.Count(c=>c.Active)==6&&Cell(slide,grid,0,0).AutoShapeType==Office.MsoAutoShapeType.msoShapePentagon,"delete first column retains rectangular active grid");
            var damaged=Cell(slide,grid,1,0);string previous=damaged.Tags[ProcessFlowTools.ColumnTag];damaged.Tags.Add(ProcessFlowTools.ColumnTag,"99");string priorXml=owned.CustomXMLParts.SelectByNamespace(ProcessGridCodec.Namespace)[1].XML;int shapeCount=slide.Shapes.Count;Cell(slide,grid,0,0).Select();bool refused=false;try{ProcessFlowTools.AddRow(app);}catch(InvalidDataException){refused=true;}Check(refused&&shapeCount==slide.Shapes.Count&&owned.CustomXMLParts.SelectByNamespace(ProcessGridCodec.Namespace)[1].XML==priorXml,"corrupt cell preflight preserves grid and outside objects");damaged.Tags.Add(ProcessFlowTools.ColumnTag,previous);
            Check(slide.Shapes.Cast<PowerPoint.Shape>().Single(s=>s.Id==outsideId).TextFrame.TextRange.Text=="Unrelated user object","outside object remains untouched");
            owned.SaveAs(file,PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation,Office.MsoTriState.msoFalse);owned.Close();owned=null;owned=app.Presentations.Open(file,Office.MsoTriState.msoFalse,Office.MsoTriState.msoFalse,Office.MsoTriState.msoTrue);slide=owned.Slides[1];grid=State(owned,gridId);Check(TextShapeCodec.Encode(TextShapeElements.Read(slide.Shapes.Cast<PowerPoint.Shape>().Single(s=>s.Id==sourceId)))==sourceSettings,"style source settings survive grid save reopen");Cell(slide,grid,1,0).Select();ProcessFlowTools.Refresh(app);Check(grid.Rows==2&&grid.Columns==3&&grid.Cells.Count(c=>!c.Active)>0,"shared versioned grid and retained history survive save/reopen");
            var edit=Cell(slide,grid,1,0);edit.TextFrame.TextRange.Text="Still editable";Cell(slide,grid,0,0).Select();ProcessFlowTools.ReorderColumn(app,true);grid=State(owned,gridId);Check(slide.Shapes.Cast<PowerPoint.Shape>().Single(s=>s.Id==edit.Id).TextFrame.TextRange.Text=="Still editable","native editing and reordering work after reopen");
            Cell(slide,grid,0,0).Select();ProcessFlowTools.Resize(app,80,80,800,160);grid=State(owned,gridId);Check(Math.Abs(Cell(slide,grid,0,0).Left-80)<.1&&Math.Abs(Cell(slide,grid,0,0).Top-80)<.1&&grid.Width==800,"whole-grid resize repositions cells while preserving identity");
            owned.Save();Console.WriteLine("PROCESS FLOW "+count+" native assertions passed. Evidence deck: "+file);
        }
        finally{if(owned!=null){owned.Saved=Office.MsoTriState.msoTrue;owned.Close();}}
    }
}

