using System;
using System.IO;
using Office=Microsoft.Office.Core;
using PowerPoint=Microsoft.Office.Interop.PowerPoint;
using PowerPointIntelligence.AddIn;
internal static class ProductivityToolOfficeTests
{
    internal static int Run(PowerPoint.Application app,string output)
    {
        if(app==null)throw new ArgumentNullException("app");if(String.IsNullOrWhiteSpace(output))throw new ArgumentException("Choose an evidence folder.");Directory.CreateDirectory(output);int count=0;Action<bool,string> check=(pass,name)=>{if(!pass)throw new Exception("FAIL productivity native "+name);Console.WriteLine("PASS productivity native "+name);count++;};PowerPoint.Presentation deck=null;string file=Path.Combine(output,"productivity-native.pptx");
        try
        {
            deck=app.Presentations.Add(Office.MsoTriState.msoTrue);deck.PageSetup.SlideWidth=960;deck.PageSetup.SlideHeight=540;deck.Windows[1].Activate();deck.Windows[1].ViewType=PowerPoint.PpViewType.ppViewNormal;var slide=deck.Slides.Add(1,PowerPoint.PpSlideLayout.ppLayoutBlank);slide.Select();
            var text=slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,80,100,400,90);int textId=text.Id;text.Name="Productivity text fixture";text.TextFrame.TextRange.Text="A1.2B";text.TextFrame.TextRange.Font.Size=22;var separator=text.TextFrame.TextRange.Characters(3,1);separator.Font.Size=11;separator.Font.Bold=Office.MsoTriState.msoTrue;separator.Font.Color.RGB=0x332211;text.Select();check(ProductivityTools.SwitchDecimals(app,false)==1,"one selected separator changed");check(text.TextFrame.TextRange.Text=="A1,2B","decimal changed in native text");separator=text.TextFrame.TextRange.Characters(3,1);check(separator.Font.Size==11&&separator.Font.Bold==Office.MsoTriState.msoTrue&&separator.Font.Color.RGB==0x332211,"changed character retains run format");check(text.TextFrame.TextRange.Characters(1,1).Font.Size==22,"adjacent format retained");
            text.TextFrame.TextRange.Characters(1,1).Select();ProductivityTools.InsertSymbol(app,"→");check(text.TextFrame.TextRange.Text=="→1,2B","selected text replaced with symbol");app.ActiveWindow.Selection.Unselect();check(app.ActiveWindow.Selection.Type!=PowerPoint.PpSelectionType.ppSelectionText,"fixture leaves text caret before new-box insertion");int before=slide.Shapes.Count;ProductivityTools.InsertSymbol(app,"✓");check(slide.Shapes.Count==before+1,"symbol creates native textbox without text selection");
            text.Select();int originalCount=slide.Shapes.Count;var picture=ProductivityTools.RasterCopy(app);int pictureId=picture.Id;check(slide.Shapes.Count==originalCount+1&&picture.Type==Office.MsoShapeType.msoPicture,"raster picture added");check(text.Id==textId&&text.TextFrame.TextRange.Text=="→1,2B","original editable object retained");check(Math.Abs(picture.Width-text.Width)<.1&&Math.Abs(picture.Height-text.Height)<.1,"raster picture bounds retained");
            deck.SaveAs(file,PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);deck.Close();deck=null;deck=app.Presentations.Open(file,Office.MsoTriState.msoFalse,Office.MsoTriState.msoFalse,Office.MsoTriState.msoTrue);var reopened=deck.Slides[1];PowerPoint.Shape savedText=null,savedPicture=null;for(int i=1;i<=reopened.Shapes.Count;i++){if(reopened.Shapes[i].Id==textId)savedText=reopened.Shapes[i];if(reopened.Shapes[i].Id==pictureId)savedPicture=reopened.Shapes[i];}check(savedText!=null&&savedPicture!=null,"native source and picture survive reopen");check(savedText.TextFrame.TextRange.Text=="→1,2B","edited text persists");savedText.TextFrame.TextRange.Text="Still editable";check(savedText.TextFrame.TextRange.Text=="Still editable","source editable after reopen");
            Console.WriteLine(count+" productivity native assertions passed. Picker/workbook-dialog UI, caret insertion and raster visual equality require separate desktop checks.");return count;
        }
        finally{if(deck!=null){deck.Saved=Office.MsoTriState.msoTrue;deck.Close();}}
    }
}
