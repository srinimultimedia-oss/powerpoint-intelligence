using System;
using System.Globalization;
using System.IO;
using System.Linq;
using PowerPointIntelligence.AddIn;
using Office=Microsoft.Office.Core;
using PowerPoint=Microsoft.Office.Interop.PowerPoint;
internal static class ChapterAgendaOfficeTests
{
    private static int count;
    private static void Require(bool value,string name){if(!value)throw new Exception("FAIL chapter agenda Office: "+name);count++;Console.WriteLine("PASS chapter agenda Office: "+name);}
    private static PowerPoint.Shape Field(PowerPoint.Slide slide,string role){return slide.Shapes.Cast<PowerPoint.Shape>().Single(s=>String.Equals(s.Tags[PresentationAgenda.RoleTag],role,StringComparison.OrdinalIgnoreCase)&&String.Equals(s.Tags[PresentationAgenda.OwnerTag],slide.Tags[PresentationAgenda.OwnerTag],StringComparison.OrdinalIgnoreCase));}
    private static PowerPoint.Shape Row(PowerPoint.Slide slide,int id){return Field(slide,"chapter:"+id.ToString(CultureInfo.InvariantCulture));}
    private static PowerPoint.Slide Content(PowerPoint.Presentation deck,PowerPoint.Slide after,string text)
    {
        var slide=deck.Slides.Add(after.SlideIndex+1,PowerPoint.PpSlideLayout.ppLayoutBlank);var shape=slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,40,40,500,80);shape.TextFrame.TextRange.Text=text;slide.Select();return slide;
    }
    private static void Refused(Action action,string name){bool refused=false;try{action();}catch(Exception ex) when(ex is InvalidDataException||ex is InvalidOperationException){refused=true;}Require(refused,name);}
    internal static void Run(PowerPoint.Application app,string outputDirectory)
    {
        if(app==null)throw new ArgumentNullException("app");Directory.CreateDirectory(outputDirectory);count=0;
        string file=Path.Combine(Path.GetFullPath(outputDirectory),"chapter-agenda-native-"+Guid.NewGuid().ToString("N")+".pptx");PowerPoint.Presentation owned=null;
        Console.WriteLine("CHAPTER AGENDA Office="+app.Version+"; process="+(Environment.Is64BitProcess?"x64":"x86"));
        try
        {
            owned=app.Presentations.Add(Office.MsoTriState.msoTrue);owned.PageSetup.SlideWidth=960;owned.PageSetup.SlideHeight=540;owned.Windows[1].ViewType=PowerPoint.PpViewType.ppViewNormal;
            var start=owned.Slides.Add(1,PowerPoint.PpSlideLayout.ppLayoutBlank);start.Select();int initial=owned.Slides.Count;
            Refused(()=>ChapterAgenda.InsertChapter(app,"Invalid jump",3),"invalid insertion hierarchy refused");Require(owned.Slides.Count==initial,"invalid insertion adds no slides");
            var a=ChapterAgenda.InsertChapter(app,"Alpha",1);int aid=a.SlideID;var ac=Content(owned,a,"Alpha owned content");int acid=ac.SlideID;
            var child=ChapterAgenda.InsertChapter(app,"Alpha detail",2);int childid=child.SlideID;var cc=Content(owned,child,"Child owned content");int ccid=cc.SlideID;
            var b=ChapterAgenda.InsertChapter(app,"Beta",1);int bid=b.SlideID;var bc=Content(owned,b,"Beta owned content");int bcid=bc.SlideID;
            ChapterAgenda.RefreshAll(app);Require(ChapterAgenda.Read(app).Chapters.Select(r=>r.SlideId).SequenceEqual(new[]{aid,childid,bid}),"explicit chapter model follows actual slides");
            b.Select();var toc=ChapterAgenda.InsertContents(app);int tocid=toc.SlideID;
            Require(toc.SlideIndex<a.SlideIndex&&ChapterAgenda.IsManaged(toc)&&toc.Tags[PresentationAgenda.SchemaTag]=="2","contents inserted before first chapter");
            Require(Row(toc,aid).Visible==Office.MsoTriState.msoTrue&&Row(toc,bid).Visible==Office.MsoTriState.msoTrue,"contents exposes root chapters");
            Require(!toc.Shapes.Cast<PowerPoint.Shape>().Any(s=>s.Visible==Office.MsoTriState.msoTrue&&String.Equals(s.Tags[PresentationAgenda.RoleTag],"chapter:"+childid,StringComparison.OrdinalIgnoreCase)),"contents hides nested levels by default");
            Require(Row(toc,aid).Tags["PPI_AGENDA_HIGHLIGHT"]=="0"&&Row(toc,bid).Tags["PPI_AGENDA_HIGHLIGHT"]=="0","contents does not highlight chapters");
            Row(toc,aid).TextFrame.TextRange.Text="Alpha renamed natively";ChapterAgenda.RefreshAll(app);
            Require(ChapterAgenda.Read(app).Chapters.Single(r=>r.SlideId==aid).Title=="Alpha renamed natively"&&Row(a,aid).TextFrame.TextRange.Text=="Alpha renamed natively"&&Row(b,aid).TextFrame.TextRange.Text=="Alpha renamed natively","native caption edit propagates across views");
            string canonical=Row(a,aid).TextFrame.TextRange.Text;Row(a,aid).TextFrame.TextRange.Text="Conflicting A";Row(toc,aid).TextFrame.TextRange.Text="Conflicting B";string priorXml=owned.CustomXMLParts.SelectByNamespace(ChapterAgendaCodec.Namespace)[1].XML;int before=owned.Slides.Count;
            Refused(()=>ChapterAgenda.RefreshAll(app),"conflicting native edits refused");Require(owned.Slides.Count==before&&owned.CustomXMLParts.SelectByNamespace(ChapterAgendaCodec.Namespace)[1].XML==priorXml&&Row(a,aid).TextFrame.TextRange.Text=="Conflicting A"&&Row(toc,aid).TextFrame.TextRange.Text=="Conflicting B","conflict preserves edits and registry");
            Row(a,aid).TextFrame.TextRange.Text=canonical;Row(toc,aid).TextFrame.TextRange.Text=canonical;ChapterAgenda.RefreshAll(app);
            Field(toc,"title").TextFrame.TextRange.Text="Custom contents heading";Row(toc,aid).TextFrame.TextRange.Font.Name="Arial";Row(toc,aid).TextFrame.TextRange.Font.Size=21;
            var note=toc.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,30,490,400,25);note.Name="chapter_workflow_note";note.TextFrame.TextRange.Text="Unrelated native note";
            ChapterAgenda.RefreshAll(app);Require(Field(toc,"title").TextFrame.TextRange.Text=="Custom contents heading"&&Row(toc,aid).TextFrame.TextRange.Font.Name=="Arial"&&Math.Abs(Row(toc,aid).TextFrame.TextRange.Font.Size-21)<.1,"title and native font survive refresh");
            ChapterAgenda.ChangeChapter(app,childid,"Promoted detail",1);Require(ChapterAgenda.Read(app).Chapters.Single(r=>r.SlideId==childid).Level==1&&Row(toc,childid).Visible==Office.MsoTriState.msoTrue,"hierarchy promotion updates contents");
            ChapterAgenda.ChangeChapter(app,childid,"Alpha detail",2);Require(ChapterAgenda.Read(app).Chapters.Single(r=>r.SlideId==childid).Level==2&&Row(toc,childid).Visible==Office.MsoTriState.msoFalse,"hierarchy demotion retains hidden native field");
            before=owned.Slides.Count;Refused(()=>ChapterAgenda.ChangeChapter(app,aid,canonical,2),"missing-parent demotion refused");Require(owned.Slides.Count==before&&ChapterAgenda.Read(app).Chapters.Single(r=>r.SlideId==aid).Level==1,"invalid change preserves model");
            var duplicate=b.Duplicate()[1];duplicate.MoveTo(owned.Slides.Count);int duplicateid=duplicate.SlideID;Row(duplicate,bid).TextFrame.TextRange.Text="Copied Beta";ChapterAgenda.RefreshAll(app);
            Require(duplicate.Tags[ChapterAgenda.Anchor]==duplicateid.ToString(CultureInfo.InvariantCulture)&&ChapterAgenda.Read(app).Chapters.Any(r=>r.SlideId==duplicateid&&r.Title=="Copied Beta"),"duplicated chapter gets actual slide anchor and canonical record");
            Require(Row(duplicate,duplicateid).Visible==Office.MsoTriState.msoTrue&&Row(b,bid).TextFrame.TextRange.Text=="Beta","duplicate gets independent editable heading without renaming the source");
            ChapterAgenda.Move(app,aid,true);
            Require(owned.Slides.FindBySlideID(bid).SlideIndex<owned.Slides.FindBySlideID(aid).SlideIndex&&owned.Slides.FindBySlideID(acid).SlideIndex==owned.Slides.FindBySlideID(aid).SlideIndex+1&&owned.Slides.FindBySlideID(childid).SlideIndex==owned.Slides.FindBySlideID(acid).SlideIndex+1&&owned.Slides.FindBySlideID(ccid).SlideIndex==owned.Slides.FindBySlideID(childid).SlideIndex+1,"chapter move retains descendant and content IDs as contiguous block");
            Require(owned.Slides.FindBySlideID(bcid).SlideIndex==owned.Slides.FindBySlideID(bid).SlideIndex+1,"sibling content remains attached");
            duplicate.Select();Refused(()=>ChapterAgenda.SplitHere(app),"split on managed chapter is refused before mutation");Require(String.IsNullOrEmpty(duplicate.Tags[ChapterAgenda.Split]),"refused split preserves chapter metadata");var splitter=Content(owned,duplicate,"Independent agenda scope");int splitid=splitter.SlideID;ChapterAgenda.SplitHere(app);var c=ChapterAgenda.InsertChapter(app,"Scoped Gamma",1);int cid=c.SlideID;
            Require(ChapterAgenda.Read(app).Chapters.Single(r=>r.SlideId==cid).Scope==splitid&&ChapterAgenda.Read(app).Chapters.Single(r=>r.SlideId==aid).Scope==0,"splitter isolates subsequent chapter scope");
            Refused(()=>ChapterAgenda.Move(app,cid,false),"chapter cannot move across splitter");
            var link=Row(toc,aid).ActionSettings[PowerPoint.PpMouseActivation.ppMouseClick];Require(link.Action==PowerPoint.PpActionType.ppActionHyperlink&&link.Hyperlink.SubAddress==aid+","+owned.Slides.FindBySlideID(aid).SlideIndex+",","agenda hyperlink follows actual moved slide identity");
            c.Select();var emptyParent=ChapterAgenda.InsertChapter(app,"Empty parent",1);var emptyChild=ChapterAgenda.InsertChapter(app,"Empty child",2);ChapterAgenda.RefreshAll(app);Require(emptyParent.SlideShowTransition.Hidden==Office.MsoTriState.msoTrue,"empty parent before immediate subchapter is hidden");var parentContent=Content(owned,emptyParent,"Parent now has content");ChapterAgenda.RefreshAll(app);Require(emptyParent.SlideShowTransition.Hidden==Office.MsoTriState.msoFalse,"parent unhides when content is inserted");parentContent.Delete();emptyChild.Delete();emptyParent.Delete();ChapterAgenda.RefreshAll(app);
            owned.SaveAs(file,PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation,Office.MsoTriState.msoFalse);owned.Close();owned=null;owned=app.Presentations.Open(file,Office.MsoTriState.msoFalse,Office.MsoTriState.msoFalse,Office.MsoTriState.msoTrue);
            toc=owned.Slides.FindBySlideID(tocid);a=owned.Slides.FindBySlideID(aid);c=owned.Slides.FindBySlideID(cid);var restored=ChapterAgenda.Read(app);
            Require(restored.Chapters.Count==5&&restored.Chapters.Single(r=>r.SlideId==childid).Level==2&&restored.Chapters.Single(r=>r.SlideId==cid).Scope==splitid,"chapter identity hierarchy scope persist after reopen");
            Require(Field(toc,"title").TextFrame.TextRange.Text=="Custom contents heading"&&toc.Shapes["chapter_workflow_note"].TextFrame.TextRange.Text=="Unrelated native note","custom title and unrelated object persist");
            Row(a,aid).TextFrame.TextRange.Text="Editable after reopen";ChapterAgenda.RefreshAll(app);Require(Row(toc,aid).TextFrame.TextRange.Text=="Editable after reopen","continued native edits propagate after reopen");
            Require(Row(toc,aid).TextFrame.TextRange.Font.Name=="Arial"&&Math.Abs(Row(toc,aid).TextFrame.TextRange.Font.Size-21)<.1,"native font retained after reopen and rename");
            owned.Save();Console.WriteLine("CHAPTER AGENDA "+count+" native assertions passed. Evidence deck: "+file);
        }
        finally{if(owned!=null){owned.Saved=Office.MsoTriState.msoTrue;owned.Close();}}
    }
}
