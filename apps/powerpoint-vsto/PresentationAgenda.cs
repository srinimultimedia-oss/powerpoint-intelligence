using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using Office=Microsoft.Office.Core;
using PowerPoint=Microsoft.Office.Interop.PowerPoint;
namespace PowerPointIntelligence.AddIn
{
    internal sealed class AgendaSlideInfo
    {
        internal int SlideId,SlideIndex,SectionIndex;
        internal string Title;
        internal bool Hidden,IsAgenda;
    }
    internal sealed class AgendaSectionInfo
    {
        internal int SectionIndex;
        internal string Name;
    }
    internal sealed class AgendaEntry
    {
        internal int TargetSlideId,TargetSlideIndex;
        internal string Caption,Key;
        internal bool Highlight;
        internal float Left,Top,Width,Height;
        internal string SubAddress {get{return TargetSlideId.ToString(CultureInfo.InvariantCulture)+","+TargetSlideIndex.ToString(CultureInfo.InvariantCulture)+",";}}
    }
    internal sealed class AgendaPlan
    {
        internal bool UsesSections;
        internal float Width,Height;
        internal IReadOnlyList<AgendaEntry> Entries;
    }
    internal static class AgendaPlanner
    {
        internal static AgendaPlan Build(IEnumerable<AgendaSlideInfo> slides,IEnumerable<AgendaSectionInfo> sections,float width,float height,int currentSlideId=0,bool highlightCurrent=true)
        {
            if(slides==null||sections==null)throw new ArgumentNullException("Agenda sources are required.");
            if(!Finite(width)||!Finite(height)||width<400||height<300||width>10000||height>10000)throw new InvalidDataException("Agenda needs valid slide dimensions of at least 400 by 300 points.");
            var all=slides.ToArray();var chapters=sections.ToArray();
            if(all.Length>10000||all.Any(s=>s==null||s.SlideId<1||s.SlideIndex<1||s.SectionIndex<0)||all.Select(s=>s.SlideId).Distinct().Count()!=all.Length||all.Select(s=>s.SlideIndex).Distinct().Count()!=all.Length)throw new InvalidDataException("Invalid or duplicate agenda slide identities.");
            if(chapters.Any(s=>s==null||s.SectionIndex<1)||chapters.Select(s=>s.SectionIndex).Distinct().Count()!=chapters.Length)throw new InvalidDataException("Invalid or duplicate agenda sections.");
            var usable=all.Where(s=>!s.Hidden&&!s.IsAgenda).OrderBy(s=>s.SlideIndex).ToArray();var current=all.SingleOrDefault(s=>s.SlideId==currentSlideId);var entries=new List<AgendaEntry>();
            if(chapters.Length>0){if(usable.Any(s=>s.SectionIndex>0&&!chapters.Any(c=>c.SectionIndex==s.SectionIndex)))throw new InvalidDataException("A slide references an unknown section.");var unsectioned=usable.FirstOrDefault(s=>s.SectionIndex==0);if(unsectioned!=null)entries.Add(Entry(unsectioned,"Unsectioned slides",highlightCurrent&&current!=null&&current.SectionIndex==0));foreach(var chapter in chapters.OrderBy(s=>s.SectionIndex)){var first=usable.FirstOrDefault(s=>s.SectionIndex==chapter.SectionIndex);if(first==null)continue;entries.Add(Entry(first,chapter.Name,highlightCurrent&&current!=null&&current.SectionIndex==chapter.SectionIndex));}}
            else foreach(var slide in usable)entries.Add(Entry(slide,String.IsNullOrWhiteSpace(slide.Title)?"Slide "+slide.SlideIndex.ToString(CultureInfo.InvariantCulture):slide.Title,highlightCurrent&&slide.SlideId==currentSlideId));
            if(entries.Count==0)throw new InvalidOperationException("No visible content slides are available for an agenda.");
            float left=40,top=128,bottom=40,gap=24;int rows=Math.Max(1,(int)Math.Floor((height-top-bottom)/34));int columns=(int)Math.Ceiling(entries.Count/(double)rows);
            if(columns>3||entries.Count>90)throw new InvalidOperationException("This agenda cannot fit on one slide at a readable size. Use fewer sections or a larger slide.");
            float cellWidth=(width-left*2-gap*(columns-1))/columns;if(cellWidth<160)throw new InvalidOperationException("Agenda columns would be too narrow. Use fewer sections.");
            int rowsUsed=(int)Math.Ceiling(entries.Count/(double)columns);float cellHeight=Math.Min(46,(height-top-bottom)/rowsUsed);
            for(int i=0;i<entries.Count;i++){var entry=entries[i];int column=i/rowsUsed,row=i%rowsUsed;entry.Left=left+column*(cellWidth+gap);entry.Top=top+row*cellHeight;entry.Width=cellWidth;entry.Height=cellHeight-4;}
            return new AgendaPlan {UsesSections=chapters.Length>0,Width=width,Height=height,Entries=entries.AsReadOnly()};
        }
        private static AgendaEntry Entry(AgendaSlideInfo slide,string text,bool highlight)
        {string caption=(text??"").Replace('\r',' ').Replace('\n',' ').Trim();if(caption.Length==0||caption.Length>240)throw new InvalidDataException("Agenda names must contain 1 to 240 characters. Shorten the section name or slide title.");return new AgendaEntry {TargetSlideId=slide.SlideId,TargetSlideIndex=slide.SlideIndex,Caption=caption,Key="item:"+slide.SlideId.ToString(CultureInfo.InvariantCulture),Highlight=highlight};}
        private static bool Finite(float value){return !Single.IsNaN(value)&&!Single.IsInfinity(value);}
    }
    internal static class PresentationAgenda
    {
        internal const string SchemaTag="PPI_AGENDA_SCHEMA",OwnerTag="PPI_AGENDA_ID",RoleTag="PPI_AGENDA_ROLE",ContextTag="PPI_AGENDA_CONTEXT";
        private static AgendaPlan Capture(PowerPoint.Presentation deck,int context,bool highlight)
        {
            var sections=new List<AgendaSectionInfo>();var ranges=new List<Tuple<int,int,int>>();var source=deck.SectionProperties;
            for(int i=1;i<=source.Count;i++){sections.Add(new AgendaSectionInfo {SectionIndex=i,Name=source.Name(i)});int first=source.FirstSlide(i);int count=source.SlidesCount(i);if(first>0&&count>0)ranges.Add(Tuple.Create(i,first,first+count-1));}
            var slides=new List<AgendaSlideInfo>();foreach(PowerPoint.Slide slide in deck.Slides){var range=ranges.FirstOrDefault(r=>slide.SlideIndex>=r.Item2&&slide.SlideIndex<=r.Item3);slides.Add(new AgendaSlideInfo {SlideId=slide.SlideID,SlideIndex=slide.SlideIndex,SectionIndex=range==null?0:range.Item1,Title=slide.Shapes.HasTitle==Office.MsoTriState.msoTrue?slide.Shapes.Title.TextFrame.TextRange.Text:null,Hidden=slide.SlideShowTransition.Hidden==Office.MsoTriState.msoTrue,IsAgenda=!String.IsNullOrEmpty(slide.Tags[SchemaTag])});}
            return AgendaPlanner.Build(slides,sections,deck.PageSetup.SlideWidth,deck.PageSetup.SlideHeight,context,highlight);
        }
        internal static PowerPoint.Slide Insert(PowerPoint.Application app,bool highlightCurrent=true)
        {
            var active=Active(app);var deck=app.ActivePresentation;int context=active.SlideID;Capture(deck,context,highlightCurrent);app.StartNewUndoEntry();PowerPoint.Slide created=null;
            try{created=deck.Slides.Add(active.SlideIndex+1,PowerPoint.PpSlideLayout.ppLayoutBlank);created.Design=active.Design;created.FollowMasterBackground=Office.MsoTriState.msoTrue;created.Tags.Add(SchemaTag,"1");created.Tags.Add(OwnerTag,Guid.NewGuid().ToString("D"));created.Tags.Add(ContextTag,context.ToString(CultureInfo.InvariantCulture));Apply(created,Capture(deck,context,highlightCurrent));created.Select();return created;}
            catch(Exception original){if(created!=null)try{created.Delete();}catch{throw new InvalidOperationException("Agenda insertion failed and cleanup was incomplete. Use PowerPoint Undo.",original);}throw;}
        }
        internal static void RefreshSelected(PowerPoint.Application app,bool highlightCurrent=true)
        {
            var slide=Active(app);Guid owner;if(slide.Tags[SchemaTag]!="1"||!Guid.TryParse(slide.Tags[OwnerTag],out owner))throw new InvalidOperationException("Select a managed agenda slide to refresh.");int context;Int32.TryParse(slide.Tags[ContextTag],NumberStyles.None,CultureInfo.InvariantCulture,out context);var plan=Capture(app.ActivePresentation,context,highlightCurrent);app.StartNewUndoEntry();Apply(slide,plan);
        }
        private static PowerPoint.Slide Active(PowerPoint.Application app)
        {if(app==null||app.Presentations.Count==0)throw new InvalidOperationException("Open a presentation first.");if(app.ActiveWindow.ViewType!=PowerPoint.PpViewType.ppViewNormal)throw new InvalidOperationException("Select a slide in Normal view first.");var slide=app.ActiveWindow.View.Slide as PowerPoint.Slide;if(slide==null)throw new InvalidOperationException("Select a slide in Normal view first.");return slide;}
        private static bool ContainsOwnedField(PowerPoint.Shape group,string owner)
        {foreach(PowerPoint.Shape shape in group.GroupItems){if(String.Equals(shape.Tags[OwnerTag],owner,StringComparison.OrdinalIgnoreCase))return true;if(shape.Type==Office.MsoShapeType.msoGroup&&ContainsOwnedField(shape,owner))return true;}return false;}
        private static void Apply(PowerPoint.Slide slide,AgendaPlan plan)
        {
            string owner=slide.Tags[OwnerTag];var existing=new Dictionary<string,PowerPoint.Shape>(StringComparer.OrdinalIgnoreCase);
            foreach(PowerPoint.Shape shape in slide.Shapes){if(shape.Type==Office.MsoShapeType.msoGroup&&ContainsOwnedField(shape,owner))throw new InvalidOperationException("Ungroup managed agenda fields before refreshing. No objects were changed.");if(String.Equals(shape.Tags[OwnerTag],owner,StringComparison.OrdinalIgnoreCase)){string role=shape.Tags[RoleTag];if(String.IsNullOrEmpty(role)||existing.ContainsKey(role))throw new InvalidOperationException("Managed agenda fields are duplicated or damaged. Preserve the slide and create a new agenda.");existing.Add(role,shape);}}
            if(existing.Count+plan.Entries.Select(e=>e.Key).Concat(new[]{"title","subtitle"}).Count(role=>!existing.ContainsKey(role))>256)throw new InvalidOperationException("Agenda history exceeds 256 managed fields. Create a new agenda to retain old edits safely.");
            var restore=new List<Action>();var active=new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try{
                Upsert(slide,existing,active,restore,owner,"title","Agenda",40,36,plan.Width-80,48,30,false,null);
                Upsert(slide,existing,active,restore,owner,"subtitle",plan.UsesSections?"Presentation sections":"Visible slide titles",40,91,plan.Width-80,22,12,false,null);
                for(int i=0;i<plan.Entries.Count;i++){var item=plan.Entries[i];Upsert(slide,existing,active,restore,owner,item.Key,(i+1).ToString("D2",CultureInfo.InvariantCulture)+"   "+item.Caption,item.Left,item.Top,item.Width,item.Height,18,item.Highlight,item);}
                foreach(var pair in existing.Where(p=>!active.Contains(p.Key))){var shape=pair.Value;var previous=shape.Visible;restore.Add(()=>shape.Visible=previous);shape.Visible=Office.MsoTriState.msoFalse;}
            }catch(Exception original){bool restored=true;for(int i=restore.Count-1;i>=0;i--)try{restore[i]();}catch{restored=false;}if(!restored)throw new InvalidOperationException("Agenda update failed and rollback was incomplete. Use PowerPoint Undo.",original);throw;}
        }
        private static void Upsert(PowerPoint.Slide slide,IDictionary<string,PowerPoint.Shape> existing,ISet<string> active,List<Action> restore,string owner,string role,string text,float left,float top,float width,float height,float fontSize,bool highlight,AgendaEntry link)
        {
            PowerPoint.Shape shape;bool created=!existing.TryGetValue(role,out shape);if(created){shape=slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,left,top,width,height);var added=shape;restore.Add(()=>added.Delete());shape.Tags.Add(OwnerTag,owner);shape.Tags.Add(RoleTag,role);shape.Fill.Visible=Office.MsoTriState.msoFalse;shape.Line.Visible=Office.MsoTriState.msoFalse;shape.TextFrame.TextRange.Font.Name="Segoe UI";shape.TextFrame.TextRange.Font.Size=fontSize;shape.TextFrame.TextRange.Font.Color.ObjectThemeColor=Office.MsoThemeColorIndex.msoThemeColorText1;shape.TextFrame.MarginLeft=shape.TextFrame.MarginRight=shape.TextFrame.MarginTop=shape.TextFrame.MarginBottom=0;shape.TextFrame.WordWrap=Office.MsoTriState.msoTrue;shape.TextFrame.VerticalAnchor=Office.MsoVerticalAnchor.msoAnchorMiddle;shape.TextFrame.TextRange.ParagraphFormat.Alignment=PowerPoint.PpParagraphAlignment.ppAlignLeft;shape.TextFrame2.AutoSize=Office.MsoAutoSize.msoAutoSizeTextToFitShape;}
            else{
                var target=shape;float x=shape.Left,y=shape.Top,w=shape.Width,h=shape.Height;string value=shape.TextFrame.TextRange.Text;var visible=shape.Visible;var bold=shape.TextFrame.TextRange.Font.Bold;var action=shape.ActionSettings[PowerPoint.PpMouseActivation.ppMouseClick];var priorAction=action.Action;string address=action.Hyperlink.Address,sub=action.Hyperlink.SubAddress,screen=action.Hyperlink.ScreenTip;
                restore.Add(()=>{target.Left=x;target.Top=y;target.Width=w;target.Height=h;target.TextFrame.TextRange.Text=value;target.Visible=visible;target.TextFrame.TextRange.Font.Bold=bold;var click=target.ActionSettings[PowerPoint.PpMouseActivation.ppMouseClick];click.Action=priorAction;click.Hyperlink.Address=address;click.Hyperlink.SubAddress=sub;click.Hyperlink.ScreenTip=screen;});
            }
            bool header=role=="title"||role=="subtitle";
            if(created||!header){shape.Left=left;shape.Top=top;shape.Width=width;shape.Height=height;if(shape.TextFrame.TextRange.Text!=text)shape.TextFrame.TextRange.Text=text;shape.TextFrame.TextRange.Font.Bold=highlight||role=="title"?Office.MsoTriState.msoTrue:Office.MsoTriState.msoFalse;}shape.Visible=Office.MsoTriState.msoTrue;
            if(link!=null){var action=shape.ActionSettings[PowerPoint.PpMouseActivation.ppMouseClick];action.Action=PowerPoint.PpActionType.ppActionHyperlink;action.Hyperlink.Address="";action.Hyperlink.SubAddress=link.SubAddress;action.Hyperlink.ScreenTip="Go to "+link.Caption;}
            active.Add(role);
        }
    }
}
