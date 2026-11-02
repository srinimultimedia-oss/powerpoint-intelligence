using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
namespace PowerPointIntelligence.AddIn
{
    internal sealed class ChapterRecord { internal int SlideId,SlideIndex,Level,Scope; internal string Title; }
    internal sealed class AgendaView { internal int SlideId,SlideIndex,Scope,CurrentChapterId; internal bool Contents,ShowAll; }
    internal sealed class ChapterCaptionEdit { internal int TargetSlideId; internal string Title; }
    internal sealed class ChapterAgendaState
    {
        internal List<ChapterRecord> Chapters=new List<ChapterRecord>();
        internal Dictionary<int,string> Baseline=new Dictionary<int,string>();
    }
    internal sealed class ChapterAgendaEntry
    {
        internal int TargetSlideId,TargetSlideIndex,Level;
        internal string Title,Number;
        internal bool Highlight,Ancestor;
        internal float Left,Top,Width,Height;
    }
    internal sealed class ChapterAgendaPlan { internal IReadOnlyList<ChapterAgendaEntry> Entries; }
    internal sealed class ChapterMovePlan { internal int FirstSlide,LastSlide,DestinationIndex; }
    internal static class ChapterAgendaPlanner
    {
        internal static ChapterMovePlan Move(IEnumerable<ChapterRecord> records,int chapterId,bool down,int totalSlides,IDictionary<int,int> scopeStartIndices)
        {
            var all=Validate(records);
            if(totalSlides<1||scopeStartIndices==null||scopeStartIndices.Count>10000||scopeStartIndices.Any(p=>p.Key<0||p.Value<1||p.Value>totalSlides||p.Key==0&&p.Value!=1)||scopeStartIndices.Where(p=>p.Key>0).Select(p=>p.Value).Distinct().Count()!=scopeStartIndices.Count(p=>p.Key>0))throw new InvalidDataException("Invalid chapter scope boundaries.");
            var starts=scopeStartIndices.Where(p=>p.Key>0).OrderBy(p=>p.Value).ToArray();
            foreach(var r in all)
            {
                if(r.SlideIndex>totalSlides||starts.Any(p=>p.Key==r.SlideId))throw new InvalidDataException("Chapter is outside the deck or shares a splitter identity.");
                int expected=0;foreach(var start in starts){if(r.SlideIndex==start.Value)throw new InvalidDataException("A splitter cannot also be a chapter slide.");if(r.SlideIndex>start.Value)expected=start.Key;else break;}
                if(r.Scope!=expected)throw new InvalidDataException("Chapter scope does not match the splitter boundaries.");
            }
            var selected=all.SingleOrDefault(r=>r.SlideId==chapterId);if(selected==null)throw new InvalidDataException("Unknown chapter to move.");
            var scoped=all.Where(r=>r.Scope==selected.Scope).ToArray();var parents=new Dictionary<int,int>();var stack=new int[4];
            foreach(var r in scoped){parents[r.SlideId]=r.Level==1?0:stack[r.Level-2];stack[r.Level-1]=r.SlideId;for(int i=r.Level;i<4;i++)stack[i]=0;}
            var siblings=scoped.Where(r=>r.Level==selected.Level&&parents[r.SlideId]==parents[selected.SlideId]).ToArray();int position=Array.IndexOf(siblings,selected),next=position+(down?1:-1);
            if(next<0||next>=siblings.Length)throw new InvalidOperationException("No sibling exists in that direction within this chapter scope.");
            int scopeStart=selected.Scope==0?1:scopeStartIndices[selected.Scope];int scopeEnd=starts.Where(p=>p.Value>scopeStart).Select(p=>p.Value-1).DefaultIfEmpty(totalSlides).First();
            Func<ChapterRecord,int> end=r=>scoped.Where(c=>c.SlideIndex>r.SlideIndex&&c.Level<=r.Level).Select(c=>c.SlideIndex-1).DefaultIfEmpty(scopeEnd).First();
            int last=end(selected),length=last-selected.SlideIndex+1;
            int destination=down?end(siblings[next])-length+1:siblings[next].SlideIndex;
            if(last<selected.SlideIndex||last>scopeEnd||destination<scopeStart||destination+length-1>scopeEnd||destination==selected.SlideIndex)throw new InvalidDataException("Chapter movement would cross or damage scope boundaries.");
            return new ChapterMovePlan {FirstSlide=selected.SlideIndex,LastSlide=last,DestinationIndex=destination};
        }
        internal static string Caption(string value)
        {
            string text=(value??"").Replace("\r\n"," ").Replace('\r',' ').Replace('\n',' ').Trim();
            if(text.Length<1||text.Length>240||text.Any(c=>Char.IsControl(c)))throw new InvalidDataException("Chapter captions require 1 to 240 printable characters; line breaks become spaces.");
            return text;
        }
        internal static ChapterRecord[] Validate(IEnumerable<ChapterRecord> records)
        {
            if(records==null)throw new ArgumentNullException("records");var all=records.Take(201).ToArray();
            if(all.Length>200||all.Any(r=>r==null||r.SlideId<1||r.SlideIndex<1||r.Scope<0||r.Level<1||r.Level>4)||all.Select(r=>r.SlideId).Distinct().Count()!=all.Length||all.Select(r=>r.SlideIndex).Distinct().Count()!=all.Length)throw new InvalidDataException("Invalid, duplicate or excessive chapter identities.");
            foreach(var r in all)Caption(r.Title);
            foreach(var scope in all.GroupBy(r=>r.Scope)){int prior=0;foreach(var r in scope.OrderBy(r=>r.SlideIndex)){if(r.Level>prior+1)throw new InvalidDataException("Chapter hierarchy requires an existing parent and cannot skip levels.");prior=r.Level;}}
            return all.OrderBy(r=>r.SlideIndex).ToArray();
        }
        internal static ChapterAgendaPlan Build(IEnumerable<ChapterRecord> records,AgendaView view,float width,float height)
        {
            var all=Validate(records);if(view==null||view.SlideId<1||view.SlideIndex<1||view.Scope<0)throw new InvalidDataException("Invalid agenda view identity.");
            if(Single.IsNaN(width)||Single.IsInfinity(width)||Single.IsNaN(height)||Single.IsInfinity(height)||width<400||height<300||width>10000||height>10000)throw new InvalidDataException("Agenda requires dimensions between 400 by 300 and 10000 by 10000 points.");
            var scoped=all.Where(r=>r.Scope==view.Scope).ToArray();var parents=new Dictionary<int,int>();var numbers=new Dictionary<int,string>();var stack=new int[4];var counts=new int[4];
            foreach(var r in scoped){parents[r.SlideId]=r.Level==1?0:stack[r.Level-2];stack[r.Level-1]=r.SlideId;counts[r.Level-1]++;for(int i=r.Level;i<4;i++){stack[i]=0;counts[i]=0;}numbers[r.SlideId]=String.Join(".",counts.Take(r.Level).Select(n=>n.ToString(CultureInfo.InvariantCulture)));}
            var current=scoped.SingleOrDefault(r=>r.SlideId==view.CurrentChapterId);if(!view.Contents&&current==null)throw new InvalidDataException("Chapter view references an unknown chapter in its scope.");
            var path=new HashSet<int>();if(current!=null){int id=current.SlideId;while(id!=0){path.Add(id);id=parents[id];}}
            int currentParent=current==null?0:parents[current.SlideId];
            var visible=scoped.Where(r=>view.ShowAll||(view.Contents?r.Level==1:r.Level==1||path.Contains(r.SlideId)||parents[r.SlideId]==currentParent||parents[r.SlideId]==current.SlideId)).ToArray();
            var entries=visible.Select(r=>new ChapterAgendaEntry {TargetSlideId=r.SlideId,TargetSlideIndex=r.SlideIndex,Level=r.Level,Title=Caption(r.Title),Number=numbers[r.SlideId],Highlight=!view.Contents&&r.SlideId==view.CurrentChapterId,Ancestor=!view.Contents&&r.SlideId!=view.CurrentChapterId&&path.Contains(r.SlideId)}).ToList();
            const float margin=32,top=104,gap=20,row=36;int maxRows=(int)Math.Floor((height-top-margin)/row);int columns=Math.Max(1,(int)Math.Ceiling(entries.Count/(double)maxRows));
            float cellWidth=(width-2*margin-gap*(columns-1))/columns;if(columns>3||cellWidth<180)throw new InvalidOperationException("Chapter agenda cannot fit without overflow. Use fewer entries or a larger slide.");
            int used=Math.Max(1,(int)Math.Ceiling(entries.Count/(double)columns));for(int i=0;i<entries.Count;i++){var e=entries[i];float indent=(e.Level-1)*12;e.Left=margin+(i/used)*(cellWidth+gap)+indent;e.Top=top+(i%used)*row;e.Width=cellWidth-indent;e.Height=row-4;}
            return new ChapterAgendaPlan {Entries=entries.AsReadOnly()};
        }
        internal static Dictionary<int,string> ResolveEdits(IDictionary<int,string> baseline,IEnumerable<ChapterCaptionEdit> edits)
        {
            if(baseline==null||edits==null)throw new ArgumentNullException("Caption baseline and edits are required.");if(baseline.Count>200||baseline.Keys.Any(k=>k<1))throw new InvalidDataException("Invalid caption baseline.");
            var result=baseline.ToDictionary(p=>p.Key,p=>Caption(p.Value));var changed=new Dictionary<int,string>();int count=0;
            foreach(var edit in edits){if(++count>10000||edit==null||!result.ContainsKey(edit.TargetSlideId))throw new InvalidDataException("Unknown or excessive chapter caption edits.");string text=Caption(edit.Title);if(text==Caption(baseline[edit.TargetSlideId]))continue;string prior;if(changed.TryGetValue(edit.TargetSlideId,out prior)&&prior!=text)throw new InvalidDataException("Conflicting chapter caption edits; choose one caption before synchronizing.");changed[edit.TargetSlideId]=text;}
            foreach(var pair in changed)result[pair.Key]=pair.Value;return result;
        }
    }
    internal static class ChapterAgendaCodec
    {
        internal const string Namespace="urn:powerpoint-intelligence:agenda:v2";
        private static readonly XNamespace Ns=Namespace;
        private static void State(ChapterAgendaState state)
        {
            if(state==null||state.Baseline==null)throw new InvalidDataException("Missing chapter state.");var records=ChapterAgendaPlanner.Validate(state.Chapters);if(state.Baseline.Count>200||state.Baseline.Keys.Any(k=>!records.Any(r=>r.SlideId==k)))throw new InvalidDataException("Caption baseline references an unknown chapter.");foreach(var value in state.Baseline.Values)ChapterAgendaPlanner.Caption(value);
        }
        internal static string Encode(ChapterAgendaState state)
        {
            State(state);var root=new XElement(Ns+"agenda",new XAttribute("version","2"),new XElement(Ns+"chapters",state.Chapters.OrderBy(r=>r.SlideIndex).Select(r=>new XElement(Ns+"chapter",new XAttribute("slideId",r.SlideId),new XAttribute("slideIndex",r.SlideIndex),new XAttribute("level",r.Level),new XAttribute("scope",r.Scope),new XAttribute("title",ChapterAgendaPlanner.Caption(r.Title))))),new XElement(Ns+"baseline",state.Baseline.OrderBy(p=>p.Key).Select(p=>new XElement(Ns+"caption",new XAttribute("slideId",p.Key),new XAttribute("title",ChapterAgendaPlanner.Caption(p.Value))))));return root.ToString(SaveOptions.DisableFormatting);
        }
        private static void Shape(XElement element,string name,params string[] attrs)
        {
            if(element.Name!=Ns+name||element.Attributes().Any(a=>!a.IsNamespaceDeclaration&&!attrs.Contains(a.Name.LocalName)||!a.IsNamespaceDeclaration&&a.Name.NamespaceName!="")||attrs.Any(a=>element.Attribute(a)==null)||element.Nodes().Any(n=>!(n is XElement)&&!(n is XText&&String.IsNullOrWhiteSpace(((XText)n).Value))))throw new InvalidDataException("Unknown or malformed agenda XML fields.");
        }
        private static int Number(XElement element,string key){int result;if(!Int32.TryParse((string)element.Attribute(key),NumberStyles.None,CultureInfo.InvariantCulture,out result))throw new InvalidDataException("Invalid agenda integer.");return result;}
        internal static ChapterAgendaState Decode(string xml)
        {
            if(String.IsNullOrEmpty(xml)||xml.Length>300000)throw new InvalidDataException("Missing or excessive agenda XML.");
            try
            {
                var settings=new XmlReaderSettings {DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null,MaxCharactersInDocument=300000,IgnoreComments=false,IgnoreProcessingInstructions=false};XDocument doc;using(var reader=XmlReader.Create(new StringReader(xml),settings))doc=XDocument.Load(reader,LoadOptions.None);
                if(doc.Nodes().Any(n=>!(n is XElement)))throw new InvalidDataException("Unexpected agenda document nodes.");var root=doc.Root;if(root==null)throw new InvalidDataException("Missing agenda root.");Shape(root,"agenda","version");if((string)root.Attribute("version")!="2"||root.Elements().Count()!=2)throw new InvalidDataException("Unsupported agenda schema.");var chapters=root.Element(Ns+"chapters");var baseline=root.Element(Ns+"baseline");if(chapters==null||baseline==null)throw new InvalidDataException("Missing agenda collections.");Shape(chapters,"chapters");Shape(baseline,"baseline");
                var state=new ChapterAgendaState();foreach(var item in chapters.Elements()){Shape(item,"chapter","slideId","slideIndex","level","scope","title");if(item.HasElements||state.Chapters.Count>=200)throw new InvalidDataException("Invalid chapter collection.");state.Chapters.Add(new ChapterRecord {SlideId=Number(item,"slideId"),SlideIndex=Number(item,"slideIndex"),Level=Number(item,"level"),Scope=Number(item,"scope"),Title=ChapterAgendaPlanner.Caption((string)item.Attribute("title"))});}
                foreach(var item in baseline.Elements()){Shape(item,"caption","slideId","title");int id=Number(item,"slideId");if(item.HasElements||state.Baseline.Count>=200||state.Baseline.ContainsKey(id))throw new InvalidDataException("Invalid caption baseline.");state.Baseline.Add(id,ChapterAgendaPlanner.Caption((string)item.Attribute("title")));}State(state);return state;
            }
            catch(XmlException ex){throw new InvalidDataException("Malformed agenda XML.",ex);}
        }
    }
}

