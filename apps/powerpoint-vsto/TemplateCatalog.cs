using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
namespace PowerPointIntelligence.AddIn
{
    internal sealed class LibraryDrafts
    {
        private readonly Dictionary<string,Dictionary<string,string>> saved=new Dictionary<string,Dictionary<string,string>>(StringComparer.Ordinal);
        internal Dictionary<string,string> Load(SlideTemplate template)
        {
            var result=template.Build().Where(n=>n.Kind=="text").ToDictionary(n=>n.Key,n=>n.Text);
            Dictionary<string,string> values;if(saved.TryGetValue(template.Id,out values))foreach(var pair in values)result[pair.Key]=pair.Value;
            return result;
        }
        internal void Save(SlideTemplate template,IDictionary<string,string> values)
        {
            var keys=new HashSet<string>(template.Build().Where(n=>n.Kind=="text").Select(n=>n.Key));
            if(values.Any(p=>!keys.Contains(p.Key)||p.Value==null||p.Value.Length>2000))throw new System.IO.InvalidDataException("Use known fields and no more than 2000 characters per field.");
            saved[template.Id]=new Dictionary<string,string>(values);
        }
    }
    internal static class TemplateTheme
    {
        // Office MsoThemeColorIndex: Text1=13, Background1=14, Accent1..6=5..10.
        internal static Tuple<int,float> Binding(string value,TemplatePalette palette,bool text)
        {
            int slot=Array.FindIndex(palette.Colors,c=>String.Equals(c,value,StringComparison.OrdinalIgnoreCase));
            if(slot==9||String.Equals(value,"#FFFFFF",StringComparison.OrdinalIgnoreCase))return Tuple.Create(14,0f);
            if(text)return Tuple.Create(13,0f);
            if(slot<0)return Tuple.Create(13,0f);
            return Tuple.Create(5+slot%6,slot>=6?.55f:0f);
        }
    }
    internal static class TemplateAlignment
    {
        internal static readonly string[] Horizontal={"Left","Center","Right"};
        internal static readonly string[] Vertical={"Top","Middle","Bottom"};
        internal static void Validate(string horizontal,string vertical)
        {if(!Horizontal.Contains(horizontal)||!Vertical.Contains(vertical))throw new System.IO.InvalidDataException("Choose a supported text alignment.");}
        internal static List<TemplateNode> Prepare(List<TemplateNode> nodes)
        {
            var labels=nodes.Where(n=>n.Kind=="text").ToArray();
            foreach(var box in nodes.Where(n=>n.Kind!="text"&&n.Kind!="line"))
            {
                var inside=labels.Where(t=>t.Stage==box.Stage&&t.X+t.W/2>=box.X&&t.X+t.W/2<=box.X+box.W&&t.Y+t.H/2>=box.Y&&t.Y+t.H/2<=box.Y+box.H).ToArray();
                if(inside.Length!=1)continue;
                var text=inside[0];float insetX=box.Kind=="chevron"?box.W*.23f:box.Kind=="diamond"?box.W*.22f:box.Kind=="polygon"?box.W*.23f:8;
                float insetY=box.Kind=="ellipse"||box.Kind=="diamond"?box.H*.15f:4;
                text.X=box.X+insetX;text.Y=box.Y+insetY;text.W=box.W-2*insetX;text.H=box.H-2*insetY;
            }
            return nodes;
        }
    }
    internal sealed class TemplatePalette
    {
        internal string Id,Name;
        internal string[] Colors;
        public override string ToString(){return Name;}
    }
    internal static class TemplatePalettes
    {
        internal const string DefaultId="01";
        internal static readonly IReadOnlyList<TemplatePalette> Items=Array.AsReadOnly(new[]{
            new TemplatePalette {Id="01",Name="Executive Slate",Colors=new[]{"#34495E","#536779","#708494","#90A1AE","#B1BEC7","#D1D9DF","#E9EDF0","#718B91","#A6B5B0","#F5F7F8"}},
            new TemplatePalette {Id="02",Name="Sage & Stone",Colors=new[]{"#52665B","#6E8272","#8FA391","#AFC0AB","#CDD8C8","#E5EAE0","#9B9A8C","#BEB8A8","#DED8C9","#F6F5F0"}},
            new TemplatePalette {Id="03",Name="Nordic Mist",Colors=new[]{"#4B6571","#688390","#86A0AB","#A8BBC2","#CAD6D9","#E5EBEC","#8D9F9D","#B3C0BB","#D5DEDA","#F5F8F7"}},
            new TemplatePalette {Id="04",Name="Dusty Rose",Colors=new[]{"#755C62","#96767C","#B39499","#CDB0B3","#E1CBCD","#F1E3E3","#9C8A88","#B9A9A3","#DDD1C8","#F9F5F2"}},
            new TemplatePalette {Id="05",Name="Muted Terracotta",Colors=new[]{"#795C50","#987466","#B58D7D","#CCA99A","#DEC4B8","#EFE1D9","#9C9484","#BBB2A1","#DCD5C7","#F8F4EE"}},
            new TemplatePalette {Id="06",Name="Coastal Blue",Colors=new[]{"#3F6274","#5C7D8C","#7897A3","#9CB3BA","#C0CFD1","#E0E8E8","#7D9C98","#A8BFBA","#D1DDD7","#F4F8F7"}},
            new TemplatePalette {Id="07",Name="Lavender Haze",Colors=new[]{"#625E78","#807B94","#9E99B0","#BBB5C9","#D5D0DE","#EAE7EF","#938CA1","#B2A9B9","#D8D1D7","#F8F6FA"}},
            new TemplatePalette {Id="08",Name="Warm Greige",Colors=new[]{"#635D57","#827A72","#A1998F","#BCB4AA","#D6CEC3","#EBE5DC","#8F8D84","#B1AFA4","#D9D6CD","#F8F6F1"}},
            new TemplatePalette {Id="09",Name="Olive & Linen",Colors=new[]{"#5D6650","#798267","#98A083","#B5BA9F","#D3D6BF","#E9E9D8","#938C70","#B5AC8E","#DAD3B9","#F8F7EE"}},
            new TemplatePalette {Id="10",Name="Powder Blue",Colors=new[]{"#526C86","#6D89A0","#8CA6BA","#ADC2D0","#CDDAE2","#E8EFF3","#8396A6","#A9B7C0","#D2DCE0","#F7F9FB"}},
            new TemplatePalette {Id="11",Name="Muted Jewel",Colors=new[]{"#4C6267","#657D7F","#849898","#A5B5B2","#CAD3CD","#E6EAE4","#81788A","#A99CAA","#D4C9D1","#F6F4F6"}},
            new TemplatePalette {Id="12",Name="Sand & Sea",Colors=new[]{"#536D70","#72908D","#92AAA3","#B6C6BA","#D6DFD3","#EDF0E7","#A79883","#C4B6A0","#E0D6C4","#F9F7F1"}},
            new TemplatePalette {Id="13",Name="Blush & Taupe",Colors=new[]{"#6E6060","#8F7B79","#AE9894","#CBB5AF","#E2D0C9","#F1E5DF","#9D9189","#BEB0A3","#DFD3C5","#FAF6F2"}},
            new TemplatePalette {Id="14",Name="Moss & Clay",Colors=new[]{"#566454","#72826D","#92A08A","#B3C0AA","#D2DACB","#E9EEE3","#A18C7D","#BFA99A","#DDD0C3","#F7F5EF"}},
            new TemplatePalette {Id="15",Name="Quiet Burgundy",Colors=new[]{"#705761","#906E78","#AF8D96","#C9ADB4","#DFCCD0","#F0E4E7","#93858A","#B8A7AA","#D9CCCB","#F9F5F6"}},
            new TemplatePalette {Id="16",Name="Cloudy Teal",Colors=new[]{"#456B6D","#628688","#81A1A0","#A4BDB9","#C8D7D0","#E4ECE7","#839E9A","#AABDB4","#D5DFD5","#F5F8F5"}},
            new TemplatePalette {Id="17",Name="Champagne Neutral",Colors=new[]{"#6E6558","#8B8070","#AA9C87","#C7B7A0","#E0D3BC","#F1E9D8","#A6A093","#C3BFB2","#E1DDD2","#FAF8F2"}},
            new TemplatePalette {Id="18",Name="Smoky Plum",Colors=new[]{"#625A69","#807585","#9F93A2","#BDB0BD","#D7CBD5","#EEE5EC","#938D96","#B5ACB6","#DAD2D9","#F8F5F8"}},
            new TemplatePalette {Id="19",Name="Soft Autumn",Colors=new[]{"#746453","#92806A","#AE9A81","#C8B69D","#DFD0B9","#F0E7D8","#9C8D78","#B9AB96","#DED2C0","#FAF7F0"}},
            new TemplatePalette {Id="20",Name="Modern Minimal",Colors=new[]{"#48545A","#657277","#859195","#A7B0B1","#CBD1D0","#E7EAE8","#89938D","#B1B9AF","#D9DED7","#F8F9F7"}},
        });
        internal static TemplatePalette Get(string id=null){var palette=Items.SingleOrDefault(p=>p.Id==(id??DefaultId));if(palette==null)throw new System.IO.InvalidDataException("Unknown library palette.");return palette;}
        private static string Map(string color,TemplatePalette p)
        {
            if(color==null)return null;
            string[] accents={"#2563EB","#0F766E","#7C3AED","#C2410C","#0369A1","#A21CAF","#475569"};int index=Array.IndexOf(accents,color);
            if(index>=0)return p.Colors[new[]{0,1,2,7,3,4,8}[index]];
            switch(color){case "#0F172A":return p.Colors[0];case "#FFFFFF":return p.Colors[9];case "#F1F5F9":return p.Colors[6];case "#DBEAFE":return p.Colors[5];case "#E2E8F0":return p.Colors[4];case "#CBD5E1":return p.Colors[3];case "#94A3B8":return p.Colors[2];default:return color;}
        }
        private static double Luminance(string hex)
        {var values=new[]{1,3,5}.Select(i=>Convert.ToInt32(hex.Substring(i,2),16)/255.0).Select(v=>v<=.04045?v/12.92:Math.Pow((v+.055)/1.055,2.4)).ToArray();return values[0]*.2126+values[1]*.7152+values[2]*.0722;}
        internal static double Contrast(string a,string b){double x=Luminance(a),y=Luminance(b);return (Math.Max(x,y)+.05)/(Math.Min(x,y)+.05);}
        internal static string TextColor(string background,TemplatePalette palette)
        {if(Contrast(palette.Colors[0],background)>=4.5)return palette.Colors[0];if(Contrast(palette.Colors[9],background)>=4.5)return palette.Colors[9];return Contrast("#000000",background)>Contrast("#FFFFFF",background)?"#000000":"#FFFFFF";}
        internal static List<TemplateNode> Apply(List<TemplateNode> nodes,string paletteId=null)
        {
            var palette=Get(paletteId);
            foreach(var node in nodes){node.Fill=Map(node.Fill,palette);node.Stroke=Map(node.Stroke,palette);}
            for(int i=0;i<nodes.Count;i++)if(nodes[i].Kind=="text")
            {
                var text=nodes[i];string background=palette.Colors[9];
                for(int b=0;b<i;b++){var box=nodes[b];if(box.Kind!="text"&&box.Kind!="line"&&box.Stage==text.Stage&&text.X+text.W/2>=box.X&&text.X+text.W/2<=box.X+box.W&&text.Y+text.H/2>=box.Y&&text.Y+text.H/2<=box.Y+box.H)background=box.Fill;}
                text.Fill=TextColor(background,palette);
            }
            return nodes;
        }
    }
    internal sealed class TemplateNode
    {
        internal string Kind,Key,Text,Fill,Stroke;
        internal float X,Y,W,H,Font=16;
        internal bool Bold;
        internal int Stage;
        internal float[] Points;
    }
    internal sealed class SlideTemplate
    {
        internal string Id,Name,Category,Purpose;
        internal int Family,Variant;
        internal List<TemplateNode> Build(string paletteId=null) {return TemplateAlignment.Prepare(TemplatePalettes.Apply(TemplateCatalog.Build(this),paletteId));}
        public override string ToString(){return Name;}
    }
    // Original procedural designs. 26 families x 20 structural configurations.
    // Themes and user edits are not counted as additional templates.
    internal static class TemplateCatalog
    {
        internal static readonly string[] Categories={"Agendas","Timelines","Processes","Cycles","Pyramids","Funnels","Organization","Decisions","Matrices","SWOT","Comparisons","Roadmaps","Swimlanes","Customer journeys","Portfolios","Scorecards","Teams","Quotes and evidence","Opening slides","Section navigation","Tables","Architecture","Ecosystems","Cubes","Overlapping sets","Experiment plans"};
        private static readonly string[] Purposes={
            "Structure a meeting around topics, owners and outcomes", "Place events and milestones in time", "Explain ordered activities and handoffs", "Explain recurring activities and feedback", "Rank layers of a hierarchy", "Explain stage gates and narrowing choices", "Describe reporting relationships and responsibilities", "Branch from a decision into alternative actions", "Position items against two dimensions", "Separate strengths, weaknesses, opportunities and threats", "Compare alternatives against common criteria", "Coordinate releases across workstreams", "Assign work to lanes and phases", "Connect touchpoints to customer needs", "Present an initiative with supporting opportunities", "Track metrics, owners and planned actions", "Introduce roles and responsibilities", "Connect a statement with supporting evidence", "Introduce a topic and the questions it answers", "Navigate a presentation with section progress", "Compare structured records in a native editable table", "Explain system components and interfaces", "Place a central organization among its partners", "Explain segmented dimensions with editable isometric faces", "Show intersections among related sets", "Connect hypotheses, evidence, measures and next steps"};
        private static readonly string[] Names={
            "Decision meeting|Weekly operations|Board agenda|Strategy workshop|Project kickoff|Customer review|Design critique|Budget review|Risk committee|Quarterly planning|Retrospective|Research session|Partner summit|Product review|Town hall|Sales briefing|Learning session|Investment review|Delivery standup|Steering committee",
            "Product history|Launch milestones|Implementation calendar|Transformation stages|Customer adoption|Research history|Market entry|Acquisition milestones|Policy rollout|Migration window|Hiring plan|Funding stages|Contract cycle|Audit calendar|Release history|Expansion phases|Learning pathway|Approval calendar|Incident chronology|Partnership launch",
            "Order fulfillment|Recruitment pipeline|Expense approval|Customer onboarding|Product discovery|Design review|Service recovery|Procurement flow|Compliance review|Incident response|Lead qualification|Contract negotiation|Data publishing|Change request|Quality assurance|Content publishing|Partner onboarding|Research workflow|Investment screening|Knowledge transfer",
            "Continuous improvement|Planning cadence|Customer feedback|Learning cycle|Risk monitoring|Product iteration|Performance review|Service management|Security response|Research loop|Inventory review|Experiment cycle|Content renewal|Partner governance|Budget cycle|Quality control|Demand planning|Talent development|Audit followup|Market sensing",
            "Strategic priorities|Capability maturity|Customer needs|Evidence hierarchy|Operating principles|Skills ladder|Value hierarchy|Governance layers|Service tiers|Brand architecture|Data maturity|Risk controls|Learning objectives|Portfolio priorities|Decision authority|Quality levels|Support levels|Impact hierarchy|Security layers|Transformation ambition",
            "Sales qualification|Candidate selection|Innovation screening|Investment shortlist|Idea prioritization|Customer activation|Procurement shortlist|Grant selection|Partner qualification|Lead conversion|Research screening|Market selection|Solution evaluation|Content review|Quality gates|Risk escalation|Opportunity assessment|Tender process|Acquisition screening|Experiment selection",
            "Leadership structure|Project governance|Product organization|Regional leadership|Service ownership|Program roles|Delivery squads|Risk ownership|Research team|Support organization|Partner governance|Decision rights|Operating model|Crossfunctional team|Center of excellence|Transformation office|Sales territories|Incident command|Portfolio governance|Capability ownership",
            "Build or buy|Market entry choice|Risk treatment|Product investment|Service escalation|Vendor selection|Architecture choice|Release readiness|Hiring decision|Pricing choice|Customer routing|Funding allocation|Policy exception|Quality disposition|Partner approval|Capacity choice|Incident triage|Research selection|Expansion choice|Contract renewal",
            "Impact and effort|Risk and likelihood|Value and feasibility|Growth and share|Urgency and importance|Capability and demand|Cost and benefit|Maturity and readiness|Reach and confidence|Value and complexity|Dependency and risk|Fit and attractiveness|Influence and interest|Quality and speed|Size and opportunity|Adoption and impact|Control and exposure|Skills and criticality|Demand and capacity|Evidence and uncertainty",
            "Business strategy|Product direction|Market position|Service portfolio|Team capability|Customer segment|Regional expansion|Digital maturity|Partner ecosystem|Operating model|Brand position|Supply chain|Research capability|Investment thesis|Talent strategy|Sales channel|Technology estate|Project readiness|Competitive position|Sustainability plan",
            "Vendor evaluation|Product alternatives|Pricing packages|Operating models|Technology options|Market opportunities|Service tiers|Investment choices|Delivery approaches|Policy alternatives|Channel strategies|Talent options|Site selection|Partner options|Architecture options|Research methods|Customer segments|Portfolio choices|Capability paths|Contract alternatives",
            "Product releases|Transformation delivery|Platform migration|Market expansion|Capability development|Compliance program|Customer experience|Data platform|Talent development|Brand launch|Partner enablement|Service improvement|Security program|Research agenda|Sustainability program|Integration delivery|Infrastructure renewal|Sales enablement|Operating model|Digital services",
            "Service delivery|Customer onboarding|Procurement handoffs|Recruitment responsibilities|Incident coordination|Product launch teams|Sales operations|Approval ownership|Partner delivery|Research collaboration|Risk management|Finance workflow|Change management|Implementation roles|Customer support|Data stewardship|Quality management|Marketing delivery|Program governance|Release coordination",
            "New customer|Returning customer|Enterprise buyer|Digital subscriber|Service requester|Partner applicant|Job candidate|Product evaluator|Support customer|Loyalty member|Mobile shopper|Business traveler|New employee|Account renewal|Research participant|Public service user|Channel partner|Learning participant|Investor relations|Customer recovery",
            "Strategic initiatives|Product portfolio|Research portfolio|Investment thesis|Market opportunities|Service catalog|Transformation portfolio|Customer segments|Partner opportunities|Capability investments|Innovation agenda|Regional priorities|Technology portfolio|Growth opportunities|Risk initiatives|Talent programs|Brand initiatives|Operational priorities|Sustainability actions|Experience improvements",
            "Executive performance|Sales performance|Delivery health|Customer outcomes|Service performance|Financial health|Product performance|People performance|Risk indicators|Research progress|Partner outcomes|Market performance|Quality indicators|Security posture|Capacity planning|Change readiness|Portfolio outcomes|Learning outcomes|Operational health|Sustainability outcomes",
            "Leadership introduction|Project team|Delivery roles|Research group|Customer account team|Partner team|Regional team|Product squad|Steering committee|Advisory board|Support team|Transformation office|Sales team|Service team|Hiring panel|Program sponsors|Design team|Data team|Operations team|Expert network",
            "Customer voice|Research finding|Leadership statement|Market observation|Employee voice|Partner perspective|Design principle|Investment argument|Project lesson|Service promise|Strategic belief|Quality principle|Culture statement|Change narrative|Product insight|Policy rationale|Evidence summary|User insight|Risk insight|Opportunity statement",
            "Strategy presentation|Project proposal|Quarterly review|Product launch|Research readout|Customer briefing|Investment proposal|Transformation plan|Team introduction|Partner proposal|Market analysis|Operating plan|Innovation pitch|Service review|Budget proposal|Risk review|Learning workshop|Leadership update|Program launch|Annual outlook",
            "Strategy sections|Project chapters|Research sections|Review agenda|Proposal structure|Planning chapters|Product narrative|Customer story|Investment narrative|Transformation chapters|Market analysis sections|Operating plan sections|Service review chapters|Learning modules|Risk review sections|Partner story|Annual report sections|Workshop stages|Decision narrative|Program stages",
            "Action register|Risk register|Decision log|Requirements table|Assumptions register|Issue tracker|Stakeholder register|Resource plan|Budget worksheet|Release checklist|Research inventory|Supplier register|Benefits register|Dependency register|Compliance checklist|Training plan|Communication plan|Service catalog|Experiment register|Responsibility register",
            "Application landscape|Data pipeline|Service architecture|Integration map|Security architecture|Analytics platform|Channel architecture|Cloud components|Operating platform|Customer systems|Knowledge platform|Delivery platform|Identity architecture|Research platform|Product platform|Enterprise layers|Partner integration|Information flow|API landscape|Infrastructure layers",
            "Partner landscape|Customer ecosystem|Supplier network|Innovation ecosystem|Platform ecosystem|Research network|Channel ecosystem|Service network|Stakeholder network|Community ecosystem|Talent network|Investment ecosystem|Regional ecosystem|Knowledge network|Delivery ecosystem|Technology partners|Public stakeholders|Value network|Market ecosystem|Strategic alliances",
            "Capability dimensions|Market segmentation|Portfolio dimensions|Risk dimensions|Product dimensions|Customer dimensions|Operating dimensions|Service dimensions|Value dimensions|Maturity dimensions|Quality dimensions|Growth dimensions|Technology dimensions|Research dimensions|Talent dimensions|Security dimensions|Investment dimensions|Channel dimensions|Governance dimensions|Sustainability dimensions",
            "Shared capabilities|Customer overlap|Partner overlap|Product intersections|Market intersections|Skills overlap|Service overlap|Responsibility overlap|Data intersections|Technology overlap|Research intersections|Stakeholder overlap|Value intersections|Channel overlap|Regional intersections|Risk intersections|Knowledge overlap|Resource intersections|Capability convergence|Strategic alignment",
            "Product hypothesis|Customer interview|Pricing experiment|Channel experiment|Service pilot|Process experiment|Market test|Message test|Usability study|Feature experiment|Partner pilot|Operational trial|Policy pilot|Training pilot|Research study|Quality experiment|Demand test|Retention experiment|Adoption pilot|Value proposition test"};
        internal static readonly IReadOnlyList<SlideTemplate> Items=Create();
        private static IReadOnlyList<SlideTemplate> Create()
        {
            var result=new List<SlideTemplate>();
            for(int f=0;f<Categories.Length;f++){var names=Names[f].Split('|');if(names.Length!=20)throw new InvalidOperationException("Template titles must have 20 entries.");for(int v=0;v<20;v++)result.Add(new SlideTemplate {Id="ppi-library-v1-"+f.ToString("D2")+"-"+v.ToString("D2"),Name=names[v],Category=Categories[f],Purpose=Purposes[f],Family=f,Variant=v});}
            return result.AsReadOnly();
        }
        private static readonly Lazy<Dictionary<string,string>> SearchText=new Lazy<Dictionary<string,string>>(()=>Items.ToDictionary(t=>t.Id,t=>t.Name+" "+t.Category+" "+t.Purpose+" "+t.Id+" "+String.Join(" ",Build(t).Where(n=>n.Kind=="text").Select(n=>n.Text))));
        internal static IEnumerable<SlideTemplate> Search(string category,string query)
        {
            var terms=System.Text.RegularExpressions.Regex.Split((query??"").Trim(),@"\s+").Where(t=>t.Length>0).ToArray();
            return Items.Where(t=>(String.IsNullOrEmpty(category)||t.Category==category)&&terms.All(term=>SearchText.Value[t.Id].IndexOf(term,StringComparison.OrdinalIgnoreCase)>=0));
        }
        private static readonly string[] Palette={"#2563EB","#0F766E","#7C3AED","#C2410C","#0369A1","#A21CAF","#475569"};
        internal static List<TemplateNode> Build(SlideTemplate template)
        {
            var s=new Composer();int f=template.Family,v=template.Variant,n=3+v%5,mode=v/5;
            s.Text("title",template.Name,44,30,872,48,30,true,0);
            s.Text("subtitle","[Add the main message for this slide]",44,80,872,30,16,false,0);
            float top=136,bottom=474,width=872,height=338;
            switch(f)
            {
                case 0: // Agenda: owner rows, split columns, numbered rail, timed workshop.
                    for(int i=0;i<n;i++){float y=top+i*height/n,x=mode==1?44+(i%2)*444:44,w=mode==1?420:872;if(mode==1)y=top+(i/2)*(height/(float)Math.Ceiling(n/2.0));s.Shape("rect",x,y,w,mode==1?height/(float)Math.Ceiling(n/2.0)-12:height/n-8,"#F1F5F9",i+1);s.Text("topic"+i,"[Topic "+(i+1)+"]",x+16,y+8,w*(mode==3?.45f:.65f),26,18,true,i+1);if(mode>1)s.Text("owner"+i,mode==3?"[Time / owner / outcome]":"[Owner and expected outcome]",x+w*.48f,y+8,w*.49f,30,13,false,i+1);}
                    break;
                case 1:
                    s.Line(70,mode%2==0?290:top,mode%2==0?880:70,mode%2==0?290:bottom,"#CBD5E1");
                    if(mode==2)s.Text("era","[Time horizon / decision window]",260,430,500,32,18,true,0);
                    for(int i=0;i<n;i++){float x=mode%2==0?70+i*800f/(n-1):100+(mode==3?i%2*390:0),y=mode%2==0?290:top+i*300f/n;s.Shape("ellipse",x-7,y-7,14,14,Palette[i%7],i+1);s.Text("event"+i,"[Milestone "+(i+1)+"]",mode%2==0?Math.Min(x-25,785):x+20,mode%2==0?(i%2==0?195:325):y,mode%2==0?130:330,42,16,true,i+1);s.Text("date"+i,"[Date / outcome]",mode%2==0?Math.Min(x-25,785):x+20,mode%2==0?(i%2==0?235:368):y+30,mode%2==0?130:330,32,12,false,i+1);}
                    break;
                case 2:
                    for(int i=0;i<n;i++){bool vertical=mode==1;float x=vertical?160:44+i*width/n,y=vertical?top+i*height/n:(mode==2?200+(i%2)*130:mode==3?160+i*22:250),w=vertical?640:width/n-6,h=vertical?height/n-8:86;s.Shape("chevron",x,y,w,h,Palette[i%7],i+1);s.Text("step"+i,"[Step "+(i+1)+"]",x+w*.23f,y+14,w*.55f,h-24,16,true,i+1,"#FFFFFF");}
                    break;
                case 3: case 22:
                    if(f==3)for(int i=0;i<n;i++){double aa=2*Math.PI*i/n-Math.PI/2+mode*.22,bb=aa+2*Math.PI/n;float rx=mode==3?310:300,ry=mode==2?120:140;s.Line(480+(float)Math.Cos(aa)*rx,292+(float)Math.Sin(aa)*ry,480+(float)Math.Cos(bb)*rx,292+(float)Math.Sin(bb)*ry,"#CBD5E1");}
                    if(f==22)for(int i=0;i<n;i++){double a=2*Math.PI*i/n-Math.PI/2+mode*.22;s.Line(480,292,480+(float)Math.Cos(a)*(mode==3?310:300),292+(float)Math.Sin(a)*(mode==2?120:140),"#CBD5E1");}
                    s.Shape("ellipse",385,245,190,95,"#E2E8F0",0);s.Text("center","[Core purpose]",395,267,170,48,19,true,0);
                    for(int i=0;i<n;i++){double a=2*Math.PI*i/n-Math.PI/2+mode*.22;float x=480+(float)Math.Cos(a)*(mode==3?310:300),y=292+(float)Math.Sin(a)*(mode==2?120:140);s.Shape(f==3?"ellipse":"round",x-80,y-35,160,70,Palette[i%7],i+1);s.Text("node"+i,f==3?"[Cycle stage "+(i+1)+"]":"[Partner / role "+(i+1)+"]",x-70,y-22,140,48,15,true,i+1,"#FFFFFF");}
                    break;
                case 4: case 5:
                    for(int i=0;i<n;i++){float y=top+i*height/n,h=height/n-6,w=f==4?150+i*(620f/n):790-i*(580f/n),x=mode%2==0?480-w/2:60;if(mode%2==1)w=Math.Min(w,460);float shift=Math.Min((mode/2)*120,w*.22f);if(mode/2>0)x=Math.Min(x,60);s.Polygon(new[]{x+shift,y,x+w-shift,y,x+w,y+h,x,y+h},Palette[i%7],i+1);s.Text("layer"+i,"["+(f==4?"Priority ":"Gate ")+(i+1)+"]",x+Math.Min(shift+12,w/4),y+4,w-Math.Min(2*shift+24,w/2),h-6,15,true,i+1,"#FFFFFF");if(mode%2==1)s.Text("detail"+i,"[Owner / evidence / action]",570,y+4,340,h-6,14,false,i+1);}
                    break;
                case 6: case 7:
                    s.Shape(f==7?"diamond":"round",370,top,220,58,"#0F172A",0);s.Text("root",f==7?"[Decision question]":"[Accountable leader]",390,top+10,180,40,16,true,0,"#FFFFFF");
                    if(mode==2){s.Shape("round",390,205,180,36,"#E2E8F0",0);s.Text("coordination","[Coordination hub]",400,211,160,26,13,true,0);}
                    for(int i=0;i<n;i++){int cc=mode==0?n:mode==1?2:mode==2?3:4;float ww=840f/cc-12,xx=60+i%cc*840f/cc,yy=(mode==2?275:235)+i/cc*((mode==2?170f:210f)/(float)Math.Ceiling(n/(double)cc));s.Line(480,mode==2?241:194,xx+ww/2,yy,"#CBD5E1");}
                    for(int i=0;i<n;i++){int columns=mode==0?n:mode==1?2:mode==2?3:4;int row=i/columns,col=i%columns;float w=840f/columns-12,x=60+col*840f/columns,y=(mode==2?275:235)+row*((mode==2?170f:210f)/(float)Math.Ceiling(n/(double)columns));s.Shape(mode==3?(f==7?"diamond":"ellipse"):"rect",x,y,w,60,Palette[i%7],i+1);s.Text("branch"+i,f==7?"[Option / condition "+(i+1)+"]":"[Role / team "+(i+1)+"]",x+9,y+9,w-18,44,14,true,i+1,"#FFFFFF");}
                    break;
                case 8:
                    int rows=mode+2;float cw=width/n,ch=height/rows;
                    for(int r=0;r<rows;r++)for(int c=0;c<n;c++){s.Shape("rect",44+c*cw,top+r*ch,cw-4,ch-4,r==0?"#DBEAFE":"#F1F5F9",r+1);s.Text("cell"+r+"-"+c,r==0?"[Dimension "+(c+1)+"]":"[Item / position]",50+c*cw,top+8+r*ch,cw-16,ch-16,14,r==0,r+1);}
                    break;
                case 9:
                    string[] swot={"Strengths","Weaknesses","Opportunities","Threats"};
                    for(int q=0;q<4;q++){
                        float x=mode==2?44:mode==1?44+q*222:44+q%2*(mode==3?500:444),y=mode==1?top:mode==2?top+q*84.5f:top+q/2*(mode==3?140:169);
                        float w=mode==1?210:mode==2?872:mode==3?(q%2==0?484:372):428,h=mode==1?338:mode==2?76.5f:mode==3?(q/2==0?128:186):157;
                        int stage=q+1;s.Shape("rect",x,y,w,h,q==mode?"#DBEAFE":"#F1F5F9",stage);
                        s.Text("quadrant"+q,swot[q],x+12,y+8,mode==2?160:w-24,26,18,true,stage);
                        for(int i=0;i<n;i++)s.Text("item"+q+"-"+i,"[Add "+swot[q].ToLowerInvariant()+" "+(i+1)+"]",mode==2?x+180+i*680f/n:x+12,mode==2?y+8:y+39+i*(h-44)/n,mode==2?680f/n-6:w-24,mode==2?60:(h-44)/n,12,false,stage);
                    }
                    break;
                case 10:
                    int options=2+mode;for(int c=0;c<options;c++){float w=width/options,x=44+c*w;s.Shape("rect",x,top,w-10,338,c==0?"#DBEAFE":"#F1F5F9",c+1);s.Text("option"+c,"[Option "+(c+1)+"]",x+12,top+10,w-34,30,20,true,c+1);for(int r=0;r<n;r++){s.Line(x+12,190+r*270f/n,x+w-24,190+r*270f/n,"#CBD5E1");s.Text("criterion"+c+"-"+r,"[Criterion "+(r+1)+": evidence]",x+12,196+r*270f/n,w-34,270f/n-6,13,false,c+1);}}
                    break;
                case 11: case 12:
                    int lanes=f==11?mode+2:n,phases=f==11?n:mode+3;float laneH=height/lanes;
                    for(int l=0;l<lanes;l++){float y=top+l*laneH;s.Text("lane"+l,"[Workstream "+(l+1)+"]",44,y+8,160,laneH-12,14,true,l+1);for(int c=0;c<phases;c++){float x=220+c*680f/phases;if(f==12||c==(l+v)%phases||c==(l+v+1)%phases){s.Shape(f==11?"round":"rect",x,y+8,680f/phases-8,laneH-16,Palette[l%7],l+1);s.Text("work"+l+"-"+c,f==11?"[Deliverable]":"[Owner / handoff]",x+5,y+15,680f/phases-18,laneH-30,12,true,l+1,"#FFFFFF");}else s.Line(x,y+laneH-1,x+680f/phases-6,y+laneH-1,"#CBD5E1");}}
                    break;
                case 13:
                    for(int i=0;i<n;i++){float x=44+i*width/n,w=width/n-8;s.Shape("round",x,top,w,58,Palette[i%7],i+1);s.Text("touch"+i,"[Touchpoint "+(i+1)+"]",x+6,top+10,w-12,40,14,true,i+1,"#FFFFFF");for(int r=0;r<mode+2;r++){float y=220+r*230f/(mode+2);s.Text("need"+i+"-"+r,new[]{"[Customer need]","[Friction]","[Opportunity]","[Owner / measure]","[Follow-up]"}[r],x+5,y,w-10,230f/(mode+2)-8,13,false,i+1);}}
                    break;
                case 14:
                    bool left=mode%2==0;float fx=left?44:560;s.Shape("rect",fx,top,350,338,"#0F172A",0);s.Text("featured","[Lead initiative]",fx+24,top+30,300,52,26,true,0,"#FFFFFF");s.Text("thesis","[State the opportunity and next decision]",fx+24,top+105,300,130,20,false,0,"#FFFFFF");for(int i=0;i<n;i++){float x=left?420:44,y=top+i*height/n;s.Shape(mode<2?"rect":"round",x,y,490,height/n-8,"#F1F5F9",i+1);s.Text("initiative"+i,"[Initiative "+(i+1)+" / owner / next action]",x+12,y+10,465,height/n-22,15,false,i+1);}
                    break;
                case 15: case 16:
                    int cols=mode+2,rr=(int)Math.Ceiling(n/(double)cols);for(int i=0;i<n;i++){
                        float w=width/cols-12,h=height/rr-12,x=44+i%cols*width/cols,y=top+i/cols*height/rr;
                        s.Shape(f==15?"rect":"round",x,y,w,h,"#F1F5F9",i+1);
                        if(f==16){
                            bool compact=h<160;float avatar=compact?Math.Min(42,h*.4f):62;
                            s.Shape("ellipse",x+12,y+12,avatar,avatar,Palette[i%7],i+1);
                            float tx=compact?x+avatar+24:x+14,ty=compact?y+8:y+88,tw=compact?w-avatar-38:w-28;
                            s.Text("name"+i,"[Name / role "+(i+1)+"]",tx,ty,tw,compact?26:42,compact?14:17,true,i+1);
                            float dy=compact?y+38:y+145;s.Text("value"+i,"[Responsibility / contact]",tx,dy,tw,y+h-dy-8,compact?12:15,false,i+1);
                        } else {
                            s.Text("name"+i,"[Metric "+(i+1)+"]",x+14,y+12,w-28,32,17,true,i+1);
                            s.Text("value"+i,"[Value / target / status]",x+14,y+h*.55f,w-28,h*.4f,15,false,i+1);
                        }
                    }
                    break;
                case 17:
                    float quoteW=mode==2?310:mode%2==0?460:872;s.Shape("rect",44,top,quoteW,mode%2==0?338:120,"#0F172A",0);s.Text("quote","[Insert the exact statement]",68,top+24,quoteW-48,mode%2==0?190:68,mode%2==0?30:24,true,0,"#FFFFFF");s.Text("attribution","[Speaker / source / date]",68,mode%2==0?410:top+90,quoteW-48,30,14,false,0,"#FFFFFF");for(int i=0;i<n;i++){float x=mode==2?390:mode%2==0?535:44+i*width/n,y=mode%2==0?top+i*height/n:mode==3?310:290,w=mode==2?520:mode%2==0?380:width/n-8,h=mode%2==0?height/n-8:mode==3?155:175;s.Text("evidence"+i,"[Evidence "+(i+1)+" and implication]",x,y,w,h,mode<2?16:14,false,i+1);}
                    break;
                case 18: case 19:
                    s.Shape("rect",mode%2==0?44:610,top,mode%2==0?290:306,338,"#0F172A",0);s.Text("promise",f==18?"[What the audience will learn]":"[Current section]",mode%2==0?68:634,top+40,240,190,28,true,0,"#FFFFFF");for(int i=0;i<n;i++){float x=mode%2==0?365:44,y=top+i*height/n,w=mode%2==0?550:535;if(f==19&&mode>=2)s.Shape("ellipse",x-24,y+height/n/2-8,16,16,Palette[i%7],i+1);s.Shape(f==19?"round":mode<2?"rect":"chevron",x,y,w,height/n-8,i==v%n?"#DBEAFE":"#F1F5F9",i+1);s.Text("chapter"+i,(i+1).ToString("D2")+"  ["+(f==18?"Question ":"Section ")+(i+1)+"]",x+14,y+10,w-28,height/n-22,f==18?18:16,true,i+1);}
                    break;
                case 20:
                    int tc=3+mode;for(int r=0;r<=n;r++)for(int c=0;c<tc;c++){float cw2=width/tc,rh=height/(n+1);s.Shape("rect",44+c*cw2,top+r*rh,cw2,rh,r==0?"#0F172A":r%2==0?"#F1F5F9":"#FFFFFF",r);s.Text("table"+r+"-"+c,r==0?"[Column "+(c+1)+"]":"[Enter value]",50+c*cw2,top+r*rh+5,cw2-12,rh-10,13,r==0,r,r==0?"#FFFFFF":"#0F172A");}
                    break;
                case 21:
                    int tiers=mode+2;for(int tier=0;tier<tiers;tier++){float y=top+tier*height/tiers;s.Shape("rect",44,y,872,height/tiers-8,"#F1F5F9",tier+1);s.Text("tier"+tier,"[Layer "+(tier+1)+"]",54,y+9,130,32,15,true,tier+1);for(int i=0;i<n;i++){float w=700f/n-8,x=200+i*700f/n;s.Shape("round",x,y+16,w,height/tiers-37,Palette[(i+tier)%7],tier+1);s.Text("component"+tier+"-"+i,"[Component]",x+6,y+20,w-12,height/tiers-43,12,true,tier+1,"#FFFFFF");}}
                    break;
                case 23:
                    for(int i=0;i<n;i++){float x=mode==0?160+i*58:mode==1?150+i%3*205:mode==2?180+i*85:190+i%2*340,y=mode==0?235-i*9:mode==1?190+i/3*100:mode==2?320-i*18:170+i/2*75,sz=mode==0?115:mode==1?90:mode==2?80:65;int stage=i+1;s.Polygon(new[]{x,y,x+sz*.6f,y-sz*.35f,x+sz*1.2f,y,x+sz*.6f,y+sz*.35f},Palette[i%7],stage);s.Polygon(new[]{x,y,x+sz*.6f,y+sz*.35f,x+sz*.6f,y+sz*1.1f,x,y+sz*.75f},"#DBEAFE",stage);s.Polygon(new[]{x+sz*.6f,y+sz*.35f,x+sz*1.2f,y,x+sz*1.2f,y+sz*.75f,x+sz*.6f,y+sz*1.1f},"#94A3B8",stage);s.Text("dimension"+i,"[Dimension "+(i+1)+"]",mode<2?70+i*120:60+i*120,435,118,38,13,false,stage);}
                    break;
                case 24:
                    for(int i=0;i<n;i++){double angle=2*Math.PI*i/n+mode*.3;float x=480+(float)Math.Cos(angle)*(70+mode*12),y=310+(float)Math.Sin(angle)*(65+mode*5),size=210+mode*10;s.Shape("ellipse",x-size/2,y-size/2,size,size,"#DBEAFE",i+1);}
                    for(int i=0;i<n;i++){double angle=2*Math.PI*i/n+mode*.3;float x=480+(float)Math.Cos(angle)*300,y=310+(float)Math.Sin(angle)*185;s.Text("set"+i,"[Set "+(i+1)+"]",x-50,y-13,100,26,15,true,i+1);}
                    s.Text("intersection","[Shared value]",405,297,150,32,16,true,n+1);
                    break;
                case 25:
                    for(int i=0;i<n;i++){float x=44+i*width/n,w=width/n-8;s.Text("phase"+i,"[Test "+(i+1)+"]",x+5,top,w-10,32,18,true,i+1);for(int r=0;r<mode+2;r++){float y=185+r*280f/(mode+2);s.Shape("rect",x,y,w,280f/(mode+2)-10,"#F1F5F9",i+1);s.Text("test"+i+"-"+r,new[]{"[Hypothesis]","[Evidence / method]","[Success measure]","[Next decision]","[Owner / timing]"}[r],x+8,y+10,w-16,280f/(mode+2)-28,14,false,i+1);}}
                    break;
            }
            try {Validate(s.Nodes);}catch(InvalidOperationException ex){throw new InvalidOperationException(template.Id+": "+ex.Message,ex);}return s.Nodes;
        }
        internal static void Validate(List<TemplateNode> nodes)
        {
            if(nodes.Count<3||nodes.Count>300)throw new InvalidOperationException("Template size outside supported bounds.");
            if(nodes.Where(x=>x.Kind=="text").Select(x=>x.Key).Distinct().Count()!=nodes.Count(x=>x.Kind=="text"))throw new InvalidOperationException("Duplicate editable field key.");
            foreach(var node in nodes){if(node.W<=0||node.H<=0||node.X<0||node.Y<0||node.X+node.W>960.1||node.Y+node.H>540.1||Single.IsNaN(node.X+node.Y+node.W+node.H)||Single.IsInfinity(node.X+node.Y+node.W+node.H))throw new InvalidOperationException("Template geometry outside canvas.");}
        }
        private sealed class Composer
        {
            internal readonly List<TemplateNode> Nodes=new List<TemplateNode>();
            internal void Text(string key,string text,float x,float y,float w,float h,float font,bool bold,int stage,string color="#0F172A") {Nodes.Add(new TemplateNode {Kind="text",Key=key,Text=text,X=x,Y=y,W=w,H=h,Font=font,Bold=bold,Fill=color,Stage=stage});}
            internal void Shape(string kind,float x,float y,float w,float h,string color,int stage){Nodes.Add(new TemplateNode {Kind=kind,X=x,Y=y,W=w,H=h,Fill=color,Stroke="#FFFFFF",Stage=stage});}
            internal void Line(float x,float y,float x2,float y2,string color){Nodes.Add(new TemplateNode {Kind="line",X=Math.Min(x,x2),Y=Math.Min(y,y2),W=Math.Max(1,Math.Abs(x2-x)),H=Math.Max(1,Math.Abs(y2-y)),Points=new[]{x,y,x2,y2},Stroke=color,Stage=0});}
            internal void Polygon(float[] points,string color,int stage){var xx=points.Where((_,i)=>i%2==0).ToArray();var yy=points.Where((_,i)=>i%2==1).ToArray();Nodes.Add(new TemplateNode {Kind="polygon",X=xx.Min(),Y=yy.Min(),W=xx.Max()-xx.Min(),H=yy.Max()-yy.Min(),Points=points,Fill=color,Stroke="#FFFFFF",Stage=stage});}
        }
    }
}



