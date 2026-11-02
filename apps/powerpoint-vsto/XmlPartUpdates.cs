using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using Office=Microsoft.Office.Core;
using PowerPoint=Microsoft.Office.Interop.PowerPoint;
namespace PowerPointIntelligence.AddIn
{
    // LoadXML initializes a new part. Updating a populated part uses a verified replacement.
    internal static class XmlPartUpdates
    {
        internal static void Validate(string xml)
        {if(String.IsNullOrEmpty(xml)||xml.Length>1000000)throw new InvalidDataException("XML update is missing or exceeds one million characters.");using(var reader=XmlReader.Create(new StringReader(xml),new XmlReaderSettings {DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null,MaxCharactersInDocument=1000000})){while(reader.Read()){} }}
        private static Office.CustomXMLPart Add(PowerPoint.CustomerData owner,string xml)
        {var p=owner.Add();try{if(!p.LoadXML(xml))throw new InvalidDataException("Replacement XML rejected.");Validate(p.XML);if(!Equivalent(p.XML,xml))throw new InvalidDataException("Replacement XML verification failed.");return p;}catch{owner.Delete(p.Id);throw;}}
        internal static bool SamePartId(string a,string b)
        {Guid x,y;return Guid.TryParse(a,out x)&&Guid.TryParse(b,out y)?x==y:String.Equals(a,b,StringComparison.Ordinal);}
        private static XmlNode[] Content(XmlNode node)
        {
            int depth=0;bool preserve=false,spaceResolved=false;for(var element=node as XmlElement;element!=null;element=element.ParentNode as XmlElement){if(++depth>128)throw new InvalidDataException("XML comparison nesting exceeds 128 levels.");var space=element.GetAttributeNode("space","http://www.w3.org/XML/1998/namespace");if(!spaceResolved&&space!=null){preserve=space.Value=="preserve";spaceResolved=true;}}
            var children=node.ChildNodes.Cast<XmlNode>().ToArray();bool elementContent=children.Any(n=>n is XmlElement)&&!children.Any(n=>n.NodeType==XmlNodeType.Text||n.NodeType==XmlNodeType.CDATA||n.NodeType==XmlNodeType.SignificantWhitespace);
            return children.Where(n=>n.NodeType!=XmlNodeType.XmlDeclaration&&!(n.NodeType==XmlNodeType.Whitespace&&elementContent&&!preserve)).ToArray();
        }
        private static string AttributeKey(XmlAttribute a){return "{"+a.NamespaceURI+"}"+a.LocalName;}
        internal static bool Equivalent(string a,string b)
        {
            Validate(a);Validate(b);var x=new XmlDocument {XmlResolver=null,PreserveWhitespace=true};var y=new XmlDocument {XmlResolver=null,PreserveWhitespace=true};x.LoadXml(a);y.LoadXml(b);var pending=new Stack<Tuple<XmlNode,XmlNode>>();pending.Push(Tuple.Create((XmlNode)x,(XmlNode)y));
            while(pending.Count>0){var pair=pending.Pop();var left=pair.Item1;var right=pair.Item2;bool lt=left.NodeType==XmlNodeType.Text||left.NodeType==XmlNodeType.CDATA||left.NodeType==XmlNodeType.Whitespace||left.NodeType==XmlNodeType.SignificantWhitespace,rt=right.NodeType==XmlNodeType.Text||right.NodeType==XmlNodeType.CDATA||right.NodeType==XmlNodeType.Whitespace||right.NodeType==XmlNodeType.SignificantWhitespace;if(lt||rt){if(!lt||!rt||left.Value!=right.Value)return false;continue;}
                if(left.NodeType!=right.NodeType||left.LocalName!=right.LocalName||left.NamespaceURI!=right.NamespaceURI||left.Value!=right.Value)return false;
                if(left is XmlElement){var la=left.Attributes.Cast<XmlAttribute>().Where(v=>v.NamespaceURI!="http://www.w3.org/2000/xmlns/").OrderBy(AttributeKey,StringComparer.Ordinal).ToArray();var ra=right.Attributes.Cast<XmlAttribute>().Where(v=>v.NamespaceURI!="http://www.w3.org/2000/xmlns/").OrderBy(AttributeKey,StringComparer.Ordinal).ToArray();if(la.Length!=ra.Length)return false;for(int i=0;i<la.Length;i++)if(AttributeKey(la[i])!=AttributeKey(ra[i])||la[i].Value!=ra[i].Value)return false;}
                var lc=Content(left);var rc=Content(right);if(lc.Length!=rc.Length)return false;for(int i=lc.Length-1;i>=0;i--)pending.Push(Tuple.Create(lc[i],rc[i]));
            }return true;
        }
        private static bool Has(PowerPoint.Shape shape,string id){int count=0;foreach(Office.CustomXMLPart p in shape.CustomerData){if(++count>64)throw new InvalidDataException("Shape XML ownership exceeds 64 parts.");if(SamePartId(p.Id,id))return true;}return false;}
        internal static T[] CanonicalParts<T>(IEnumerable<T> observed,Func<T,string> identity,Func<string,T> resolve,Func<T,string> xml,string schema)
        {
            var result=new List<T>();var ids=new HashSet<string>(StringComparer.OrdinalIgnoreCase);int scanned=0;
            foreach(var item in observed){
                if(++scanned>64)throw new InvalidDataException("Shape XML ownership exceeds 64 parts.");
                string id=identity(item);Guid parsed;if(!Guid.TryParse(id,out parsed))throw new InvalidDataException("Shape metadata part identity is invalid.");
                if(!ids.Add(parsed.ToString("D")))continue;
                var current=resolve(id);if(RootNamespace(xml(current))==schema)result.Add(current);
            }
            return result.ToArray();
        }
        internal static Office.CustomXMLPart[] AttachedParts(PowerPoint.Shape shape,string schema)
        {
            // Resolve owner membership by GUID and use serialized XML rather than cached native namespace state.
            return CanonicalParts(shape.CustomerData.Cast<Office.CustomXMLPart>(),part=>part.Id,id=>shape.CustomerData[id],part=>part.XML,schema);
        }
        internal static string RootNamespace(string xml)
        {Validate(xml);using(var reader=XmlReader.Create(new StringReader(xml),new XmlReaderSettings {DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null,MaxCharactersInDocument=1000000})){reader.MoveToContent();return reader.NamespaceURI;}}
        internal static T ReplaceAttachment<T>(Action detach,Func<T> create,Action<T> verify,Action<T> cleanup,Action restore) where T:class
        {
            T added=null;try{detach();added=create();verify(added);return added;}
            catch(Exception original){var failures=new List<Exception>();if(added!=null)try{cleanup(added);}catch(Exception e){failures.Add(e);}try{restore();}catch(Exception e){failures.Add(e);}if(failures.Count>0){failures.Insert(0,original);throw new InvalidOperationException("Shape XML replacement failed and restoration was incomplete. Use Undo.",new AggregateException(failures));}throw;}
        }
        private static Office.CustomXMLPart VerifyAttachment(PowerPoint.Shape shape,string ns,string xml,string expectedId=null,string phase="preflight")
        {
            var parts=AttachedParts(shape,ns);int inspected=shape.CustomerData.Count;bool idMatches=parts.Length==1&&(expectedId==null||SamePartId(parts[0].Id,expectedId)),payloadMatches=parts.Length==1&&Equivalent(parts[0].XML,xml);if(parts.Length!=1||!idMatches||!payloadMatches)throw new InvalidDataException("Replacement shape XML attachment could not be verified: phase="+phase+"; namespaceMatches="+parts.Length+"; inspected="+inspected+"; idMatches="+idMatches+"; payloadMatches="+payloadMatches+". No XML content was logged.");return parts[0];
        }
        internal static Office.CustomXMLPart Shape(PowerPoint.Shape shape,Office.CustomXMLPart old,string xml)
        {
            Validate(xml);string prior=old.XML,id=old.Id,ns=RootNamespace(prior);if(RootNamespace(xml)!=ns)throw new InvalidDataException("Shape XML replacement cannot change its schema namespace.");VerifyAttachment(shape,ns,prior,id);if(Equivalent(prior,xml))return old;
            var slide=shape.Parent as PowerPoint.Slide;if(slide==null)throw new InvalidOperationException("XML update requires an ungrouped managed object.");var deck=(PowerPoint.Presentation)slide.Parent;var owners=new List<PowerPoint.Shape>();int visited=0;
            Action<PowerPoint.Shape,int> visit=null;visit=(candidate,depth)=>{if(++visited>10000||depth>8)throw new InvalidOperationException("XML ownership inspection exceeds its safe bound.");if(Has(candidate,id))owners.Add(candidate);if(candidate.Type==Office.MsoShapeType.msoGroup)foreach(PowerPoint.Shape child in candidate.GroupItems)visit(child,depth+1);};
            foreach(PowerPoint.Slide s in deck.Slides)foreach(PowerPoint.Shape candidate in s.Shapes)visit(candidate,0);
            Func<PowerPoint.Shape,bool> current=owner=>owner.Id==shape.Id&&owner.Parent is PowerPoint.Slide&&((PowerPoint.Slide)owner.Parent).SlideID==slide.SlideID;
            Action<PowerPoint.Shape> restoreOwner=owner=>{var matches=AttachedParts(owner,ns);if(matches.Length==0)Add(owner.CustomerData,prior);else if(matches.Length!=1||!Equivalent(matches[0].XML,prior))throw new InvalidDataException("Shared XML owner changed unexpectedly.");VerifyAttachment(owner,ns,prior,null,"restore-owner");};
            Action restoreAll=()=>{var failures=new List<Exception>();foreach(var owner in owners)try{restoreOwner(owner);}catch(Exception error){failures.Add(error);}if(failures.Count>0)throw new AggregateException(failures);};Office.CustomXMLPart verified=null;
            // Detach first: verification must describe the owner's final attachment, never just an orphan RCW.
            ReplaceAttachment(()=>shape.CustomerData.Delete(id),()=>Add(shape.CustomerData,xml),added=>{VerifyAttachment(shape,ns,xml,added.Id,"created");foreach(var owner in owners.Where(o=>!current(o)))restoreOwner(owner);verified=VerifyAttachment(shape,ns,xml,added.Id,"commit");},added=>{if(Has(shape,added.Id))shape.CustomerData.Delete(added.Id);},restoreAll);
            return verified;
        }
        internal static Office.CustomXMLPart Deck(PowerPoint.Presentation deck,Office.CustomXMLPart old,string xml)
        {
            Validate(xml);string prior=old.XML,ns=RootNamespace(prior),id=old.Id;
            if(RootNamespace(xml)!=ns)throw new InvalidDataException("Registry replacement cannot change its namespace.");
            Func<Office.CustomXMLPart[]> attached=()=>{var selected=deck.CustomXMLParts.SelectByNamespace(ns);return Enumerable.Range(1,selected.Count).Select(index=>selected[index]).Where(part=>RootNamespace(part.XML)==ns).ToArray();};
            var before=attached();bool identityMatches=before.Length==1&&SamePartId(before[0].Id,id),payloadMatches=before.Length==1&&Equivalent(before[0].XML,prior);
            if(before.Length!=1||!identityMatches||!payloadMatches)throw new InvalidDataException("Registry changed before replacement: matches="+before.Length+"; identityMatches="+identityMatches+"; payloadMatches="+payloadMatches+". Nothing was removed; no XML content logged.");
            if(Equivalent(prior,xml))return before[0];
            // Do not add a new same-namespace stream while retaining an old stream RCW for deletion.
            return ReplaceAttachment(()=>before[0].Delete(),()=>deck.CustomXMLParts.Add(xml),added=>{
                var current=attached();if(current.Length!=1||!SamePartId(current[0].Id,added.Id)||!Equivalent(current[0].XML,xml))throw new InvalidDataException("Replacement registry attachment could not be verified.");
            },added=>{var live=deck.CustomXMLParts.Cast<Office.CustomXMLPart>().SingleOrDefault(p=>SamePartId(p.Id,added.Id));if(live!=null)live.Delete();},()=>{
                var current=attached();if(current.Length==0){var restored=deck.CustomXMLParts.Add(prior);if(!Equivalent(restored.XML,prior))throw new InvalidDataException("Restored registry verification failed.");}
                else if(current.Length!=1||!Equivalent(current[0].XML,prior))throw new InvalidDataException("Registry restoration was ambiguous; use Undo.");
            });
        }
    }
}
