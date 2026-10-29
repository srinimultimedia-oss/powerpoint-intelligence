using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
namespace PowerPointIntelligence.Library
{
    internal sealed partial class PiidxCatalog
    {
        static string MatchExpression(string query,LibrarySearchMode mode)
        {
            var terms=new List<string>();
            if(mode==LibrarySearchMode.Phrase)
            {
                var words=Regex.Matches(query,@"[\p{L}\p{N}\p{M}_]+").Cast<Match>().Select(m=>m.Value).ToArray();return words.Length==0?null:"\""+String.Join(" ",words)+"\"";
            }
            foreach(Match token in Regex.Matches(query,"\"([^\"]*)\"|(\\S+)"))
            {
                bool quoted=token.Groups[1].Success;string input=quoted?token.Groups[1].Value:token.Groups[2].Value;
                var words=Regex.Matches(input,@"[\p{L}\p{N}\p{M}_]+").Cast<Match>().Select(m=>m.Value).ToArray();if(words.Length==0)continue;
                string term="\""+String.Join(" ",words)+"\"";
                if(!quoted&&(mode==LibrarySearchMode.Prefix||input.EndsWith("*",StringComparison.Ordinal)))term+="*";
                terms.Add(term);if(terms.Count>64)throw new ArgumentException("Search has more than 64 terms.");
            }
            return terms.Count==0?null:String.Join(" AND ",terms);
        }
        internal LibrarySearchPage Search(LibrarySearchRequest request,CancellationToken cancellation)
        {
            if(request==null)throw new ArgumentNullException("request");cancellation.ThrowIfCancellationRequested();
            string query=request.Query??"";
            if(query.Length>4096||request.Limit<1||request.Limit>1000||request.Offset<0||request.Offset>10000000||!Enum.IsDefined(typeof(LibrarySearchMode),request.Mode))throw new ArgumentException("Search request exceeds supported bounds.");
            var parameters=new List<object>();string from=" FROM SLIDES s JOIN PRESENTATIONS p ON p.presentation_id=s.presentation_id JOIN LIBRARIES l ON l.library_id=p.library_id";
            var predicates=new List<string>{"p.index_status='indexed'","l.enabled=1"};bool ranked=false;
            if(!String.IsNullOrWhiteSpace(query))
            {
                if(request.Mode==LibrarySearchMode.ExactText){predicates.Add("(s.title=? OR s.text=?)");parameters.Add(query);parameters.Add(query);}
                else
                {
                    string match=MatchExpression(query,request.Mode);
                    if(match==null)return new LibrarySearchPage{Total=0,Slides=new List<LibrarySlideReference>()};
                    from+=" JOIN SLIDE_SEARCH ON SLIDE_SEARCH.slide_id=s.slide_id";predicates.Add("SLIDE_SEARCH MATCH ?");parameters.Add(match);ranked=true;
                }
            }
            if(!String.IsNullOrEmpty(request.LibraryId)){Id(request.LibraryId);predicates.Add("p.library_id=?");parameters.Add(request.LibraryId);}
            if(!String.IsNullOrEmpty(request.PresentationId)){Id(request.PresentationId);predicates.Add("p.presentation_id=?");parameters.Add(request.PresentationId);}
            if(request.Category!=null){predicates.Add("s.category=?");parameters.Add(request.Category.Replace('\\','/').Trim('/'));}
            if(!String.IsNullOrEmpty(request.Tag)){predicates.Add("EXISTS(SELECT 1 FROM SLIDE_TAGS t WHERE t.slide_id=s.slide_id AND t.normalized_tag=?)");parameters.Add(request.Tag.ToUpperInvariant());}
            if(!String.IsNullOrEmpty(request.Author)){predicates.Add("instr(upper(s.author),upper(?))>0");parameters.Add(request.Author);}
            if(!String.IsNullOrEmpty(request.Layout)){predicates.Add("(instr(upper(s.layout),upper(?))>0 OR instr(upper(s.master),upper(?))>0)");parameters.Add(request.Layout);parameters.Add(request.Layout);}
            string where=" WHERE "+String.Join(" AND ",predicates),rank=ranked?"bm25(SLIDE_SEARCH,0.0,5.0,1.0,2.0)":"0.0";
            using(var db=Connect(true))using(db.ObserveCancellation(cancellation))
            {
                db.Execute("BEGIN");try
                {
                    long total=(long)db.Query("SELECT count(*)"+from+where,parameters.ToArray()).Single()[0];cancellation.ThrowIfCancellationRequested();
                    var values=new List<object>(parameters){request.Limit,request.Offset};
                    string sql="SELECT s.slide_id,p.presentation_id,p.library_id,p.source_path,p.content_hash,s.content_fingerprint,s.native_slide_id,s.title,s.text,s.category,s.author,s.layout,s.master,s.position,s.width,s.height,p.modified_utc,"+rank+from+where+" ORDER BY "+(ranked?rank+",":"")+"p.source_path COLLATE NOCASE,s.position,s.slide_id LIMIT ? OFFSET ?";
                    var rows=new List<LibrarySlideReference>();
                    foreach(var r in db.Query(sql,values.ToArray()))
                    {
                        cancellation.ThrowIfCancellationRequested();rows.Add(new LibrarySlideReference{SlideId=(string)r[0],PresentationId=(string)r[1],LibraryId=(string)r[2],SourcePath=(string)r[3],SourceFingerprint=(string)r[4],SlideFingerprint=(string)r[5],NativeId=(string)r[6],Title=(string)r[7],Text=(string)r[8],Category=(string)r[9],Author=(string)r[10],Layout=(string)r[11],Master=(string)r[12],Position=checked((int)(long)r[13]),Width=Convert.ToDouble(r[14],CultureInfo.InvariantCulture),Height=Convert.ToDouble(r[15],CultureInfo.InvariantCulture),ModifiedUtc=Date((string)r[16]),Rank=Convert.ToDouble(r[17],CultureInfo.InvariantCulture)});
                    }
                    db.Execute("COMMIT");return new LibrarySearchPage{Total=total,Slides=rows};
                }
                catch{db.RollbackIfActive();throw;}
            }
        }
        internal LibrarySlideReference Slide(string slideId,CancellationToken cancellation)
        {
            Id(slideId);cancellation.ThrowIfCancellationRequested();using(var db=Connect(true))using(db.ObserveCancellation(cancellation))
            {
                var row=db.Query("SELECT p.library_id FROM SLIDES s JOIN PRESENTATIONS p ON p.presentation_id=s.presentation_id WHERE s.slide_id=?",slideId).SingleOrDefault();if(row==null)throw new KeyNotFoundException("Indexed slide is no longer present.");
            }
            // Direct identity lookup is independent of position or a UI result page.
            using(var db=Connect(true))using(db.ObserveCancellation(cancellation))
            {
                var r=db.Query("SELECT s.slide_id,p.presentation_id,p.library_id,p.source_path,p.content_hash,s.content_fingerprint,s.native_slide_id,s.title,s.text,s.category,s.author,s.layout,s.master,s.position,s.width,s.height,p.modified_utc FROM SLIDES s JOIN PRESENTATIONS p ON p.presentation_id=s.presentation_id WHERE s.slide_id=? AND p.index_status='indexed'",slideId).SingleOrDefault();
                if(r==null)throw new FileNotFoundException("Indexed source is missing; refresh or relink explicitly.");
                return new LibrarySlideReference{SlideId=(string)r[0],PresentationId=(string)r[1],LibraryId=(string)r[2],SourcePath=(string)r[3],SourceFingerprint=(string)r[4],SlideFingerprint=(string)r[5],NativeId=(string)r[6],Title=(string)r[7],Text=(string)r[8],Category=(string)r[9],Author=(string)r[10],Layout=(string)r[11],Master=(string)r[12],Position=checked((int)(long)r[13]),Width=Convert.ToDouble(r[14],CultureInfo.InvariantCulture),Height=Convert.ToDouble(r[15],CultureInfo.InvariantCulture),ModifiedUtc=Date((string)r[16])};
            }
        }
        internal void VerifySource(LibrarySlideReference slide,CancellationToken cancellation)
        {
            if(slide==null)throw new ArgumentNullException("slide");var current=Slide(slide.SlideId,cancellation);
            if(current.SourcePath!=slide.SourcePath||current.SourceFingerprint!=slide.SourceFingerprint||current.NativeId!=slide.NativeId||current.Position!=slide.Position||!String.Equals(PresentationSource.Fingerprint(current.SourcePath,cancellation),current.SourceFingerprint,StringComparison.Ordinal))throw new InvalidOperationException("Source or indexed selection changed; refresh and reselect before insertion.");
        }
    }
}