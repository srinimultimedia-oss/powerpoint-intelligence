using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
namespace PowerPointIntelligence.Library
{
    internal sealed class LibraryCategorySummary
    {
        internal string Category {get;set;}
        internal long Slides {get;set;}
    }
    internal sealed class LibraryPresentationSummary
    {
        internal string Id {get;set;}
        internal string File {get;set;}
        internal string LibraryId {get;set;}
        internal long Slides {get;set;}
        public override string ToString(){return System.IO.Path.GetFileName(File);}
    }
    internal sealed partial class PiidxCatalog
    {
        internal IReadOnlyList<LibraryCategorySummary> Categories(string libraryId,CancellationToken cancellation)
        {
            if(libraryId!=null)Id(libraryId);using(var db=Connect(true))using(db.ObserveCancellation(cancellation))return db.Query("SELECT s.category,count(*) FROM SLIDES s JOIN PRESENTATIONS p ON p.presentation_id=s.presentation_id JOIN LIBRARIES l ON l.library_id=p.library_id WHERE p.index_status='indexed' AND l.enabled=1"+(libraryId==null?"":" AND p.library_id=?")+" GROUP BY s.category ORDER BY s.category",libraryId==null?new object[0]:new object[]{libraryId}).Select(r=>new LibraryCategorySummary{Category=(string)r[0],Slides=(long)r[1]}).ToList();
        }
        internal IReadOnlyList<LibraryPresentationSummary> Presentations(string libraryId,CancellationToken cancellation)
        {
            if(libraryId!=null)Id(libraryId);using(var db=Connect(true))using(db.ObserveCancellation(cancellation))return db.Query("SELECT p.presentation_id,p.source_path,p.library_id,p.slide_count FROM PRESENTATIONS p JOIN LIBRARIES l ON l.library_id=p.library_id WHERE p.index_status='indexed' AND l.enabled=1"+(libraryId==null?"":" AND p.library_id=?")+" ORDER BY p.source_path",libraryId==null?new object[0]:new object[]{libraryId}).Select(r=>new LibraryPresentationSummary{Id=(string)r[0],File=(string)r[1],LibraryId=(string)r[2],Slides=(long)r[3]}).ToList();
        }
    }
}