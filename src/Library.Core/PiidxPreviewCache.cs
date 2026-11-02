using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Threading;
namespace PowerPointIntelligence.Library
{
    internal static class PiidxPng
    {
        internal const int MaximumBytes=8*1024*1024;
        static readonly byte[] Signature={137,80,78,71,13,10,26,10};
        static readonly uint[] Table=BuildTable();
        static uint[] BuildTable(){var table=new uint[256];for(uint n=0;n<256;n++){uint value=n;for(int k=0;k<8;k++)value=(value&1)!=0?0xedb88320u^(value>>1):value>>1;table[n]=value;}return table;}
        static uint Integer(byte[] bytes,int at){return ((uint)bytes[at]<<24)|((uint)bytes[at+1]<<16)|((uint)bytes[at+2]<<8)|bytes[at+3];}
        static uint Crc(byte[] bytes,int at,int length){uint value=0xffffffffu;for(int i=0;i<length;i++)value=Table[(value^bytes[at+i])&255]^(value>>8);return value^0xffffffffu;}
        internal static void Validate(byte[] png,int width,int height)
        {
            if(png==null||png.Length<45||png.Length>MaximumBytes||width<1||height<1||width>4096||height>4096||!png.Take(8).SequenceEqual(Signature))throw new InvalidDataException("Preview PNG header or size is invalid.");
            int at=8;bool header=false,data=false,end=false,palette=false,dataEnded=false;int color=-1;
            while(at<png.Length)
            {
                if(png.Length-at<12)throw new InvalidDataException("Preview PNG chunk is truncated.");uint rawLength=Integer(png,at);if(rawLength>MaximumBytes||rawLength>(uint)(png.Length-at-12))throw new InvalidDataException("Preview PNG chunk exceeds bounds.");int length=(int)rawLength;
                string type=System.Text.Encoding.ASCII.GetString(png,at+4,4);
                if(type.Any(c=>!(c>='A'&&c<='Z')&&!(c>='a'&&c<='z'))||type[2]<'A'||type[2]>'Z')throw new InvalidDataException("Preview PNG chunk type is invalid.");
                if(Crc(png,at+4,length+4)!=Integer(png,at+8+length))throw new InvalidDataException("Preview PNG checksum failed.");
                if(!header&&type!="IHDR")throw new InvalidDataException("Preview PNG must start with IHDR.");
                if(type=="IHDR")
                {
                    if(header||length!=13||Integer(png,at+8)!=(uint)width||Integer(png,at+12)!=(uint)height)throw new InvalidDataException("Preview PNG dimensions differ from metadata.");
                    int bits=png[at+16];color=png[at+17];bool valid=(color==0&&(bits==1||bits==2||bits==4||bits==8||bits==16))||((color==2||color==4||color==6)&&(bits==8||bits==16))||(color==3&&(bits==1||bits==2||bits==4||bits==8));
                    if(!valid||png[at+18]!=0||png[at+19]!=0||png[at+20]>1)throw new InvalidDataException("Preview PNG encoding is unsupported.");header=true;
                }
                else if(type=="PLTE"){if(palette||data||length==0||length>768||length%3!=0)throw new InvalidDataException("Preview PNG palette is invalid.");palette=true;}
                else if(type=="IDAT"){if(dataEnded||length==0||(color==3&&!palette))throw new InvalidDataException("Preview PNG image data is invalid.");data=true;}
                else if(type=="IEND"){if(!data||length!=0||at+12!=png.Length)throw new InvalidDataException("Preview PNG ending is invalid.");end=true;}
                else{if(type[0]>='A'&&type[0]<='Z')throw new InvalidDataException("Unknown critical preview PNG chunk.");if(data)dataEnded=true;}
                at+=length+12;
            }
            if(!end)throw new InvalidDataException("Preview PNG is incomplete.");
        }
        internal static string Hash(byte[] bytes){using(var hash=SHA256.Create())return BitConverter.ToString(hash.ComputeHash(bytes)).Replace("-","");}
    }
    internal sealed partial class PiidxCatalog
    {
        internal void PutPreview(LibrarySlideReference slide,byte[] png,int width,int height,CancellationToken cancellation)
        {
            if(slide==null)throw new ArgumentNullException("slide");Id(slide.SlideId);cancellation.ThrowIfCancellationRequested();
            png=png==null?null:(byte[])png.Clone();PiidxPng.Validate(png,width,height);string checksum=PiidxPng.Hash(png);
            using(WriterLock(cancellation))using(var db=Connect(false))using(db.ObserveCancellation(cancellation))
            {
                db.Execute("BEGIN IMMEDIATE");try
                {
                    var current=db.Query("SELECT s.content_fingerprint,p.content_hash,p.source_path FROM SLIDES s JOIN PRESENTATIONS p ON p.presentation_id=s.presentation_id WHERE s.slide_id=? AND p.index_status='indexed'",slide.SlideId).SingleOrDefault();
                    if(current==null||(string)current[0]!=slide.SlideFingerprint||(string)current[1]!=slide.SourceFingerprint||(string)current[2]!=slide.SourcePath)throw new InvalidOperationException("Preview selection is stale; source snapshot changed.");
                    db.Execute("DELETE FROM SLIDE_PREVIEWS WHERE slide_id=? AND width=? AND height=?",slide.SlideId,width,height);
                    db.Execute("INSERT INTO SLIDE_PREVIEWS VALUES(?,?,?,'image/png',?,?,?,?)",slide.SlideId,width,height,png,checksum,slide.SlideFingerprint,Utc());
                    long bytes=(long)db.Query("SELECT coalesce(sum(length(image)),0) FROM SLIDE_PREVIEWS").Single()[0];
                    while(bytes>256L*1024*1024){var victim=db.Query("SELECT rowid,length(image) FROM SLIDE_PREVIEWS ORDER BY rendered_utc,slide_id,width,height LIMIT 1").Single();db.Execute("DELETE FROM SLIDE_PREVIEWS WHERE rowid=?",victim[0]);bytes-=(long)victim[1];}
                    cancellation.ThrowIfCancellationRequested();db.Execute("COMMIT");
                }
                catch{db.RollbackIfActive();throw;}
            }
        }
        internal byte[] Preview(string slideId,int width,int height,CancellationToken cancellation)
        {
            Id(slideId);cancellation.ThrowIfCancellationRequested();if(width<1||height<1||width>4096||height>4096)throw new ArgumentException("Preview dimensions exceed supported bounds.");
            using(var db=Connect(true))using(db.ObserveCancellation(cancellation))
            {
                var row=db.Query("SELECT c.image,c.content_hash,c.source_fingerprint,s.content_fingerprint FROM SLIDE_PREVIEWS c JOIN SLIDES s ON s.slide_id=c.slide_id JOIN PRESENTATIONS p ON p.presentation_id=s.presentation_id WHERE c.slide_id=? AND c.width=? AND c.height=? AND p.index_status='indexed'",slideId,width,height).SingleOrDefault();
                if(row==null||!Object.Equals(row[2],row[3]))return null;byte[] png=(byte[])row[0];
                if(PiidxPng.Hash(png)!=(string)row[1])throw new InvalidDataException("Cached preview content hash failed; regenerate it from its source.");
                PiidxPng.Validate(png,width,height);return png;
            }
        }
    }
}