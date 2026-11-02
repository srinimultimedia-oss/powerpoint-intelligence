using System;
using System.IO;
using System.Linq;
using PowerPointIntelligence.AddIn;
internal static class ProductivityToolTests
{
    internal static int Run()
    {
        int count=0;Action<bool,string> check=(pass,name)=>{if(!pass)throw new Exception("FAIL productivity "+name);Console.WriteLine("PASS productivity "+name);count++;};
        check(ProductivityTools.Ids.Distinct().Count()==4,"four unique tools");check(ProductivityTools.SearchSymbols(null).Length==ProductivityTools.Symbols.Length,"empty search");check(ProductivityTools.SearchSymbols(" CHECK ").Any(s=>s.Glyph=="✓"),"case insensitive name search");check(ProductivityTools.SearchSymbols("U+2192").Single().Glyph=="→","codepoint search");check(ProductivityTools.SearchSymbols("→").Single().Glyph=="→","glyph search");check(ProductivityTools.SearchSymbols("not-a-symbol").Length==0,"no fake results");check(ProductivityTools.Symbols.Select(s=>s.Glyph).Distinct().Count()==ProductivityTools.Symbols.Length,"unique symbols");check(ProductivityTools.Symbols.All(s=>s.Glyph.Length>0&&s.Name.Length>0),"named catalog entries");
        var changes=ProductivityTools.DecimalChanges("Revenue 1,234.50; ratio 2.4 and words, punctuation.");check(changes.Length==3,"digit-adjacent separators only");check(changes[0].Before==','&&changes[0].After=='.',"comma to dot");check(changes[1].Before=='.'&&changes[1].After==',',"dot to comma");check(changes.All(c=>c.Index>0),"zero based change positions");
        check(ProductivityTools.DecimalChanges("").Length==0,"empty input");check(ProductivityTools.DecimalChanges(".1 2. a,b a.b ,").Length==0,"edge and nonnumeric punctuation retained");check(ProductivityTools.DecimalChanges("١٫٢").Length==0,"unsupported separator retained");check(ProductivityTools.DecimalChanges("😀1.2").Single().Index==3,"UTF16 offset after surrogate pair");check(ProductivityTools.DecimalChanges("1.2.3").Length==2,"date-like literal policy explicit");check(ProductivityTools.DecimalChanges("１.２").Length==1,"Unicode digits recognized");
        var source="1.2 and 3,4";var chars=source.ToCharArray();foreach(var c in ProductivityTools.DecimalChanges(source))chars[c.Index]=c.After;var swapped=new string(chars);check(swapped=="1,2 and 3.4","planned substitutions");chars=swapped.ToCharArray();foreach(var c in ProductivityTools.DecimalChanges(swapped))chars[c.Index]=c.After;check(new string(chars)==source,"swap reversible");
        Action<Action,string> rejects=(action,name)=>{bool rejected=false;try{action();}catch(ArgumentException){rejected=true;}catch(InvalidDataException){rejected=true;}check(rejected,name);};rejects(()=>ProductivityTools.DecimalChanges(null),"null rejected");rejects(()=>ProductivityTools.DecimalChanges(new string('1',20001)),"oversize rejected");rejects(()=>ProductivityTools.SearchSymbols(new string('q',101)),"search bound");check(ProductivityTools.DecimalChanges(new string('1',20000)).Length==0,"maximum accepted");
        Console.WriteLine(count+" productivity pure assertions passed; Office integration not executed.");return count;
    }
}
