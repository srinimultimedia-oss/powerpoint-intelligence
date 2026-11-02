using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using PowerPointIntelligence.AddIn;
internal static class ChapterAgendaTests
{
    private static int count;
    private static void Check(bool value,string name){if(!value)throw new Exception("FAIL chapter agenda: "+name);count++;}
    private static void Reject(Action action,string name){try{action();}catch(Exception ex) when(ex is InvalidDataException||ex is InvalidOperationException){count++;return;}throw new Exception("FAIL expected rejection: "+name);}
    private static ChapterRecord R(int id,int level=1,int scope=0){return new ChapterRecord {SlideId=id,SlideIndex=id,Level=level,Scope=scope,Title="Chapter "+id};}
    private static AgendaView V(int current=1,bool contents=false,bool all=false,int scope=0){return new AgendaView {SlideId=1000,SlideIndex=1000,CurrentChapterId=current,Contents=contents,ShowAll=all,Scope=scope};}
    internal static void Run()
    {
        count=0;var records=new[]{R(1),R(2,2),R(3,3),R(4,4),R(5,3),R(6,2),R(7),R(8,2),R(9,1,99),R(10,2,99)};
        var plan=ChapterAgendaPlanner.Build(records.Reverse(),V(3),960,540);
        Check(plan.Entries.Select(e=>e.TargetSlideId).SequenceEqual(new[]{1,2,3,4,5,7}),"roots path siblings and children");
        Check(plan.Entries.Single(e=>e.TargetSlideId==3).Number=="1.1.1","dotted numbering");
        Check(plan.Entries.Count(e=>e.Highlight)==1&&plan.Entries.Single(e=>e.Highlight).TargetSlideId==3,"current highlighted");
        Check(plan.Entries.Where(e=>e.Ancestor).Select(e=>e.TargetSlideId).SequenceEqual(new[]{1,2}),"ancestor roles");
        var toc=ChapterAgendaPlanner.Build(records,V(0,true),960,540);Check(toc.Entries.Select(e=>e.TargetSlideId).SequenceEqual(new[]{1,7})&&!toc.Entries.Any(e=>e.Highlight||e.Ancestor),"contents roots without highlight");
        Check(ChapterAgendaPlanner.Build(records,V(0,true,true),960,540).Entries.Count==8,"contents all levels");
        Check(ChapterAgendaPlanner.Build(records,V(10,false,false,99),960,540).Entries.Select(e=>e.TargetSlideId).SequenceEqual(new[]{9,10}),"scope isolated");
        Check(ChapterAgendaPlanner.Build(new ChapterRecord[0],V(0,true),400,300).Entries.Count==0,"empty contents allowed");
        foreach(var e in plan.Entries)Check(e.Left>=0&&e.Top>=0&&e.Width>0&&e.Height>0&&e.Left+e.Width<=960&&e.Top+e.Height<=540,"layout inside canvas");
        Reject(()=>ChapterAgendaPlanner.Build(records,V(99),960,540),"unknown current");Reject(()=>ChapterAgendaPlanner.Build(records,V(10),960,540),"current wrong scope");Reject(()=>ChapterAgendaPlanner.Build(new[]{R(1,2)},V(),960,540),"missing root");Reject(()=>ChapterAgendaPlanner.Build(new[]{R(1),R(2,3)},V(),960,540),"level jump");Reject(()=>ChapterAgendaPlanner.Build(new[]{R(1),R(1)},V(),960,540),"duplicate identity");
        var sameIndex=new[]{R(1),R(2)};sameIndex[1].SlideIndex=1;Reject(()=>ChapterAgendaPlanner.Build(sameIndex,V(),960,540),"duplicate position");Reject(()=>ChapterAgendaPlanner.Build(new[]{R(1,5)},V(),960,540),"level bound");Reject(()=>ChapterAgendaPlanner.Build(new[]{R(1,1,-1)},V(),960,540),"scope bound");Reject(()=>ChapterAgendaPlanner.Build(records,V(),399,300),"minimum size");Reject(()=>ChapterAgendaPlanner.Build(records,V(),Single.NaN,540),"finite size");Reject(()=>ChapterAgendaPlanner.Build(Enumerable.Range(1,201).Select(i=>R(i)),V(),960,540),"record bound");Reject(()=>ChapterAgendaPlanner.Build(Enumerable.Range(1,30).Select(i=>R(i)),V(1,false,true),400,300),"overflow refused");
        var baseline=new Dictionary<int,string>{{1,"Original"},{2,"Second"}};var resolved=ChapterAgendaPlanner.ResolveEdits(baseline,new[]{new ChapterCaptionEdit {TargetSlideId=1,Title="Original"},new ChapterCaptionEdit {TargetSlideId=1,Title="New\r\nname"},new ChapterCaptionEdit {TargetSlideId=1,Title="New name"}});Check(resolved[1]=="New name"&&resolved[2]=="Second"&&baseline[1]=="Original","edits normalized without baseline mutation");
        Reject(()=>ChapterAgendaPlanner.ResolveEdits(baseline,new[]{new ChapterCaptionEdit {TargetSlideId=1,Title="A"},new ChapterCaptionEdit {TargetSlideId=1,Title="B"}}),"conflicting edits");Reject(()=>ChapterAgendaPlanner.ResolveEdits(baseline,new[]{new ChapterCaptionEdit {TargetSlideId=3,Title="A"}}),"unknown edit");Reject(()=>ChapterAgendaPlanner.Caption(new string('a',241)),"caption bound");Reject(()=>ChapterAgendaPlanner.Caption("\t"),"empty caption");Reject(()=>ChapterAgendaPlanner.Caption("bad\u0001"),"control caption");
        var state=new ChapterAgendaState {Chapters=records.ToList(),Baseline=new Dictionary<int,string>{{1,"Caption < & >"}}};string xml=ChapterAgendaCodec.Encode(state);var copy=ChapterAgendaCodec.Decode(xml);Check(copy.Chapters.Count==10&&copy.Baseline[1]=="Caption < & >","XML roundtrip");Check(ChapterAgendaCodec.Encode(copy)==xml,"deterministic codec");
        Reject(()=>ChapterAgendaCodec.Decode(xml.Replace("version=\"2\"","version=\"3\"")),"unknown schema");Reject(()=>ChapterAgendaCodec.Decode(xml.Replace("agenda:v2","agenda:v9")),"unknown namespace");Reject(()=>ChapterAgendaCodec.Decode(xml.Replace("<chapters>","<chapters surprise=\"yes\">")),"unknown attribute");Reject(()=>ChapterAgendaCodec.Decode(xml.Replace("</chapters>","<unknown /></chapters>")),"unknown node");Reject(()=>ChapterAgendaCodec.Decode(xml.Replace("<chapters>","<chapters><!--comment-->")),"comments refused");Reject(()=>ChapterAgendaCodec.Decode(xml.Replace("<chapters>","<chapters>unexpected")),"text refused");Reject(()=>ChapterAgendaCodec.Decode("<!DOCTYPE agenda [<!ENTITY x 'test'>]>"+xml),"DTD refused");Reject(()=>ChapterAgendaCodec.Decode("<!DOCTYPE agenda SYSTEM 'file:///no-read'>"+xml),"external DTD refused");Reject(()=>ChapterAgendaCodec.Decode(xml.Substring(1)),"malformed XML");Reject(()=>ChapterAgendaCodec.Decode(new string('x',300001)),"XML size bound");Reject(()=>ChapterAgendaCodec.Decode(xml.Replace("slideId=\"1\" title=", "slideId=\"999\" title=")),"unknown baseline");
        // Content gaps and nested chapters are part of the moved block, never detached.
        var moves=new[]{R(2),R(4,2),R(6,3),R(8,2),R(11),R(13,2),R(17,1,900),R(20,1,900),R(22,2,900)};
        var boundaries=new Dictionary<int,int>{{900,16}};
        var move=ChapterAgendaPlanner.Move(moves,2,true,25,boundaries);
        Check(move.FirstSlide==2&&move.LastSlide==10&&move.DestinationIndex==7,"root down moves descendants and content through sibling block");
        move=ChapterAgendaPlanner.Move(moves,11,false,25,boundaries);
        Check(move.FirstSlide==11&&move.LastSlide==15&&move.DestinationIndex==2,"root up preserves next splitter");
        move=ChapterAgendaPlanner.Move(moves,4,true,25,boundaries);
        Check(move.FirstSlide==4&&move.LastSlide==7&&move.DestinationIndex==7,"nested down includes grandchild and content");
        move=ChapterAgendaPlanner.Move(moves,8,false,25,boundaries);
        Check(move.FirstSlide==8&&move.LastSlide==10&&move.DestinationIndex==4,"nested up before prior sibling");
        move=ChapterAgendaPlanner.Move(moves,17,true,25,boundaries);
        Check(move.FirstSlide==17&&move.LastSlide==19&&move.DestinationIndex==23,"last scope down includes trailing content");
        move=ChapterAgendaPlanner.Move(moves,20,false,25,boundaries);
        Check(move.FirstSlide==20&&move.LastSlide==25&&move.DestinationIndex==17,"last scope up ends at deck end");
        Reject(()=>ChapterAgendaPlanner.Move(moves,11,true,25,boundaries),"cannot move across splitter");
        Reject(()=>ChapterAgendaPlanner.Move(moves,2,false,25,boundaries),"first sibling cannot move up");
        Reject(()=>ChapterAgendaPlanner.Move(moves,6,true,25,boundaries),"different parent is not sibling");
        Reject(()=>ChapterAgendaPlanner.Move(moves,13,true,25,boundaries),"child cannot leave parent");
        Reject(()=>ChapterAgendaPlanner.Move(moves,999,true,25,boundaries),"unknown moving identity");
        Reject(()=>ChapterAgendaPlanner.Move(moves,2,true,21,boundaries),"record outside deck");
        Reject(()=>ChapterAgendaPlanner.Move(moves,2,true,25,new Dictionary<int,int>()),"missing scope boundary");
        Reject(()=>ChapterAgendaPlanner.Move(moves,2,true,25,new Dictionary<int,int>{{900,17}}),"splitter overlaps chapter");
        Reject(()=>ChapterAgendaPlanner.Move(moves,2,true,25,new Dictionary<int,int>{{900,15},{901,15}}),"duplicate splitter positions");
        Reject(()=>ChapterAgendaPlanner.Move(moves,2,true,25,new Dictionary<int,int>{{0,2},{900,16}}),"default scope starts at one");
        Reject(()=>ChapterAgendaPlanner.Move(moves,2,true,25,new Dictionary<int,int>{{900,26}}),"splitter outside deck");
        Reject(()=>ChapterAgendaPlanner.Move(moves,2,true,25,new Dictionary<int,int>{{900,12}}),"chapter membership mismatch");
        // Simulate PowerPoint final insertion-index semantics and verify entire-block order.
        move=ChapterAgendaPlanner.Move(moves,2,true,25,boundaries);var sequence=Enumerable.Range(1,25).ToList();var block=sequence.GetRange(move.FirstSlide-1,move.LastSlide-move.FirstSlide+1);sequence.RemoveRange(move.FirstSlide-1,block.Count);sequence.InsertRange(move.DestinationIndex-1,block);
        Check(sequence.Take(15).SequenceEqual(new[]{1,11,12,13,14,15,2,3,4,5,6,7,8,9,10})&&sequence[15]==16,"move index matches final contiguous range semantics");
        Check(moves[0].SlideIndex==2&&boundaries[900]==16,"move planning does not mutate records or boundaries");
        Console.WriteLine("Chapter agenda tests passed: "+count);
    }
}

