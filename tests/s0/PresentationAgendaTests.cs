using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using PowerPointIntelligence.AddIn;
internal static class PresentationAgendaTests
{
    private static int passed;
    private static void Check(bool value,string label){if(!value)throw new Exception("Agenda: "+label);passed++;}
    private static void Reject(Action action,string label){bool rejected=false;try{action();}catch(Exception ex) when(ex is InvalidDataException||ex is InvalidOperationException||ex is ArgumentException){rejected=true;}Check(rejected,label);}
    private static AgendaSlideInfo Slide(int id,int index,int section,string title="Slide title",bool hidden=false,bool agenda=false){return new AgendaSlideInfo {SlideId=id,SlideIndex=index,SectionIndex=section,Title=title,Hidden=hidden,IsAgenda=agenda};}
    internal static void Run()
    {
        var slides=new[]{Slide(11,1,1,"Opening"),Slide(12,2,1,"Hidden",true),Slide(13,3,2,"Existing agenda",false,true),Slide(14,4,2,"Findings"),Slide(15,5,3,"Hidden chapter",true)};
        var sections=new[]{new AgendaSectionInfo {SectionIndex=1,Name="Introduction"},new AgendaSectionInfo {SectionIndex=2,Name="Results"},new AgendaSectionInfo {SectionIndex=3,Name="Unshown"},new AgendaSectionInfo {SectionIndex=4,Name="Empty"}};
        var plan=AgendaPlanner.Build(slides,sections,960,540,14);
        Check(plan.UsesSections&&plan.Entries.Count==2,"native sections take precedence over slide titles");
        Check(plan.Entries[0].Caption=="Introduction"&&plan.Entries[1].Caption=="Results","section names preserved");
        Check(plan.Entries[1].TargetSlideId==14&&plan.Entries[1].TargetSlideIndex==4,"hidden and managed agenda targets excluded");
        Check(plan.Entries[1].Highlight&&!plan.Entries[0].Highlight,"current section highlight");
        Check(plan.Entries[1].SubAddress=="14,4,","stable slide ID and current index hyperlink without title parsing");
        Check(AgendaPlanner.Build(slides,sections,960,540,14,false).Entries.All(e=>!e.Highlight),"highlight can be disabled");
        var fallback=AgendaPlanner.Build(new[]{Slide(4,2,0,"Line one\nline two"),Slide(3,1,0,null),Slide(5,3,0,"skip",false,true)},new AgendaSectionInfo[0],720,540,4);
        Check(!fallback.UsesSections&&fallback.Entries.Count==2,"title fallback only when sections absent");
        Check(fallback.Entries[0].Caption=="Slide 1"&&fallback.Entries[1].Caption=="Line one line two","missing and multiline title handling");
        Check(fallback.Entries[0].Key=="item:3"&&fallback.Entries[1].Highlight,"fallback retains target identity");
        foreach(var dimensions in new[]{new[]{960f,540f},new[]{720f,540f},new[]{540f,960f}}){
            var many=AgendaPlanner.Build(Enumerable.Range(1,20).Select(i=>Slide(100+i,i,0,"Topic "+i)),new AgendaSectionInfo[0],dimensions[0],dimensions[1]);
            Check(many.Entries.Count==20,"all source entries retained");
            for(int i=0;i<many.Entries.Count;i++){var a=many.Entries[i];Check(a.Left>=40&&a.Top>=128&&a.Left+a.Width<=dimensions[0]-40+.01&&a.Top+a.Height<=dimensions[1]-40+.01,"entry stays within slide bounds");for(int j=i+1;j<many.Entries.Count;j++){var b=many.Entries[j];Check(a.Left+a.Width<=b.Left||b.Left+b.Width<=a.Left||a.Top+a.Height<=b.Top||b.Top+b.Height<=a.Top,"entries do not overlap");}}
        }
        Reject(()=>AgendaPlanner.Build(new AgendaSlideInfo[0],sections,960,540),"empty content rejected");
        Reject(()=>AgendaPlanner.Build(new[]{Slide(1,1,0,hidden:true)},new AgendaSectionInfo[0],960,540),"hidden-only content rejected");
        Reject(()=>AgendaPlanner.Build(new[]{Slide(1,1,0),Slide(1,2,0)},new AgendaSectionInfo[0],960,540),"duplicate source slide IDs rejected");
        Reject(()=>AgendaPlanner.Build(new[]{Slide(1,1,0),Slide(2,1,0)},new AgendaSectionInfo[0],960,540),"duplicate source slide indexes rejected");
        Reject(()=>AgendaPlanner.Build(slides,new[]{sections[0],sections[0]},960,540),"duplicate sections rejected");
        Reject(()=>AgendaPlanner.Build(slides,sections,Single.NaN,540),"nonfinite dimensions rejected");
        Reject(()=>AgendaPlanner.Build(slides,sections,300,540),"unreadable canvas rejected");
        Reject(()=>AgendaPlanner.Build(new[]{Slide(1,1,0,new String('x',241))},new AgendaSectionInfo[0],960,540),"oversize captions rejected rather than truncated");
        Reject(()=>AgendaPlanner.Build(Enumerable.Range(1,100).Select(i=>Slide(i,i,0)),new AgendaSectionInfo[0],960,540),"overflowing fallback rejected before Office mutation");
        Check(AgendaPlanner.Build(new[]{Slide(21,1,0),Slide(22,2,1)},new[]{sections[0]},960,540).Entries[0].Caption=="Unsectioned slides","unsectioned prefix is represented");
        Reject(()=>AgendaPlanner.Build(new[]{Slide(22,1,7)},new[]{sections[0]},960,540),"unknown section references rejected");
        var moved=AgendaPlanner.Build(new[]{Slide(14,1,2,"Findings"),Slide(11,2,1,"Opening")},sections,960,540);
        Check(moved.Entries.Single(e=>e.TargetSlideId==14).SubAddress=="14,1,","refresh rebuilds index hint while retaining stable target ID");
        Console.WriteLine("PASS "+passed+" agenda planning/identity/layout assertions. Native insertion, hyperlinks and save/reopen remain unverified.");
    }
}
