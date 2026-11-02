using System;
using System.Drawing;
using System.Windows.Forms;
internal sealed class ContextForegroundPrompt:Form
{
    internal bool StartRequested {get;private set;}
    internal bool RequestAccepted {get;private set;}
    private readonly Timer deadline=new Timer {Interval=60000};
    internal ContextForegroundPrompt(Func<bool> request)
    {
        Text="PowerPoint test · ready to continue";ClientSize=new Size(510,180);Font=new Font("Segoe UI",10);StartPosition=FormStartPosition.CenterScreen;FormBorderStyle=FormBorderStyle.FixedDialog;MaximizeBox=MinimizeBox=false;
        var message=new Label {Left=20,Top=18,Width=470,Height=88,Text="Windows declined the automatic focus request.\r\nClick Start test to activate the owned PowerPoint test window.\r\nKeep PowerPoint in front while the checks run."};
        var start=new Button {Text="Start test",Name="StartTest",Left=270,Top=125,Width=110,Height=32};var cancel=new Button {Text="Cancel",Left=390,Top=125,Width=100,Height=32,DialogResult=DialogResult.Cancel};Controls.AddRange(new Control[]{message,start,cancel});AcceptButton=start;CancelButton=cancel;
        start.Click+=(s,e)=>{if(StartRequested)return;StartRequested=true;start.Enabled=false;Hide();RequestAccepted=request();DialogResult=DialogResult.OK;};
        Shown+=(s,e)=>deadline.Start();deadline.Tick+=(s,e)=>{deadline.Stop();DialogResult=DialogResult.Cancel;Close();};
    }
    protected override void Dispose(bool disposing){if(disposing)deadline.Dispose();base.Dispose(disposing);}
}