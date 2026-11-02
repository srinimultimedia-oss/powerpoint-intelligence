using System;
using System.Windows.Forms;

namespace PowerPointIntelligence.AddIn
{
    // Per-instance input ports. Production uses the real dialogs when a delegate
    // is absent. Tests can exercise document commands with explicit fixture inputs
    // without claiming that the native picker UI itself has been visually tested.
    internal sealed class DemoInputs
    {
        internal Func<string,string,string> Prompt { get; set; }
        internal Func<string,string> OpenFile { get; set; }
        internal Func<string,string,string> SaveFile { get; set; }
        internal Func<string> Folder { get; set; }
        internal Func<string,string,bool> Confirm { get; set; }
        internal Action<string,string> Notify { get; set; }
        internal Func<string[,],string[,]> EditTable { get; set; }
        internal Func<string[],int> ChooseAsset { get; set; }
        internal Action<Form> ShowToolbar { get; set; }
        internal Action<Microsoft.Office.Interop.PowerPoint.Presentation> PresentationCreated { get; set; }
    }
}
