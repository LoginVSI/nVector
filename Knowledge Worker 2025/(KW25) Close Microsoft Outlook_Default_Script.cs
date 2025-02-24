// TARGET:outlook
// START_IN:

/////////////
// Outlook Close
// Workload: KnowledgeWorker 2025
// Version: 1.0
/////////////

using LoginPI.Engine.ScriptBase;

public class Close_Outlook_DefaultScript : ScriptBase
{
    void Execute()
    {
        Wait(seconds:2, showOnScreen:true, onScreenText:"Closing Outlook if open.");
        Log("Closing Outlook if open.");
        var MainWindow = FindWindow(processName:"outlook", title:"* - Outlook*", timeout:2, continueOnError:true);
        MainWindow?.Focus();
        MainWindow?.Close();
    }
}