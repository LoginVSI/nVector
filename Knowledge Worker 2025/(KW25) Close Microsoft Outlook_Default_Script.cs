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
        var MainWindow = FindWindow(processName:"outlook", timeout:2, continueOnError:true);
        MainWindow?.Close();
    }
}