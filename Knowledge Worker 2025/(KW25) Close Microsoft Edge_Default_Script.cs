// TARGET:msedge
// START_IN:

/////////////
// Browser Close
// Workload: KnowledgeWorker 2025
// Version: 1.0
/////////////

using LoginPI.Engine.ScriptBase;

public class Close_Browser_DefaultScript : ScriptBase
{
    void Execute()
    {
        var MainWindow = FindWindow(processName:"msedge", timeout:3, continueOnError:true);
        MainWindow?.Close();
    }
}