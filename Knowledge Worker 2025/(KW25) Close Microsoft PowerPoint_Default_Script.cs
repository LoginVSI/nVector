// TARGET:powerpnt
// START_IN:

/////////////
// PowerPoint Close
// Workload: KnowledgeWorker 2025
// Version: 1.0
/////////////

using LoginPI.Engine.ScriptBase;

public class Close_PowerPoint_DefaultScript : ScriptBase
{
    void Execute()
    {
        var MainWindow = FindWindow(processName:"powerpnt", timeout:2, continueOnError:true);
        MainWindow?.Close();
    }
}