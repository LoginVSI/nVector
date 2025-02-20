// TARGET:excel
// START_IN:

/////////////
// Excel Close
// Workload: KnowledgeWorker 2025
// Version: 1.0
/////////////

using LoginPI.Engine.ScriptBase;

public class Close_Excel_DefaultScript : ScriptBase
{
    void Execute()
    {
        var MainWindow = FindWindow(processName:"excel", timeout:3, continueOnError:true);
        MainWindow?.Close();
    }
}