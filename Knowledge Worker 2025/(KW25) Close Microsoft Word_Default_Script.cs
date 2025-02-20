// TARGET:winword
// START_IN:

/////////////
// Word Close
// Workload: KnowledgeWorker 2025
// Version: 1.0
/////////////

using LoginPI.Engine.ScriptBase;

public class Close_Word_DefaultScript : ScriptBase
{
    void Execute()
    {
        var MainWindow = FindWindow(processName:"winword", timeout:3, continueOnError:true);
        MainWindow?.Close();
    }
}