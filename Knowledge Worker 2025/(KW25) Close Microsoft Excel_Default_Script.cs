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
        Wait(seconds:2, showOnScreen:true, onScreenText:"Closing Excel if open.");
        Log("Closing Excel if open.");
        var MainWindow = FindWindow(processName:"excel", timeout:2, continueOnError:true);
        MainWindow?.Close();
    }
}