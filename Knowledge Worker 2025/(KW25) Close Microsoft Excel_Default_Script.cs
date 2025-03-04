// TARGET:excel /e
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
        int globalWaitInSeconds = 3; // Standard wait time between actions
        Wait(seconds:2, showOnScreen:true, onScreenText:"Closing Excel if open.");
        Log("Closing Excel if open.");
        var MainWindow = FindWindow(processName:"excel", timeout:5, continueOnError:true, title: "* Excel*");
        Wait(globalWaitInSeconds);
        MainWindow?.Focus();
        MainWindow?.Maximize();
        Wait(globalWaitInSeconds);
        MainWindow?.Close();
        Wait(globalWaitInSeconds);
        MainWindow?.Type("{alt+f4}")
        Wait(globalWaitInSeconds);
        MainWindow?.Type("n");
        Wait(globalWaitInSeconds);
    }
}