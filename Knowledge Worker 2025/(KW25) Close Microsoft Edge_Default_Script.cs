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
        int globalWaitInSeconds = 3; // Standard wait time between actions
        Wait(seconds:2, showOnScreen:true, onScreenText:"Closing browser if open.");
        Log("Closing browser if open.");
        var MainWindow = FindWindow(processName:"msedge", timeout:5, continueOnError:true);
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