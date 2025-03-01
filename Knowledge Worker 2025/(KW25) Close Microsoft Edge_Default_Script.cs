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
        Wait(seconds:2, showOnScreen:true, onScreenText:"Closing browser if open.");
        Log("Closing browser if open.");
        var MainWindow = FindWindow(processName:"msedge", timeout:5, continueOnError:true);
        Wait(2);
        MainWindow?.Focus();
        MainWindow?.Maximize();
        Wait(2);
        MainWindow?.Close();
    }
}