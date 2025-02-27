// TARGET:powerpnt /n
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
        Wait(seconds:2, showOnScreen:true, onScreenText:"Closing PowerPoint if open.");
        Log("Closing PowerPoint if open.");
        var MainWindow = FindWindow(processName:"powerpnt", timeout:5, continueOnError:true);
        MainWindow?.Focus();
        MainWindow?.Maximize();
        MainWindow?.Close();
    }
}