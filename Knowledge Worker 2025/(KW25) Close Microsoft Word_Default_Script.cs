// TARGET:winword /t
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
        Wait(seconds:2, showOnScreen:true, onScreenText:"Closing Word if open.");
        Log("Closing Word if open.");
        var MainWindow = FindWindow(processName:"winword", timeout:5, continueOnError:true);
        Wait(2);
        MainWindow?.Focus();
        MainWindow?.Maximize();
        Wait(2);
        MainWindow?.Close();
        Wait(1);
        Type("n");  
    }
}