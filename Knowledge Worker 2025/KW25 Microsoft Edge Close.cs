// TARGET:msedge
// START_IN:

/////////////
// Edge Close
// Workload: Knowledge Worker 2025
// Version: 0.1.0
/////////////

using LoginPI.Engine.ScriptBase;

public class Edge_Close : ScriptBase
{
    void Execute()
    {
        int globalWaitInSeconds = 3; // Standard wait time between actions
        Wait(seconds:2, showOnScreen:true, onScreenText:"Closing browser if open.");
        
        Log("Closing browser if open");        
        var MainWindow = FindWindow(processName:"msedge", title: "*- Microsoft​ Edge*", timeout:2, continueOnError:true);
        if (MainWindow != null) {
            MainWindow.Focus();
            MainWindow.Maximize();
            Wait(globalWaitInSeconds);
            MainWindow.Close();
        }
    }
}