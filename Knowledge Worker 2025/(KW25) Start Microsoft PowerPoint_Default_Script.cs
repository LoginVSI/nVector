// TARGET:powerpnt.exe /n
// START_IN:

/////////////
// PowerPoint Start
// Workload: KnowledgeWorker 2025
// Version: 1.0
/////////////

using LoginPI.Engine.ScriptBase;

public class Start_PowerPoint_DefaultScript : ScriptBase
{
    void Execute()
    {
        Wait(seconds:2, showOnScreen:true, onScreenText:"Starting PowerPoint");
        Log("Starting PowerPoint");
        START(mainWindowTitle:"*PowerPoint*", mainWindowClass:"*PPTFrameClass*", timeout:60);
        SkipFirstRunDialogs();
        MainWindow.Maximize();       
        MainWindow.Focus(); 
    }
    private void SkipFirstRunDialogs()
    {
        var dialog = FindWindow(className: "Win32 Window:NUIDialog", processName: "POWERPNT", continueOnError: true, timeout: 5);
        while (dialog != null)
        {
            dialog.Close();
            dialog = FindWindow(className: "Win32 Window:NUIDialog", processName: "POWERPNT", continueOnError: true, timeout: 10);
        }
    }
}