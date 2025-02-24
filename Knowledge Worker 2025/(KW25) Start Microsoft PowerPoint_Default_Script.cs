// TARGET:powerpnt.exe
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
        START(mainWindowTitle:"*PowerPoint*", mainWindowClass:"*PPTFrameClass*", timeout:60);
        SkipFirstRunDialogs();
    }
    private void SkipFirstRunDialogs()
    {
        var dialog = FindWindow(className: "Win32 Window:NUIDialog", processName: "POWERPNT", continueOnError: true, timeout: 1);
        while (dialog != null)
        {
            dialog.Close();
            dialog = FindWindow(className: "Win32 Window:NUIDialog", processName: "POWERPNT", continueOnError: true, timeout: 10);
        }
    }
}