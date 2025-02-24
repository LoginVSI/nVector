// TARGET:winword.exe
// START_IN:

/////////////
// Word Start
// Workload: KnowledgeWorker 2025
// Version: 1.0
/////////////

using LoginPI.Engine.ScriptBase;

public class Start_Word_DefaultScript : ScriptBase
{
    void Execute()
    {
        START(mainWindowTitle: "*Word*", mainWindowClass: "Win32 Window:OpusApp", processName: "WINWORD", timeout: 60);
        SkipFirstRunDialogs();
    }
    private void SkipFirstRunDialogs()
    {
        var dialog = FindWindow(className: "Win32 Window:NUIDialog", processName: "WINWORD", continueOnError: true, timeout: 1);
        while (dialog != null)
        {
            dialog.Close();
            dialog = FindWindow(className: "Win32 Window:NUIDialog", processName: "WINWORD", continueOnError: true, timeout: 10);
        }
    }
}