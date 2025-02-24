// TARGET:winword.exe /t
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
        Wait(seconds:2, showOnScreen:true, onScreenText:"Starting Word");
        Log("Starting Word");
        START(mainWindowTitle: "*Word*", mainWindowClass: "Win32 Window:OpusApp", processName: "WINWORD", timeout: 60);
        SkipFirstRunDialogs();
        MainWindow.Maximize();
        MainWindow.Focus();
    }
    private void SkipFirstRunDialogs()
    {
        var dialog = FindWindow(className: "Win32 Window:NUIDialog", processName: "WINWORD", continueOnError: true, timeout: 5);
        while (dialog != null)
        {
            dialog.Close();
            dialog = FindWindow(className: "Win32 Window:NUIDialog", processName: "WINWORD", continueOnError: true, timeout: 10);
        }
    }
}