// TARGET:excel.exe /e
// START_IN:

/////////////
// Excel Start
// Workload: KnowledgeWorker 2025
// Version: 1.0
/////////////

using LoginPI.Engine.ScriptBase;

public class Start_Excel_DefaultScript : ScriptBase
{
    const string ProcessName = "EXCEL";
    void Execute()
    {                
        Wait(seconds:2, showOnScreen:true, onScreenText:"Starting Excel");
        Log("Starting Excel");
        START(mainWindowTitle: "*Excel*", mainWindowClass: "*XLMAIN*", timeout: 60);
        SkipFirstRunDialogs();
        MainWindow.Maximize();
        MainWindow.Focus();
    }
    void SkipFirstRunDialogs()
    {
        var dialog = FindWindow(className: "Win32 Window:NUIDialog", processName: ProcessName, continueOnError: true, timeout: 5);
        while (dialog != null)
        {
            dialog.Close();
            dialog = FindWindow(className: "Win32 Window:NUIDialog", processName: ProcessName, continueOnError: true, timeout: 10);
        }
    }
}