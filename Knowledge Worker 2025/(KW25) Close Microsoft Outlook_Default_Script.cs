// TARGET:outlook.exe /importprf %TEMP%\LoginEnterprise\Outlook.prf
// START_IN:

/////////////
// Outlook Close
// Workload: KnowledgeWorker 2025
// Version: 1.0
/////////////

using LoginPI.Engine.ScriptBase;

public class Close_Outlook_DefaultScript : ScriptBase
{
    void Execute()
    {
        Wait(seconds:2, showOnScreen:true, onScreenText:"Closing Outlook if open.");
        Log("Closing Outlook if open.");
        var MainWindow = FindWindow(processName:"outlook", title:"* - Outlook*", timeout:5, continueOnError:true);
        Wait(2);
        MainWindow?.Focus();
        MainWindow?.Maximize();
        Wait(2);
        MainWindow?.Close();
    }
}