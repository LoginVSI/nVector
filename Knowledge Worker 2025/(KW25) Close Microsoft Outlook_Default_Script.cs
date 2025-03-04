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
        int globalWaitInSeconds = 3; // Standard wait time between actions
        Wait(seconds:2, showOnScreen:true, onScreenText:"Closing Outlook if open.");
        Log("Closing Outlook if open.");
        var MainWindow = FindWindow(processName:"outlook", title:"* - Outlook*", timeout:5, continueOnError:true);
        Wait(globalWaitInSeconds);
        MainWindow?.Focus();
        MainWindow?.Maximize();
        Wait(globalWaitInSeconds);
        MainWindow?.Close();
        Wait(globalWaitInSeconds);
        MainWindow?.Type("{alt+f4}")
        Wait(globalWaitInSeconds);
        MainWindow?.Type("n");
        Wait(globalWaitInSeconds);
    }
}