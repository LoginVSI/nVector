// TARGET:outlook.exe /importprf %TEMP%\LoginEnterprise\Outlook.prf
// START_IN:

/////////////
// Outlook Close
// Workload: Knowledge Worker 2025
// Version: 0.1.0
/////////////

using LoginPI.Engine.ScriptBase;

public class Outlook_Close : ScriptBase
{
    void Execute()
    {
        int globalWaitInSeconds = 3; // Standard wait time between actions
        int timeoutSeconds = 2;      // 2-second timeout for window searches

        Wait(seconds: 2, showOnScreen: true, onScreenText: "Closing Outlook if open.");
        Log("Closing Outlook if open.");
        var mainWindow = FindWindow(processName: "outlook", title: "* - Outlook*", timeout: 5, continueOnError: true);
        if (mainWindow != null)
        {
            mainWindow.Focus();
            mainWindow.Maximize();
            Wait(globalWaitInSeconds);

            // Send the ALT+F4 key to attempt to close Outlook.
            mainWindow.Type("{ALT+F4}", hideInLogging: false);
            Wait(globalWaitInSeconds);

            // Check if the Outlook window still exists.
            var checkWindow = FindWindow(processName: "outlook", title: "* - Outlook*", timeout: timeoutSeconds, continueOnError: true);
            if (checkWindow != null)
            {
                Wait(globalWaitInSeconds);
                checkWindow.Type("{ALT+N}", hideInLogging: false);
                Wait(globalWaitInSeconds);
            }
        }
        else
        {
            Log("Outlook window not found.");
        }
    }
}
