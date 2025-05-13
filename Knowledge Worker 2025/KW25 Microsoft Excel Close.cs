// TARGET:excel /e
// START_IN:

/////////////
// Excel Close
// Workload: Knowledge Worker 2025
// Version: 0.1.0
/////////////

using LoginPI.Engine.ScriptBase;

public class Excel_Close : ScriptBase
{
    void Execute()
    {
        int closeTimeoutSeconds = 2; // Use a 2-second timeout for find operations in this workload.
        
        // Close extra windows with titles matching "*loginvsi*", "*edited*", and "*Book*"
        CloseExtraWindow("EXCEL", "*loginvsi*", closeTimeoutSeconds);
        CloseExtraWindow("EXCEL", "*edited*", closeTimeoutSeconds);
        CloseExtraWindow("EXCEL", "*Book*", closeTimeoutSeconds);
    }

    void CloseExtraWindow(string processName, string titleMask, int timeoutSeconds)
    {
        int globalWaitInSeconds = 3; // Standard wait time between actions.
        int maxAttempts = 1;         // Maximum attempts to close the window.
        for (int attempt = 0; attempt < maxAttempts; attempt++)
        {
            var extraWindow = FindWindow(title: titleMask, processName: processName, timeout: timeoutSeconds, continueOnError: true);
            if (extraWindow == null)
            {
                // The window is already closed.
                break;
            }
            
            // Attempt to close the window.
            Wait(globalWaitInSeconds);
            extraWindow.Focus();
            extraWindow.Maximize();
            Wait(globalWaitInSeconds);
            extraWindow.Type("{ESC}", hideInLogging: false);
            Wait(globalWaitInSeconds);
            extraWindow.Type("{ALT+F4}", hideInLogging: false);
            Wait(globalWaitInSeconds);
            
            // Check if the window still exists (could be due to a modified document confirmation popup).
            extraWindow = FindWindow(title: titleMask, processName: processName, timeout: timeoutSeconds, continueOnError: true);
            if (extraWindow != null)
            {
                // Wait a bit and then dismiss the confirmation with {ALT+N}.
                Wait(globalWaitInSeconds);
                extraWindow.Type("{ALT+N}", hideInLogging: false);
                Wait(globalWaitInSeconds);
            }
        }
    }
}