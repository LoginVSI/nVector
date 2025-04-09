// TARGET:excel /e
// START_IN:

/////////////
// Excel Close
// Workload: KnowledgeWorker 2025
// Version: 1.0
/////////////

using LoginPI.Engine.ScriptBase;

public class Close_Excel_DefaultScript : ScriptBase
{
    void Execute()
    {
        int closeTimeoutSeconds = 2; // Use a 2-second timeout for find operations in this workload.
        
        // Close extra windows with titles matching "*loginvsi*", "*edited*", and "*Book*"
        CloseExtraWindow("EXCEL", "*loginvsi*", closeTimeoutSeconds);
        CloseExtraWindow("EXCEL", "*edited*", closeTimeoutSeconds);
        CloseExtraWindow("EXCEL", "*Book*", closeTimeoutSeconds);
    }

    /// <summary>
    /// Attempts to close a window matching the title mask (within the specified process) and
    /// handles any confirmation dialogs by sending {ALT+N} if needed.
    /// </summary>
    /// <param name="processName">The process name (e.g., "EXCEL").</param>
    /// <param name="titleMask">Window title mask to search for (e.g., "*loginvsi*").</param>
    /// <param name="timeoutSeconds">Timeout for find operations in seconds.</param>
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