// TARGET:winword /e
// START_IN:

/////////////
// Word Close
// Workload: KnowledgeWorker 2025
// Version: 1.0
/////////////

using LoginPI.Engine.ScriptBase;

public class Close_Word_DefaultScript : ScriptBase
{
    // Global wait time between actions (in seconds). Modify as needed.
    private int globalWaitInSeconds = 3;

    void Execute()
    {
        int closeTimeoutSeconds = 2; // Use a 2-second timeout for find operations in this workload.

        // Close extra windows with titles matching "*loginvsi*", "*edited*", and "*Document*"
        CloseExtraWindow("WINWORD", "*loginvsi*", closeTimeoutSeconds);
        CloseExtraWindow("WINWORD", "*edited*", closeTimeoutSeconds);
        CloseExtraWindow("WINWORD", "*Document*", closeTimeoutSeconds);

        // Handle the specific Microsoft Word confirmation dialog.
        CloseMicrosoftWordDialog();
    }

    /// <summary>
    /// Attempts to close a window matching the title mask (within the specified process) and
    /// handles any confirmation dialogs by sending {ALT+N} if needed.
    /// </summary>
    /// <param name="processName">The process name (e.g., "WINWORD").</param>
    /// <param name="titleMask">Window title mask to search for (e.g., "*loginvsi*").</param>
    /// <param name="timeoutSeconds">Timeout for find operations in seconds.</param>
    void CloseExtraWindow(string processName, string titleMask, int timeoutSeconds)
    {
        int maxAttempts = 1; // Maximum attempts to close the window.
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
            
            // Check if the window still exists (could be due to a confirmation dialog).
            extraWindow = FindWindow(title: titleMask, processName: processName, timeout: timeoutSeconds, continueOnError: true);
            if (extraWindow != null)
            {
                Wait(globalWaitInSeconds);
                extraWindow.Type("{ALT+N}", hideInLogging: false);
                Wait(globalWaitInSeconds);
            }
        }
    }

    // Handling a specific Microsoft Word confirmation dialog: "Do you want to keep the last item you copied?"
    void CloseMicrosoftWordDialog()
    {
        int timeoutSeconds = 2; 
        var msWordWindow = FindWindow(className: "Win32 Window:#32770", title: "Microsoft Word", processName: "WINWORD", timeout: timeoutSeconds, continueOnError: true);
        if (msWordWindow != null)
        {
            msWordWindow.Focus();
            msWordWindow.Maximize();
            Wait(globalWaitInSeconds);
            msWordWindow.Type("{ALT+N}", hideInLogging: false);
            Wait(globalWaitInSeconds);
        }
    }
}
