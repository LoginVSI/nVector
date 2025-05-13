// TARGET:winword /e
// START_IN:

/////////////
// Word Close
// Workload: Knowledge Worker 2025
// Version: 0.1.0
/////////////

using LoginPI.Engine.ScriptBase;

public class Word_Close : ScriptBase
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
