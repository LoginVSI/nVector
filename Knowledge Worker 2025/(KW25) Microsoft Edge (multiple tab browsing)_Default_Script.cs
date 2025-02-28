// TARGET:msedge
// START_IN:

/////////////
// Browser Application
// Workload: KnowledgeWorker 2025
// Version: 1.0
/////////////

using LoginPI.Engine.ScriptBase;
using System;
using System.Runtime.InteropServices;

public class Browser_MultipleTabs_DefaultScript : ScriptBase
{
    // =====================================================
    // Import and Constants
    // =====================================================
    [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
    public static extern void mouse_event(uint dwFlags, uint dx, uint dy, int dwData, UIntPtr dwExtraInfo);
    public const uint MOUSEEVENTF_WHEEL = 0x0800; // Constant for a mouse wheel event

    // =====================================================
    // Configurable Variables
    // =====================================================
    // Global timings and iterations
    int ctrlTabIterations = 5;                     // Number of iterations for tab switching and scrolling interactions
    int ctrlTabWaitSecondsBeforeScroll = 3;         // Wait time before scrolling to allow the page to load
    int ctrlTabWaitSecondsAfterScroll = 1;          // Wait time after scrolling before next iteration
    string browserProcessName = "msedge";           // Process name for Microsoft Edge

    // Scrolling parameters
    int scrollDownCount = 15;                       // Number of scroll events for scrolling down
    int scrollUpCount = 15;                         // Number of scroll events for scrolling up
    double scrollWaitTime = 0.2;                    // Wait time between each scroll event

    // Additional global wait times
    int globalWaitInSeconds = 3;                    // Standard wait time between actions
    int waitMessageboxInSeconds = 2;                // Duration for onscreen wait messages
    int startMenuWaitInSeconds = 5;                // Duration for Start Menu wait between interactions

    private void Execute()
    {
        // =====================================================
        // Simulate Start Menu Interaction
        // =====================================================
        Log("Simulating Start Menu interaction.");
        Wait(startMenuWaitInSeconds);
        Type("{LWIN}",hideInLogging:false);
        Wait(seconds: startMenuWaitInSeconds, showOnScreen: true, onScreenText: "Opening Start Menu");
        Type("{LWIN}",hideInLogging:false);
        Wait(seconds: 1, showOnScreen: true, onScreenText: "Closing Start Menu");
        Type("{ESC}",hideInLogging:false);
        Log("Start Menu simulation complete.");
        Wait(startMenuWaitInSeconds);

        // =====================================================
        // Bring Browser Window into Focus
        // =====================================================
        var browserWindow = FindWindow(
            className: "Win32 Window:Chrome_WidgetWin_1",
            title: "*Microsoft​ Edge",
            processName: browserProcessName);
        browserWindow.Minimize();
        Wait(seconds: globalWaitInSeconds, showOnScreen: true, onScreenText: "Minimizing Browser");
        browserWindow.Maximize();
        browserWindow.Focus();
        Wait(seconds: globalWaitInSeconds, showOnScreen: true, onScreenText: "Focusing Browser");

        // =====================================================
        // Setup Iteration Message and Wait Time
        // =====================================================
        int totalWaitPerIteration = ctrlTabWaitSecondsBeforeScroll + ctrlTabWaitSecondsAfterScroll;
        int totalCtrlTabTime = ctrlTabIterations * totalWaitPerIteration;
        string message = $"Performing {ctrlTabIterations} iterations with {ctrlTabWaitSecondsBeforeScroll} sec wait after scrolling and {ctrlTabWaitSecondsAfterScroll} sec wait after scrolling. Total wait time: {totalCtrlTabTime} sec.";
        Wait(seconds: waitMessageboxInSeconds, showOnScreen: true, onScreenText: message);
        Log(message);

        // =====================================================
        // Iterate Over Tabs with Scrolling Interactions
        // =====================================================
        for (int i = 0; i < ctrlTabIterations; i++)
        {
            Log($"Iteration {i + 1} started.");
            Wait(seconds: ctrlTabWaitSecondsBeforeScroll);

            // Ensure browser window is maximized and in focus
            browserWindow.Maximize();
            browserWindow.MoveMouseToCenter(continueOnError: true, hoverTimeAfterMove: 1);
            browserWindow.Focus();
            browserWindow.Maximize();

            if (i > 0)
            {
                Log("Switching to next tab with Ctrl+Tab.");
                browserWindow.Type("{ctrl+tab}", hideInLogging: false);
                browserWindow.Type("{f5}",hideInLogging:false);
                Wait(seconds: ctrlTabWaitSecondsBeforeScroll);
                browserWindow.Maximize();
                browserWindow.MoveMouseToCenter(continueOnError: true, hoverTimeAfterMove: 1);
                browserWindow.Focus();
                Log("Switched tab and refocused window.");
            }
            
            // =====================================================
            // Helper: Scroll Function
            // =====================================================
            // Usage of Scroll():
            //   - direction: "Down" to scroll down or "Up" to scroll up.
            //   - scrollCount: Number of scroll events to send.
            //   - notches: Number of notches per event (1 notch is typically 120).
            //   - waitTime: Time in seconds to wait between each scroll event.
            // Example:
            //   Scroll("Down", 20, 1, 0.2);
            //   Scroll("Up", 10, 2, 0.3);
            // =====================================================
            // Scroll Interactions on Active Tab
            // =====================================================
            Log("Starting scroll interactions on the active tab.");
            Scroll("Down", scrollDownCount, 1, scrollWaitTime);
            Scroll("Up", scrollUpCount, 1, scrollWaitTime);
            Log("Scroll interactions completed for this iteration.");

            Wait(seconds: ctrlTabWaitSecondsAfterScroll, showOnScreen: true, onScreenText: "Waiting after scrolling");
            Log($"Iteration {i + 1} complete. Waiting {ctrlTabWaitSecondsAfterScroll} seconds before next iteration.");
        }
        Log("All iterations completed.");
    }

    void Scroll(string direction, int scrollCount, int notches, double waitTime)
    {
        if (waitTime <= 0)
        {
            throw new ArgumentException("Scroll waitTime must be greater than 0 seconds.");
        }

        int sign = direction.Equals("Down", StringComparison.OrdinalIgnoreCase) ? -1 : 1;
        int delta = sign * 120 * notches;

        Log($"Scrolling mouse {direction} {scrollCount} times, {notches} notch(es) per scroll, with {waitTime} sec between each scroll.");
        for (int i = 0; i < scrollCount; i++)
        {
            mouse_event(MOUSEEVENTF_WHEEL, 0, 0, delta, UIntPtr.Zero);
            Wait(seconds: waitTime);
        }
        Log($"Completed scrolling mouse {direction} {scrollCount} times.");
    }
}
