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

public class MicrosoftEdgeMultipleTabs_DefaultScript : ScriptBase
{
    int ctrlTabIterations = 10; // Number of iterations for scrolling interactions
    int ctrlTabWaitSecondsBeforeScroll = 3; // Wait time (in seconds) before scrolling to allow the page to load
    int ctrlTabWaitSecondsAfterScroll = 1;  // Wait time (in seconds) after scrolling before next iteration
    string browserProcessName = "msedge"; 

    // Import the user32.dll function to simulate mouse events.
    [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
    public static extern void mouse_event(uint dwFlags, uint dx, uint dy, int dwData, UIntPtr dwExtraInfo);

    public const uint MOUSEEVENTF_WHEEL = 0x0800; // Constant for a mouse wheel event

    private void Execute()
    {
        // Find the browser window (adjust className/title as needed).
        var browserWindow = FindWindow(
            className: "Win32 Window:Chrome_WidgetWin_1",
            title: "*Microsoft​ Edge",
            processName: browserProcessName);

        // Calculate total wait time per iteration and overall.
        int totalWaitPerIteration = ctrlTabWaitSecondsBeforeScroll + ctrlTabWaitSecondsAfterScroll;
        int totalCtrlTabTime = ctrlTabIterations * totalWaitPerIteration;
        string message = $"Performing {ctrlTabIterations} iterations with {ctrlTabWaitSecondsBeforeScroll} sec wait before scrolling and {ctrlTabWaitSecondsAfterScroll} sec wait after scrolling. Total wait time: {totalCtrlTabTime} sec.";
        Wait(3, showOnScreen: true, onScreenText: message);
        Log(message);

        // For each iteration:
        //   - On the first iteration, scroll the initial tab without switching.
        //   - On subsequent iterations, send Ctrl+Tab to switch to the next tab before scrolling.
        for (int i = 0; i < ctrlTabIterations; i++)
        {
            Wait(ctrlTabWaitSecondsBeforeScroll);
            browserWindow.Maximize();
            browserWindow.MoveMouseToCenter(continueOnError: true, hoverTimeAfterMove: 1);
            browserWindow.Focus();

            if (i > 0)
            {
                // For iterations beyond the first, switch to the next tab.
                browserWindow.Type("{ctrl+tab}", hideInLogging: false);
                Wait(ctrlTabWaitSecondsBeforeScroll);
                browserWindow.Maximize();
                browserWindow.MoveMouseToCenter(continueOnError: true, hoverTimeAfterMove: 1);
                browserWindow.Focus();
            }
            
            // Usage of Scroll():
            //   - direction: "Down" to scroll down or "Up" to scroll up.
            //   - scrollCount: Number of scroll events to send.
            //   - notches: Number of notches per event (1 notch is typically 120).
            //   - waitTime: Time in seconds to wait between each scroll event.
            // Example:
            //   Scroll("Down", 20, 1, 0.2);
            //   Scroll("Up", 10, 2, 0.3);

            // Scroll interactions on the active tab after switching:
            Scroll("Down", 10, 1, 0.2);
            Scroll("Up", 5, 2, 0.3);
            
            Wait(ctrlTabWaitSecondsAfterScroll);
        }
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
            Wait(waitTime);
        }
    }
}
