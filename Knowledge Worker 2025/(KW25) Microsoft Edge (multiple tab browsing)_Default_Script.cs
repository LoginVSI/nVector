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
        Log("Beginning Execute() in MicrosoftEdgeMultipleTabs_DefaultScript.");

        // Simulate user interaction to open the Start Menu.
        Log("Simulating Start Menu interaction.");
        Wait(seconds: 2, showOnScreen: true, onScreenText: "Start Menu");
        Type("{LWIN}");
        Wait(3);
        Type("{LWIN}");
        Wait(1);
        Type("{ESC}");
        Log("Start Menu simulation complete.");
        
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
        for (int i = 0; i < ctrlTabIterations; i++)
        {
            Log($"Iteration {i + 1} started.");            
            Wait(ctrlTabWaitSecondsBeforeScroll); // Wait for the page to load on the current tab.
            
            // Prepare the browser window.
            browserWindow.Maximize();
            browserWindow.MoveMouseToCenter(continueOnError: true, hoverTimeAfterMove: 1);            
            browserWindow.Focus();            

            if (i > 0)
            {
                Log("Switching to next tab with Ctrl+Tab.");
                browserWindow.Type("{ctrl+tab}", hideInLogging: false);                
                Wait(ctrlTabWaitSecondsBeforeScroll); // Extra wait after switching.
                browserWindow.Maximize();
                browserWindow.MoveMouseToCenter(continueOnError: true, hoverTimeAfterMove: 1);
                browserWindow.Focus();
                Log("Switched tab and refocused window.");
            }
            
            // Scroll interactions on the active tab:
            // Usage of Scroll():
            //   - direction: "Down" to scroll down or "Up" to scroll up.
            //   - scrollCount: Number of scroll events to send.
            //   - notches: Number of notches per event (1 notch is typically 120).
            //   - waitTime: Time in seconds to wait between each scroll event.
            // Example:
            //   Scroll("Down", 20, 1, 0.2);
            //   Scroll("Up", 10, 2, 0.3);
            Log("Starting scroll interactions on the active tab.");
            Scroll("Down", 10, 1, 0.2);
            Scroll("Up", 10, 1, 0.2);
            Log("Scroll interactions completed for this iteration.");            
            
            Wait(ctrlTabWaitSecondsAfterScroll); // Wait after scrolling before the next iteration.
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
            Wait(waitTime);
        }
        Log($"Completed scrolling mouse {direction} {scrollCount} times.");
    }
}