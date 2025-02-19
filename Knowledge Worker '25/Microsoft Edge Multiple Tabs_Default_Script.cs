// TARGET:ping
// START_IN:
using LoginPI.Engine.ScriptBase;
using System.Text;
using System.IO;
using System.Diagnostics;

public class BrowserMultipleTabs : ScriptBase
{
    // Developed with and tested in 24H2 Win 11 x64 Pro and msedge v133.0.3065.51 x64

    // User-defined variable for the browser executable.
    // Change this if you need to use a different browser.
    string browserExecutable = "msedge.exe";
    
    // Optional: Terminate all running instances of the browser before launching new ones.
    // Set to true to terminate all existing instances before launching new tabs.
    bool terminateExistingProcesses = true;
    
    // Optional: Terminate all running instances of the browser at the end of the script.
    // Set to true to terminate all existing instances after execution is complete.
    // Default is false, so the browser remains open.
    bool terminateProcessesAtEnd = false;
    
    // Number of browser tabs (sites) to open when launching the browser.
    // Set this value to the desired number of tabs.
    int tabsToOpen = 5;

    // List of URLs to open.
    // Define your sites as a semicolon-delimited string.
    // For example: "https://www.google.com;https://www.microsoft.com;https://www.bing.com"
    // If tabsToOpen is greater than the number of URLs defined here, the URLs will cycle.
    // Good GPU impact site using WebGL: https://webglsamples.org/aquarium/aquarium.html;
    string urlsDefined = "https://euc.loginvsi.com/customer-portal/knowledge-worker-2023;http://distribution.bbb3d.renderfarming.net/video/mp4/bbb_sunflower_2160p_30fps_normal.mp4;https://www.google.com;https://www.microsoft.com;https://www.bing.com;";
    
    // Maximum wait time in seconds for the browser to initially appear.
    int waitTimeoutInSecondsMsedgeLaunch = 30;
    
    // Wait time in seconds to allow the browser to fully load the defined tabs/URLs.
    int waitInSecondsBrowserInitialize = 30;

    // --- New Variables for Iterative Ctrl+Tab ---
    // Number of times to send the Ctrl+Tab command.
    int ctrlTabIterations = 15;
    
    // Wait time in seconds between each Ctrl+Tab command.
    int ctrlTabWaitSeconds = 20;

    void Execute() 
    {
        // If requested, terminate all running instances of the browser before proceeding.
        if (terminateExistingProcesses)
        {
            // Compute the process name by stripping the ".exe" from the browserExecutable.
            string procName = Path.GetFileNameWithoutExtension(browserExecutable);
            Log("Terminating any running instances of " + procName + "...");
            
            Process[] runningProcesses = Process.GetProcessesByName(procName);
            foreach (Process proc in runningProcesses)
            {
                try 
                {
                    proc.Kill();
                    Log("Terminated process ID: " + proc.Id);
                }
                catch (System.Exception ex)
                {
                    Log("Error terminating process ID: " + proc.Id + " - " + ex.Message);
                }
            }
            // Wait one second for processes to terminate.
            Wait(1, onScreenText:"Waiting for processes to terminate...");
            
            // Ensure no processes are running (wait up to 10 seconds).
            int waitCounter = 0;
            while (Process.GetProcessesByName(procName).Length > 0 && waitCounter < 10)
            {
                Wait(1);
                waitCounter++;
            }
        }

        // Split the defined URLs into an array using semicolon as the delimiter.
        string[] urlArray = urlsDefined.Split(new char[] { ';' }, System.StringSplitOptions.RemoveEmptyEntries);

        // Build the command line string starting with the browser executable.
        // Insert the "--guest" flag immediately after the executable name.
        StringBuilder cmdBuilder = new StringBuilder();
        cmdBuilder.Append(browserExecutable);
        cmdBuilder.Append(" --guest");

        // Append each URL (cycling through the array if necessary) to the command string.
        // For example, if tabsToOpen is 5 and there are 2 URLs, it will result in:
        // msedge.exe --guest https://google.com https://microsoft.com https://google.com https://microsoft.com https://google.com
        for (int i = 0; i < tabsToOpen; i++)
        {
            // Use modulo (%) to cycle through the URL array if needed.
            string url = urlArray[i % urlArray.Length].Trim();
            cmdBuilder.Append(" " + url);
        }

        // Convert the StringBuilder to a string to be used as the command.
        string command = cmdBuilder.ToString();

        // Execute the browser with the dynamically constructed command line.
        // The options: waitForProcessEnd:false, continueOnError:false, forceKillOnExit:false
        ShellExecute(command, waitForProcessEnd:false, continueOnError:false, forceKillOnExit:false);

        // Compute the process name by stripping the ".exe" from the browserExecutable.
        string browserProcessName = Path.GetFileNameWithoutExtension(browserExecutable);

        // Find the browser window (using your existing metafunction; adjust className/title as needed).
        var browserWindow = FindWindow(
            className: "Win32 Window:Chrome_WidgetWin_1", 
            title: "*Microsoft​ Edge", 
            processName: browserProcessName, 
            timeout: waitTimeoutInSecondsMsedgeLaunch);
        
        // Wait for the browser to initialize fully.
        Wait(waitInSecondsBrowserInitialize, onScreenText:"Waiting for browser to fully load tabs");
        browserWindow.Maximize();
        
        // --- Begin iterative Ctrl+Tab code block ---
        // Calculate total time for all iterations (iterations * wait time).
        int totalCtrlTabTime = ctrlTabIterations * ctrlTabWaitSeconds;
        Wait(3, onScreenText:"Performing " + ctrlTabIterations + " iterations of Ctrl+Tab with " + ctrlTabWaitSeconds + " second(s) wait between each iteration. Total time: " + totalCtrlTabTime + " second(s).");
        Log("Performing " + ctrlTabIterations + " iterations of Ctrl+Tab with " + ctrlTabWaitSeconds + " second(s) wait between each iteration. Total time: " + totalCtrlTabTime + " second(s).");

        // Iterate, sending Ctrl+Tab and then waiting the defined interval.
        for (int i = 0; i < ctrlTabIterations; i++)
        {
            browserWindow.Focus(); // Make sure the browser window is in focus.
            browserWindow.Type("{ctrl+tab}", hideInLogging:false);
            Wait(ctrlTabWaitSeconds);
        }
        // --- End iterative Ctrl+Tab code block ---

        // Optionally terminate browser processes at the end of the script.
        if (terminateProcessesAtEnd)
        {
            Log("Terminating any running instances of " + browserProcessName + " at the end of the script...");
            Process[] runningProcesses = Process.GetProcessesByName(browserProcessName);
            foreach (Process proc in runningProcesses)
            {
                try 
                {
                    proc.Kill();
                    Log("Terminated process ID: " + proc.Id);
                }
                catch (System.Exception ex)
                {
                    Log("Error terminating process ID: " + proc.Id + " - " + ex.Message);
                }
            }
            // Wait one second for processes to terminate.
            Wait(1, onScreenText:"Waiting for processes to terminate at end...");
            
            // Ensure no processes are running (wait up to 10 seconds).
            int waitCounter = 0;
            while (Process.GetProcessesByName(browserProcessName).Length > 0 && waitCounter < 10)
            {
                Wait(1);
                waitCounter++;
            }
        }
    }
}
