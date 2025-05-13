// TARGET:msedge 
// START_IN:

/////////////
// Edge Start
// Workload: Knowledge Worker 2025
// Version: 0.1.0
/////////////

using LoginPI.Engine.ScriptBase;
using System.Text;
using System.IO;
using System.Diagnostics;

public class Edge_Start : ScriptBase
{
    // =====================================================
    // Configurable Variables
    // =====================================================
    // Browser settings
    string browserExecutable = "msedge.exe";          // Browser executable name
    int tabsToOpen = 10;                              // Number of browser tabs to open
    int waitMessageboxInSeconds = 2;                  // Duration for onscreen wait messages
    int globalWaitInSeconds = 3;                      // Standard wait time between actions

    // Browser launch and initialization timing
    int waitTimeoutInSecondsMsedgeLaunch = 60;         // Maximum wait time (in seconds) for the browser to initially appear
    int waitInSecondsBrowserInitialize = 15;           // Wait time (in seconds) to allow the browser to fully load the defined tabs/URLs

    // =====================================================
    // Execute Method
    // =====================================================
    void Execute()
    {
        Log("Starting browser open process.");
        Wait(seconds: waitMessageboxInSeconds, showOnScreen: true, onScreenText: "Starting browser open process.");

        // =====================================================
        // Setup: Create Directory and Copy PDF
        // =====================================================
        // Get the current user's TEMP folder path.
        string tempPath = GetEnvironmentVariable("TEMP");
        Log("Retrieved TEMP folder: " + tempPath);
        
        // Define the subdirectory path for LoginEnterprise.
        string loginEnterpriseDir = Path.Combine(tempPath, "LoginEnterprise");
        Directory.CreateDirectory(loginEnterpriseDir);
        Log("Ensured directory exists: " + loginEnterpriseDir);

        // Define the destination path for the PDF file.
        string pdfDestination = Path.Combine(loginEnterpriseDir, "loginvsi.pdf");

        // Copy the PDF file from the Login Enterprise appliance to the destination.
        Log("Copying PDF file from KnownFiles.PdfFile to " + pdfDestination);
        CopyFile(KnownFiles.PdfFile, pdfDestination, continueOnError: false, overwrite: true);
        Log("PDF file copied successfully.");

        // =====================================================
        // Build URL List with Hardcoded PDF Path
        // =====================================================
        // Construct the local file URL for the PDF.
        string pdfUrl = "file:///" + pdfDestination.Replace("\\", "/");
        Log("Constructed local PDF URL: " + pdfUrl);
        
        // Build the URL list.
        string urlsDefined = 
            "https://euc.loginvsi.com/customer-portal/knowledge-worker-2023;" +
            "http://distribution.bbb3d.renderfarming.net/video/mp4/bbb_sunflower_1080p_30fps_normal.mp4;" +
            pdfUrl + ";" +
            "https://images.nasa.gov/;" +
            "https://www.google.com/search?q=beautiful+mountains&udm=2;" +            
            "https://www.google.com/search?q=login+vsi&udm=2;" +
            "https://www.bing.com/images/search?q=login%20vsi&lq=0&ghsh=0&ghacc=0&first=1;" +
            "https://www.microsoft.com;";
        Log("URL list constructed.");
        // Good GPU impact site using WebGL: https://webglsamples.org/aquarium/aquarium.html;
        // And one for a high def autoplaying streaming video: http://distribution.bbb3d.renderfarming.net/video/mp4/bbb_sunflower_2160p_30fps_normal.mp4

        // Split the defined URLs into an array using semicolon as the delimiter.
        string[] urlArray = urlsDefined.Split(new char[] { ';' }, System.StringSplitOptions.RemoveEmptyEntries);
        Log("URL array created with " + urlArray.Length + " entries.");

        string firstCommand = browserExecutable + " --guest --no-session-restore";

        // Build the command using the helper method (includes URLs).
        string secondCommand = BuildCommand(urlArray, tabsToOpen);
        Log("Command built: " + secondCommand);

        StartTimer("Browser_Start");
        // Launch the msedge instance.
        ShellExecute(secondCommand, waitForProcessEnd: false, continueOnError: false, forceKillOnExit: false);
        string browserProcessName = Path.GetFileNameWithoutExtension(browserExecutable);
        var browserWindow = FindWindow(
            className: "Win32 Window:Chrome_WidgetWin_1",
            title: "*Microsoft​ Edge",
            processName: browserProcessName,
            timeout: waitTimeoutInSecondsMsedgeLaunch);
        StopTimer("Browser_Start");

        Wait(waitInSecondsBrowserInitialize, onScreenText: "Waiting for browser to fully load tabs");
        Log("Waited " + waitInSecondsBrowserInitialize + " seconds for browser initialization.");

        browserWindow.Maximize();
        browserWindow.Focus();
        Wait(globalWaitInSeconds);
        Log("Browser open process completed.");
    }

    // =====================================================
    // Helper: Build Command String
    // =====================================================
    // Constructs the command string for launching the browser with multiple URLs.
    string BuildCommand(string[] urls, int tabs)
    {
        StringBuilder cmdBuilder = new StringBuilder();
        cmdBuilder.Append(browserExecutable);
        cmdBuilder.Append(" --guest --no-session-restore");
        for (int i = 0; i < tabs; i++)
        {
            string url = urls[i % urls.Length].Trim();
            cmdBuilder.Append(" " + url);
        }
        Log("BuildCommand completed.");
        return cmdBuilder.ToString();
    }
}
