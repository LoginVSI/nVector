// TARGET:msedge 
// START_IN:

/////////////
// Browser Start
// Workload: KnowledgeWorker 2025
// Version: 1.0
/////////////

using LoginPI.Engine.ScriptBase;
using System.Text;
using System.IO;
using System.Diagnostics;

public class Start_Browser_DefaultScript : ScriptBase
{
    string browserExecutable = "msedge.exe"; // Browser executable name
    int tabsToOpen = 10; // Number of browser tabs to open

    // Maximum wait time in seconds for the browser to initially appear.
    int waitTimeoutInSecondsMsedgeLaunch = 30;
    
    // Wait time in seconds to allow the browser to fully load the defined tabs/URLs.
    int waitInSecondsBrowserInitialize = 5;

    void Execute()
    {
        Log("Starting browser open process.");

        // Get the current user's TEMP folder path.
        string tempPath = GetEnvironmentVariable("TEMP");
        
        // Define the subdirectory path for LoginEnterprise.
        string loginEnterpriseDir = Path.Combine(tempPath, "LoginEnterprise");
        // Create the directory if it doesn't exist.
        Directory.CreateDirectory(loginEnterpriseDir);

        // Define the destination path for the PDF file.
        string pdfDestination = Path.Combine(loginEnterpriseDir, "loginvsi.pdf");

        // Copy the PDF file from the Login Enterprise appliance to the destination.
        Log("Copying PDF file from KnownFiles.PdfFile to " + pdfDestination);
        CopyFile(KnownFiles.PdfFile, pdfDestination, continueOnError: false, overwrite: true);
        Log("PDF file copied successfully.");

        // Build the URL list with the local PDF file path as the second URL.
        string urlsDefined =
            "https://euc.loginvsi.com/customer-portal/knowledge-worker-2023;" +
            pdfDestination + ";" +
            "https://images.nasa.gov/;" +
            "https://www.google.com/search?q=beautiful+mountains&udm=2;" +            
            "https://www.bing.com/images/search?q=beautiful%20mountains&first=1;" +
            "https://www.google.com/search?q=nvidia&udm=2;" +
            "https://www.bing.com/images/search?q=nvidia&lq=0&ghsh=0&ghacc=0&first=1;" +
            "https://www.google.com/search?q=login+vsi&udm=2;" +
            "https://www.bing.com/images/search?q=login%20vsi&lq=0&ghsh=0&ghacc=0&first=1;" +
            "https://www.microsoft.com;";

        // Split the defined URLs into an array using semicolon as the delimiter.
        string[] urlArray = urlsDefined.Split(new char[] { ';' }, System.StringSplitOptions.RemoveEmptyEntries);
        
        // Build the command using the helper method.
        string command = BuildCommand(urlArray, tabsToOpen);

        StartTimer("Browser_Start"); // Start timer for opening the browser.

        // Execute the browser with the dynamically constructed command line.
        ShellExecute(command, waitForProcessEnd: false, continueOnError: false, forceKillOnExit: false);

        string browserProcessName = System.IO.Path.GetFileNameWithoutExtension(browserExecutable);

        // Find the browser window (adjust className/title as needed).
        FindWindow(
            className: "Win32 Window:Chrome_WidgetWin_1",
            title: "*Microsoft​ Edge",
            processName: browserProcessName,
            timeout: waitTimeoutInSecondsMsedgeLaunch);

        StopTimer("Browser_Start"); // Stop timer after the browser window is found.

        // Wait for the browser to fully load the tabs.
        Wait(waitInSecondsBrowserInitialize, onScreenText: "Waiting for browser to fully load tabs");
        Log("Waited " + waitInSecondsBrowserInitialize + " seconds for browser initialization.");

        // Maximize and focus the browser window.
        var browserWindow = FindWindow(
            className: "Win32 Window:Chrome_WidgetWin_1",
            title: "*Microsoft​ Edge",
            processName: browserProcessName);
        browserWindow.Maximize();        
        browserWindow.Focus();        

        Log("Browser open process completed.");
    }

    // Helper method to build the command string.
    string BuildCommand(string[] urls, int tabs)
    {
        StringBuilder cmdBuilder = new StringBuilder();
        cmdBuilder.Append(browserExecutable);
        cmdBuilder.Append(" --guest");
        for (int i = 0; i < tabs; i++)
        {
            string url = urls[i % urls.Length].Trim();
            cmdBuilder.Append(" " + url);
        }
        return cmdBuilder.ToString();
    }
}
