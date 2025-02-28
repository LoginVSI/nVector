// TARGET:outlook.exe /importprf %TEMP%\LoginPI\Outlook.prf
// START_IN:

/////////////
// Outlook Start
// Workload: KnowledgeWorker 2025
// Version: 1.0
/////////////

using LoginPI.Engine.ScriptBase;
using System.IO;
using System;

public class M365Outlook_InvocationScript : ScriptBase
{
    // =====================================================
    // Configurable Variables
    // =====================================================
    int globalWaitInSeconds = 3;                    // Wait time between actions
    int waitMessageboxInSeconds = 2;                // Duration for onscreen wait messages

    // =====================================================
    // Execute Method
    // =====================================================
    private void Execute()
    {
        // =====================================================
        // Setup: Create Directory and Prepare Files
        // =====================================================
        Log("Preparing Outlook configuration.");
        // Retrieve the TEMP environment variable.
        var temp = GetEnvironmentVariable("TEMP");

        // Define the target directory.
        string targetDir = $"{temp}\\LoginPI";
        // Create the "Login PI" directory if it doesn't exist.
        if (!Directory.Exists(targetDir))
        {
            Directory.CreateDirectory(targetDir);
            Log("Created directory: " + targetDir);
        }

        // =====================================================
        // Download PRF and PST Files
        // =====================================================
        string prfFile = $"{targetDir}\\Outlook.prf";
        string pstFile = $"{targetDir}\\Outlook.pst";
        Wait(seconds: waitMessageboxInSeconds, showOnScreen: true, onScreenText: "Get PRF & PST");
        Log("Downloading PRF & PST files.");
        CopyFile(KnownFiles.OutlookConfiguration, prfFile, continueOnError: true);
        CopyFile(KnownFiles.OutlookData, pstFile, continueOnError: true);

        // =====================================================
        // Update PRF File
        // =====================================================
        // Replace the placeholder %TEMP% with the actual TEMP path.
        string prfContent = File.ReadAllText(prfFile).Replace("%TEMP%", temp);
        File.WriteAllText(prfFile, prfContent);
        Log("Updated PRF file with current TEMP path.");

        // =====================================================
        // Launch Outlook
        // =====================================================
        Wait(seconds: waitMessageboxInSeconds, showOnScreen: true, onScreenText: "Starting Outlook");
        Log("Starting Outlook.");
        START(mainWindowTitle: "Inbox*", mainWindowClass: "Win32 Window:rctrl_renwnd32", processName: "OUTLOOK", timeout: 60, continueOnError: true);
        Wait(globalWaitInSeconds);
        MainWindow.Maximize();
        MainWindow.Focus();

        // =====================================================
        // Dismiss First Run Dialogs
        // =====================================================
        SkipFirstRunDialogs();
        MainWindow.Maximize();
        MainWindow.Focus();
        Log("Outlook is now ready.");
    }

    // =====================================================
    // Helper: Skip First-Run Dialogs
    // =====================================================
    private void SkipFirstRunDialogs()
    {
        int firstRunRetryCount = 4;
        for (int i = 0; i < firstRunRetryCount; i++)
        {
            var signinWindow = MainWindow.FindControlWithXPath(
                "Win32 Window:NUIDialog", 
                timeout: 5, 
                continueOnError: true);
            if (signinWindow != null)
            {
                signinWindow.Type("{ESC}", hideInLogging: false);
                Log("Dismissed a first-run dialog.");
            }
        }
    }
}
