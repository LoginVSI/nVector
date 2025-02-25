// TARGET:outlook.exe /importprf %TEMP%\LoginEnterprise\Outlook.prf
// START_IN:

/////////////
// Outlook Application
// Workload: KnowledgeWorker 2025
// Version: 1.0
/////////////

using LoginPI.Engine.ScriptBase;
using System.IO;
using System;

public class M365Outlook_InvocationScript : ScriptBase
{
    private void Execute()
    {
        // This script downloads the configuration (PRF) and data (PST) files, prepares them for Outlook, and starts Outlook.
        
        // Retrieve the TEMP environment variable.
        var temp = GetEnvironmentVariable("TEMP");
        
        // Define the target directory.
        string targetDir = $"{temp}\\LoginEnterprise";
        // Create the "Login Enterprise" directory if it doesn't exist.
        if (!Directory.Exists(targetDir))
        {
            Directory.CreateDirectory(targetDir);
        }
        
        // Download the PRF and PST files from the appliance.
        // Overwrite existing files to ensure a clean start.
        Wait(seconds:1, showOnScreen:true, onScreenText:"Get PRF & PST");
        Log("Downloading PRF & PST files");
        CopyFile(KnownFiles.OutlookConfiguration, $"{targetDir}\\Outlook.prf", overwrite:true, continueOnError:true);
        CopyFile(KnownFiles.OutlookData, $"{targetDir}\\Outlook.pst", overwrite:true, continueOnError:true);
        
        // Update the PRF file: replace the placeholder %TEMP% with the actual TEMP path.
        string prfPath = $"{targetDir}\\Outlook.prf";
        File.WriteAllText(prfPath, File.ReadAllText(prfPath).Replace("%TEMP%", $"{temp}"));
        
        // Start Outlook using its command line configuration.
        Wait(seconds:2, showOnScreen:true, onScreenText:"Starting Outlook");
        Log("Starting Outlook");
        START(mainWindowTitle:"Inbox*", mainWindowClass:"Win32 Window:rctrl_renwnd32", processName:"OUTLOOK", timeout:60, continueOnError:true);
        Wait(5);
        MainWindow.Maximize();
        MainWindow.Focus();
    }
}