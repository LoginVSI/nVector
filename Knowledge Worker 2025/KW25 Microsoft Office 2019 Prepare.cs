// TARGET:winword.exe /t
// START_IN:

/////////////
// Office 2019 Prepare
// Workload: Knowledge Worker 2025
// Version: 0.1.0
/////////////

using LoginPI.Engine.ScriptBase;
using LoginPI.Engine.ScriptBase.Components;
using System;
using System.IO;
using System.Diagnostics;

public class Office_2019_Prepare : ScriptBase
{
    private int globalWaitInSeconds = 3; // Standard wait time between actions

    private void DeleteFilesWithPatterns(string folderPath, params string[] patterns)
    {
        if (Directory.Exists(folderPath))
        {
            foreach (var pattern in patterns)
            {
                foreach (var file in Directory.GetFiles(folderPath, pattern))
                {
                    try
                    {
                        File.Delete(file);
                        Log("Deleted file: " + file);
                    }
                    catch (Exception ex)
                    {
                        Log("Failed to delete file: " + file + " - " + ex.Message);
                    }
                }
            }
        }
    }

    private void DeleteTemporaryFiles()
    {
        string wordFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft", "Word");
        string excelFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft", "Excel");
        string pptUnsavedFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft", "Office", "UnsavedFiles");
        string tempFolder = Path.GetTempPath();
        string tempEnv = GetEnvironmentVariable("TEMP");
        string loginEnterpriseDir = $"{tempEnv}\\LoginEnterprise";

        // Delete files in Word folder (AutoRecover, Backup, and Word documents)
        DeleteFilesWithPatterns(wordFolder, "*.asd", "*.wbk", "*.docx");

        // Delete files in Excel folder (Excel Binary, Archive, Documents, and Temp files)
        DeleteFilesWithPatterns(excelFolder, "*.xlsb", "*.xar", "*.xls*", "*.tmp");

        // Delete files in PowerPoint unsaved folder (PowerPoint, Temp and AutoRecover files)
        DeleteFilesWithPatterns(pptUnsavedFolder, "*.pptx", "*.tmp", "*.asd");

        // Delete files in Temp folder for Word, Excel, and PowerPoint specific patterns
        DeleteFilesWithPatterns(tempFolder, "~WRD*.tmp", "~$*.docx", "~$*.xls*", "ppt*.tmp");

        // Delete files in LoginEnterprise directory that contain "loginvsi" or "edited" in the filename
        if (Directory.Exists(loginEnterpriseDir))
        {
            foreach (var file in Directory.GetFiles(loginEnterpriseDir))
            {
                try
                {
                    string fileName = Path.GetFileName(file).ToLower();
                    if (fileName.Contains("loginvsi") || fileName.Contains("edited"))
                    {
                        File.Delete(file);
                        Log("Deleted file: " + file);
                    }
                }
                catch (Exception ex)
                {
                    Log("Failed to delete file: " + file + " - " + ex.Message);
                }
            }
        }
    }

    private void Execute()
    {
        // =====================================================
        // Pre-delete: Remove all Microsoft Office AutoRecover, backup,
        // 'loginvsi' and 'edited', and temporary files.
        // =====================================================
        // Log("Deleting all Microsoft Office AutoRecover, backup, 'loginvsi' and 'edited', and temporary files...");

        // To delete temp files in the temp Word, Excel, and PowerPoint folders, then uncomment the following line:
        // DeleteTemporaryFiles();

        // =====================================================
        // Launch new blank Word document
        // =====================================================
        try
        {
            ShellExecute("winword /t", waitForProcessEnd: false, timeout: 60, continueOnError: true, forceKillOnExit: false);
            /* Alternate start blank word document function:
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = "winword.exe",
                Arguments = "/t",
                UseShellExecute = true
            };
            Process.Start(startInfo); */
        }
        catch (Exception ex)
        {
            ABORT("Error starting process: " + ex.Message);
        }

        Wait(globalWaitInSeconds);
        var mainWindow = FindWindow(title:"*Document*Word*", processName:"WINWORD", continueOnError:false, timeout:60);
        Wait(globalWaitInSeconds);
        mainWindow.Focus();
        mainWindow.Maximize();
        Wait(globalWaitInSeconds);
        
        // Dismiss first run dialogs using detailed logic (passing the mainWord window)
        DismissFirstRunDialogs(mainWindow);
        Wait(globalWaitInSeconds);
        
        // =====================================================
        // Close Word Windows
        // =====================================================
        int closeTimeoutSeconds = 2;
        CloseExtraWindow("WINWORD", "*loginvsi*", closeTimeoutSeconds);
        CloseExtraWindow("WINWORD", "*edited*", closeTimeoutSeconds);
        CloseExtraWindow("WINWORD", "*Document*", closeTimeoutSeconds);
    }

    /// <summary>
    /// Dismisses first run dialogs for Word using detailed logic.
    /// </summary>
    /// <param name="mainWindow">The main Word window.</param>
    private void DismissFirstRunDialogs(IWindow mainWindow)
    {
        Log("Dismissing first run Word dialogs");

        int loopCount = 2; // configurable number of loops
        for (int i = 0; i < loopCount; i++)
        {
            var openDialog = mainWindow.FindControlWithXPath(
                xPath: "*:NUIDialog",
                timeout: 3,
                continueOnError: true);

            if (openDialog != null)
            {
                if (openDialog.GetTitle().StartsWith("First things", StringComparison.CurrentCultureIgnoreCase))
                {
                    Wait(seconds: 2, showOnScreen: true, onScreenText: "Closing first things first dialog if it exists");

                    openDialog.FindControl(
                        className: "RadioButton:NetUIRadioButton",
                        title: "Install updates only",
                        continueOnError: true)?.Click();

                    openDialog.FindControl(
                        className: "Button:NetUIButton",
                        title: "Accept",
                        continueOnError: true)?.Click();

                    openDialog = mainWindow.FindControlWithXPath(
                        xPath: "Pane:NUIDialog",
                        timeout: 5,
                        continueOnError: true);

                    if (openDialog != null)
                    {
                        openDialog.Type("{ALT+i}", hideInLogging: false);
                        Wait(globalWaitInSeconds);
                        openDialog.Type("{ALT+a}", hideInLogging: false);
                    }

                    openDialog = mainWindow.FindControlWithXPath(
                        xPath: "Pane:NUIDialog",
                        timeout: 5,
                        continueOnError: true);

                    if (openDialog != null)
                    {
                        ABORT("Could not close first things first dialog");
                    }
                }
                else
                {
                    Wait(globalWaitInSeconds);
                    openDialog.Type("{ESC}");
                }
            }
        }
    }

    /// <summary>
    /// Attempts to close a window matching the title mask (within the specified process) and
    /// handles any confirmation dialogs by sending {ALT+N} if needed.
    /// </summary>
    /// <param name="processName">The process name (e.g., "WINWORD").</param>
    /// <param name="titleMask">Window title mask to search for (e.g., "*loginvsi*").</param>
    /// <param name="timeoutSeconds">Timeout for find operations in seconds.</param>
    private void CloseExtraWindow(string processName, string titleMask, int timeoutSeconds)
    {
        int maxAttempts = 1; // Maximum number of attempts to close the window.

        for (int attempt = 0; attempt < maxAttempts; attempt++)
        {
            var extraWindow = FindWindow(title: titleMask, processName: processName, timeout: 2, continueOnError: true);
            if (extraWindow == null)
            {
                // Window is already closed.
                break;
            }

            Wait(globalWaitInSeconds);
            extraWindow.Focus();
            extraWindow.Maximize();
            Wait(globalWaitInSeconds);
            extraWindow.Type("{ESC}", hideInLogging: false);
            Wait(globalWaitInSeconds);
            extraWindow.Type("{ALT+F4}", hideInLogging: false);
            Wait(globalWaitInSeconds);

            // Check if the window still exists
            var checkWindow = FindWindow(title: titleMask, processName: processName, timeout: timeoutSeconds, continueOnError: true);
            if (checkWindow != null)
            {
                Wait(globalWaitInSeconds);
                checkWindow.Type("{ALT+N}", hideInLogging: false);
                Wait(globalWaitInSeconds);
            }
        }
    }
}
