// TARGET:winword.exe /t
// START_IN:

/////////////
// Office 2019 Prepare
// Workload: KnowledgeWorker 2025
// Version: 1.0
/////////////

using LoginPI.Engine.ScriptBase;
using LoginPI.Engine.ScriptBase.Components;
using System;
using System.IO;
using System.Diagnostics;

public class PrepareOffice2019_DefaultScript : ScriptBase
{
    private int globalWaitInSeconds = 3; // Standard wait time between actions

    /// <summary>
    /// Delete all files in a given folder using provided search patterns.
    /// </summary>
    /// <param name="folderPath">Directory to search for files</param>
    /// <param name="patterns">Array of file search patterns (e.g., "*.asd")</param>
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

    /// <summary>
    /// Deletes temporary files across different folders.
    /// </summary>
    private void DeleteTemporaryFiles()
    {
        string wordFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft", "Word");
        string excelFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft", "Excel");
        string pptUnsavedFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft", "Office", "UnsavedFiles");
        string tempFolder = Path.GetTempPath();
        string tempEnv = GetEnvironmentVariable("TEMP");
        string loginEnterpriseDir = $"{tempEnv}\\LoginEnterprise";

        // Delete files in Word folder
        DeleteFilesWithPatterns(wordFolder, "*.asd", "*.wbk", "*.docx");

        // Delete files in Excel folder
        DeleteFilesWithPatterns(excelFolder, "*.xlsb", "*.xar", "*.xls*", "*.tmp");

        // Delete files in PowerPoint unsaved folder
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
        Log("Deleting all Microsoft Office AutoRecover, backup, 'loginvsi' and 'edited', and temporary files...");
        DeleteTemporaryFiles();

        // =====================================================
        // Launch new blank Word document
        // =====================================================
        ShellExecute("cmd /c start \"\" winword /t", waitForProcessEnd: true, timeout: 3, continueOnError: false);
        /* This is an alternate start blank word document function 
        try
        {
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = "winword.exe",
                Arguments = "/t",
                UseShellExecute = true
            };
            Process.Start(startInfo);
        }
        catch (Exception ex)
        {
            ABORT("Error starting process: " + ex.Message);
        }
        */

        var MainWindow = FindWindow(title: "*Word*", processName: "WINWORD", className: "Win32 Window:OpusApp", continueOnError: false, timeout: 60);
        Wait(globalWaitInSeconds);
        MainWindow.Focus();
        MainWindow.Maximize();
        Wait(globalWaitInSeconds);
        Log("Dismissing first run Word dialogs");

        int loopCount = 2; // configurable number of loops
        for (int i = 0; i < loopCount; i++)
        {
            var openDialog = MainWindow.FindControlWithXPath(
                xPath: "*:NUIDialog",
                timeout: 3,
                continueOnError: true);

            if (openDialog is object)
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

                    openDialog = MainWindow.FindControlWithXPath(
                        xPath: "Pane:NUIDialog",
                        timeout: 5,
                        continueOnError: true);

                    if (openDialog is object)
                    {
                        openDialog.Type("{ALT+i}", hideInLogging: false);
                        Wait(globalWaitInSeconds);
                        openDialog.Type("{ALT+a}", hideInLogging: false);
                    }

                    openDialog = MainWindow.FindControlWithXPath(
                        xPath: "Pane:NUIDialog",
                        timeout: 5,
                        continueOnError: true);

                    if (openDialog is object)
                    {
                        ABORT("Could not close Outlook's First things first dialog");
                    }
                }
                else
                {
                    Wait(globalWaitInSeconds);
                    openDialog.Type("{ESC}");
                }
            }
        }
        Wait(globalWaitInSeconds);
        MainWindow.Close();
        Wait(globalWaitInSeconds);
    }
}
