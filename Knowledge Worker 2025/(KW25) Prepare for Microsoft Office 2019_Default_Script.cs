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
    private void Execute()
    {
        int globalWaitInSeconds = 3; // Standard wait time between actions

        // Delete all Microsoft Word, Excel, PowerPoint AutoRecover, backup, 'loginvsi' and 'edited', and temporary files
        Log("Deleting all Microsoft Office AutoRecover, backup, 'loginvsi' and 'edited', and temporary files...");

        string wordFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft", "Word");
        string excelFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft", "Excel");
        string pptUnsavedFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft", "Office", "UnsavedFiles");
        string tempFolder = Path.GetTempPath();
        string tempEnv = GetEnvironmentVariable("TEMP");
        string loginEnterpriseDir = $"{tempEnv}\\LoginEnterprise";

        try
        {
            if (Directory.Exists(wordFolder))
            {
                foreach (var file in Directory.GetFiles(wordFolder, "*.asd"))  // AutoRecover files
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

                foreach (var file in Directory.GetFiles(wordFolder, "*.wbk")) // Backup files
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

                foreach (var file in Directory.GetFiles(wordFolder, "*.docx")) // Word documents
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
        catch (Exception ex)
        {
            Log("Error accessing Word directory: " + wordFolder + " - " + ex.Message);
        }

        try
        {
            if (Directory.Exists(excelFolder))
            {
                foreach (var file in Directory.GetFiles(excelFolder, "*.xlsb")) // Excel Binary files
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

                foreach (var file in Directory.GetFiles(excelFolder, "*.xar")) // Excel Archive files
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

                foreach (var file in Directory.GetFiles(excelFolder, "*.xls*")) // Excel documents
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

                foreach (var file in Directory.GetFiles(excelFolder, "*.tmp")) // Excel temp files
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
        catch (Exception ex)
        {
            Log("Error accessing Excel directory: " + excelFolder + " - " + ex.Message);
        }

        try
        {
            if (Directory.Exists(pptUnsavedFolder))
            {
                foreach (var file in Directory.GetFiles(pptUnsavedFolder, "*.pptx")) // PowerPoint files
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

                foreach (var file in Directory.GetFiles(pptUnsavedFolder, "*.tmp")) // PowerPoint temp files
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

                foreach (var file in Directory.GetFiles(pptUnsavedFolder, "*.asd")) // PowerPoint AutoRecover
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
        catch (Exception ex)
        {
            Log("Error accessing PowerPoint directory: " + pptUnsavedFolder + " - " + ex.Message);
        }

        try
        {
            if (Directory.Exists(tempFolder))
            {
                foreach (var file in Directory.GetFiles(tempFolder, "~WRD*.tmp")) // Word temp files
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

                foreach (var file in Directory.GetFiles(tempFolder, "~$*.docx")) // Word temp document cache
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

                foreach (var file in Directory.GetFiles(tempFolder, "~$*.xls*")) // Excel temp document cache
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

                foreach (var file in Directory.GetFiles(tempFolder, "ppt*.tmp")) // PowerPoint temp files
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
        catch (Exception ex)
        {
            Log("Error accessing Temp directory: " + tempFolder + " - " + ex.Message);
        }

        try
        {
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
        catch (Exception ex)
        {
            Log("Error accessing LoginEnterprise directory: " + loginEnterpriseDir + " - " + ex.Message);
        }
        
        /*
        // Delete all Microsoft Word AutoRecover, backup, "loginvsi"/"edited", and temporary files
        Log("Deleting all Microsoft Word AutoRecover, backup, "loginvsi"/"edited", and temporary files...");

        string wordFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft", "Word");
        string tempFolder = Path.GetTempPath();
        string temp = GetEnvironmentVariable("TEMP");
        string tempEnv = GetEnvironmentVariable("TEMP");
        string loginEnterpriseDir = $"{tempEnv}\\LoginEnterprise";

        try
        {
            if (Directory.Exists(wordFolder))
            {
                foreach (var file in Directory.GetFiles(wordFolder, "*.asd"))  // AutoRecover files
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

                foreach (var file in Directory.GetFiles(wordFolder, "*.wbk")) // Backup files
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

                foreach (var file in Directory.GetFiles(wordFolder, "*.docx")) // Word documents
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
        catch (Exception ex)
        {
            Log("Error accessing Word directory: " + wordFolder + " - " + ex.Message);
        }

        try
        {
            if (Directory.Exists(tempFolder))
            {
                foreach (var file in Directory.GetFiles(tempFolder, "~WRD*.tmp")) // Temporary Word files
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

                foreach (var file in Directory.GetFiles(tempFolder, "~$*.docx")) // Temporary Word document cache
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
        catch (Exception ex)
        {
            Log("Error accessing Temp directory: " + tempFolder + " - " + ex.Message);
        }

        try
        {
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
        catch (Exception ex)
        {
            Log("Error accessing LoginEnterprise directory: " + loginEnterpriseDir + " - " + ex.Message);
        }
        */

        // =====================================================
        // Launch new blank Word document using ProcessStartInfo
        // =====================================================
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

        var MainWindow = FindWindow(title:"*Word*", processName:"WINWORD", className:"Win32 Window:OpusApp", continueOnError:false, timeout:60);
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
                        ABORT("Could not close outlooks First things first dialog");
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

        /*
        // Close Word
        Log("Closing Word...");
        try
        {
            foreach (var process in Process.GetProcessesByName("WINWORD"))
            {
                process.Kill();
                process.WaitForExit(); // Ensure the process is terminated
            }
        }
        catch (Exception ex)
        {
            ABORT("Error terminating Word process: " + ex.Message);
        }

        Wait(globalWaitInSeconds);

        // Open and close Excel and PowerPoint as a preparation
        // =====================================================
        // Launch new blank Excel document using ProcessStartInfo
        // =====================================================
        try
        {
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = "excel.exe",
                Arguments = "/s",
                UseShellExecute = true
            };
            Process.Start(startInfo);
        }
        catch (Exception ex)
        {
            ABORT("Error starting Excel process: " + ex.Message);
        }

        // =====================================================
        // Launch new blank PowerPoint presentation using ProcessStartInfo
        // =====================================================
        try
        {
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = "powerpnt.exe",
                Arguments = "/n",
                UseShellExecute = true
            };
            Process.Start(startInfo);
        }
        catch (Exception ex)
        {
            ABORT("Error starting PowerPoint process: " + ex.Message);
        }
        FindWindow(title:"*Excel*", processName:"EXCEL", className:"Win32 Window:XLMAIN", continueOnError:false, timeout:60);
        FindWindow(title:"*PowerPoint*", processName:"POWERPNT", className:"Win32 Window:PPTFrameClass", continueOnError:false, timeout:60);
        Wait(globalWaitInSeconds);

        // Close Excel and PowerPoint
        Log("Closing Excel and PowerPoint...");
        string[] processesToKill = { "EXCEL", "POWERPNT" };

        try
        {
            foreach (var processName in processesToKill)
            {
                foreach (var process in Process.GetProcessesByName(processName))
                {
                    process.Kill();
                    process.WaitForExit(); // Ensure the process is terminated
                    Log($"Terminated process: {processName}");
                }
            }
        }
        catch (Exception ex)
        {
            ABORT("Error terminating Office processes: " + ex.Message);
        }
        */

        Wait(globalWaitInSeconds);
    }
}
