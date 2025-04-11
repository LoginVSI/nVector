// TARGET:winword.exe /t
// START_IN:

/////////////
// M365 Prepare
// Workload: KnowledgeWorker 2025
// Version: 1.0
/////////////

using LoginPI.Engine.ScriptBase;
using System;
using System.IO;
using System.Diagnostics;

public class M365PrivacyPrep_DefaultScript : ScriptBase
{
    private int globalWaitInSeconds = 3; // Standard wait time between actions

    /// <summary>
    /// Deletes all files in a given directory matching each provided pattern.
    /// </summary>
    /// <param name="folderPath">Directory to check for files</param>
    /// <param name="patterns">One or more search patterns (e.g., "*.asd")</param>
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
    /// Consolidates deletion of all Office temporary and backup files into one method.
    /// </summary>
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
        DeleteFilesWithPatterns(excelFolder, "*.xlsb", "*.xar", "*.tmp", "*.xls*");

        // Delete files in PowerPoint unsaved folder (PowerPoint, Temp and AutoRecover files)
        DeleteFilesWithPatterns(pptUnsavedFolder, "*.tmp", "*.asd", "*.pptx");

        // Delete files in Temp folder (Word temp files, Excel and PowerPoint temp document caches, and PowerPoint temp files)
        DeleteFilesWithPatterns(tempFolder, "~WRD*.tmp", "~$*.docx", "~$*.xls*", "ppt*.tmp");

        // Delete files in LoginEnterprise directory that contain "loginvsi" or "edited"
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
        // DeleteTemporaryFiles();

        // =====================================================
        // Set registry values; this should be a run-once preparation
        // =====================================================
        Wait(seconds:2, showOnScreen:true, onScreenText:"Setting Reg Values");
        RegImport(create_regfile(@"HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\General", @"ShownFirstRunOptin", @"dword:00000001")); // Marks that the first run opt-in dialog has been shown.
        RegImport(create_regfile(@"HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\Licensing", @"DisableActivationUI", @"dword:00000001")); // Disables the Office activation UI prompt.
        RegImport(create_regfile(@"HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Registration", @"AcceptAllEulas", @"dword:00000001")); // Automatically accepts all End User License Agreements (EULAs).

        RegImport(create_regfile(@"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\excel\Security\ProtectedView", @"DisableAttachmentsInPV", @"dword:00000001")); // Disables attachments in Protected View for Excel.
        RegImport(create_regfile(@"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\excel\Security\ProtectedView", @"DisableInternetFilesInPV", @"dword:00000001")); // Disables files from internet sources in Protected View for Excel.
        RegImport(create_regfile(@"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\excel\Security\ProtectedView", @"DisableUnsafeLocationsInPV", @"dword:00000001")); // Disables files from unsafe locations in Protected View for Excel.
        RegImport(create_regfile(@"HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\excel\Options", @"SaveAutoRecoverInfoEvery", @"dword:00000000")); // Disables the automatic saving of AutoRecover information in Excel.
        RegImport(create_regfile(@"HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Excel\Resiliency", @"DisableAutoRecover", @"dword:00000001")); // Disables the AutoRecover feature in Excel.

        RegImport(create_regfile(@"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Word\Security\ProtectedView", @"DisableAttachmentsInPV", @"dword:00000001")); // Disables attachments in Protected View for Word.
        RegImport(create_regfile(@"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Word\Security\ProtectedView", @"DisableInternetFilesInPV", @"dword:00000001")); // Disables files from internet sources in Protected View for Word.
        RegImport(create_regfile(@"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Word\Security\ProtectedView", @"DisableUnsafeLocationsInPV", @"dword:00000001")); // Disables files from unsafe locations in Protected View for Word.

        RegImport(create_regfile(@"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Powerpoint\Security\ProtectedView", @"DisableAttachmentsInPV", @"dword:00000001")); // Disables attachments in Protected View for PowerPoint.
        RegImport(create_regfile(@"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Powerpoint\Security\ProtectedView", @"DisableInternetFilesInPV", @"dword:00000001")); // Disables files from internet sources in Protected View for PowerPoint.
        RegImport(create_regfile(@"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Powerpoint\Security\ProtectedView", @"DisableUnsafeLocationsInPV", @"dword:00000001")); // Disables files from unsafe locations in Protected View for PowerPoint.

        RegImport(create_regfile(@"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Powerpoint\options", @"DisableHardwareNotification", @"dword:00000001")); // Disables hardware acceleration notifications in PowerPoint.

        RegImport(create_regfile(@"HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Powerpoint\Options", @"SaveAutoRecoverInfoEvery", @"dword:00000000")); // Disables AutoRecover information saving in PowerPoint.
        RegImport(create_regfile(@"HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\PowerPoint", @"AutoRecover", @"dword:00000000")); // Disables the AutoRecover feature in PowerPoint.
        
        RegImport(create_regfile(@"HKEY_CURRENT_USER\software\microsoft\office\16.0\common\sharepointintegration", @"hidelearnmorelink", @"dword:00000001")); // Hides the "Learn More" link for SharePoint integration.
        RegImport(create_regfile(@"HKEY_CURRENT_USER\software\microsoft\office\16.0\common\graphics", @"disablehardwareacceleration", @"dword:00000001")); // Disables hardware acceleration for Office graphics.
        RegImport(create_regfile(@"HKEY_CURRENT_USER\software\microsoft\office\16.0\common\graphics", @"disableanimations", @"dword:00000001")); // Disables animations within Office applications.
        RegImport(create_regfile(@"HKEY_CURRENT_USER\software\microsoft\office\16.0\common\general", @"skydrivesigninoption", @"dword:00000000")); // Disables the SkyDrive (OneDrive) sign-in option in Office.
        RegImport(create_regfile(@"HKEY_CURRENT_USER\software\microsoft\office\16.0\common\general", @"disableboottoofficestart", @"dword:00000001")); // Disables booting directly to the Office start screen.
        RegImport(create_regfile(@"HKEY_CURRENT_USER\software\microsoft\office\16.0\firstrun", @"disablemovie", @"dword:00000001")); // Disables the first run introductory movie/video.
        RegImport(create_regfile(@"HKEY_CURRENT_USER\software\microsoft\office\16.0\firstrun", @"bootedrtm", @"dword:00000001")); // Marks Office as having completed the initial first run experience.
        RegImport(create_regfile(@"HKEY_CURRENT_USER\software\microsoft\office\16.0\excel\options", @"defaultformat", @"dword:00000051")); // Sets the default file format for Excel (value 0x51 likely corresponds to XLSX).
        RegImport(create_regfile(@"HKEY_CURRENT_USER\software\microsoft\office\16.0\powerpoint\options", @"defaultformat", @"dword:00000027")); // Sets the default file format for PowerPoint (value 0x27 likely corresponds to PPTX).
        RegImport(create_regfile(@"HKEY_CURRENT_USER\software\microsoft\office\16.0\word\options", @"defaultformat", @"")); // Leaves the default file format for Word unchanged or resets it to system default.
        RegImport(create_regfile(@"HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Word\Options", @"SaveAutoRecoverInfoEvery", @"dword:00000000")); // Disables AutoRecover information saving in Word.
        RegImport(create_regfile(@"HKEY_CURRENT_USER\software\microsoft\office\16.0\common\options", @"PrivacyNoticeShown", @"dword:00000002")); // Indicates that the Privacy Notice has been shown or dismissed.
        RegImport(create_regfile(@"HKEY_CURRENT_USER\software\microsoft\office\16.0\common\PromoDialogShown", @"FluentWelcomeDialogShown", @"dword:00000001")); // Marks the Fluent Welcome dialog as already shown.

        RegImport(create_regfile(@"HKEY_CURRENT_USER\software\microsoft\office\16.0\Outlook\Preferences", @"ReopenWindowsOption", @"dword:00000001")); // Enables the option to reopen previous windows in Outlook.

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
        var MainWindow = FindWindow(title:"*Document*Word*", processName:"WINWORD", continueOnError:false, timeout:60);
        Wait(globalWaitInSeconds);
        MainWindow.Focus();
        MainWindow.Maximize();
        Wait(globalWaitInSeconds);
        SkipFirstRunDialogs();
        Wait(globalWaitInSeconds);
        
        // =====================================================
        // Close Word Windows
        // =====================================================
        int closeTimeoutSeconds = 2;
        CloseExtraWindow("WINWORD", "*loginvsi*", closeTimeoutSeconds);
        CloseExtraWindow("WINWORD", "*edited*", closeTimeoutSeconds);
        CloseExtraWindow("WINWORD", "*Document*", closeTimeoutSeconds);
    }

    private void SkipFirstRunDialogs()
    {
        int loopCount = 2; // configurable number of loops
        for (int i = 0; i < loopCount; i++)
        {
            var dialog = FindWindow(
                className: "Win32 Window:NUIDialog",
                processName: "WINWORD",
                continueOnError: true,
                timeout: 3);
            while (dialog != null)
            {
                Wait(globalWaitInSeconds);
                dialog.Close();
                dialog = FindWindow(
                    className: "Win32 Window:NUIDialog",
                    processName: "WINWORD",
                    continueOnError: true,
                    timeout: 3);
            }
        }
    } 

    private string create_regfile(string key, string value, string data)
    {            
        System.Text.StringBuilder sb = new System.Text.StringBuilder();
        var file = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "reg.reg");

        sb.AppendLine("Windows Registry Editor Version 5.00");
        sb.AppendLine();
        sb.AppendLine($"[{key}]");
        if(data.ToLower().Contains("dword"))
        {
            sb.AppendLine($"\"{value}\"={data.ToLower()}");
        }
        else
        {
            sb.AppendLine($"\"{value}\"=\"{data}\"");
        }
        sb.AppendLine();

        System.IO.File.WriteAllText(file, sb.ToString());

        return file;
    }

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
