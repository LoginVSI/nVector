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
    private void Execute()
    {           
        int globalWaitInSeconds = 3; // Standard wait time between actions
        
        // Delete all Microsoft Word AutoRecover, backup, and temporary files
        Log("Deleting all Microsoft Word AutoRecover, backup, *loginvsi*/*edited*, and temporary files...");

        string wordFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft", "Word");
        string tempFolder = Path.GetTempPath();
        string temp = GetEnvironmentVariable("TEMP");
        string loginEnterpriseDir = $"{temp}\\LoginEnterprise";

        if (Directory.Exists(wordFolder))
        {
            foreach (var file in Directory.GetFiles(wordFolder, "*.asd"))
            {
                File.Delete(file);
                Log("Deleted file: " + file);
            }
            foreach (var file in Directory.GetFiles(wordFolder, "*.wbk"))
            {
                File.Delete(file);
                Log("Deleted file: " + file);
            }
            foreach (var file in Directory.GetFiles(wordFolder, "*.docx"))
            {
                File.Delete(file);
                Log("Deleted file: " + file);
            }
        }

        if (Directory.Exists(tempFolder))
        {
            foreach (var file in Directory.GetFiles(tempFolder, "~WRD*.tmp"))
            {
                File.Delete(file);
                Log("Deleted file: " + file);
            }
            foreach (var file in Directory.GetFiles(tempFolder, "~$*.docx"))
            {
                File.Delete(file);
                Log("Deleted file: " + file);
            }
            /* Commented out because it may delete other important temp files 
            foreach (var file in Directory.GetFiles(tempFolder, "*.tmp"))
            {
                File.Delete(file);
                Log("Deleted file: " + file);
            } */
        }

        // Delete all "loginvsi*" and "edited*" files in LoginEnterprise directory
        if (Directory.Exists(loginEnterpriseDir))
        {
            foreach (var file in Directory.GetFiles(loginEnterpriseDir))
            {
                string fileName = Path.GetFileName(file).ToLower();
                if (fileName.Contains("loginvsi") || fileName.Contains("edited"))
                {
                    File.Delete(file);
                    Log("Deleted file: " + file);
                }
            }
        }

        // Set registry values; this should be a run-once preparation
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
        // RegImport(create_regfile(@"HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\PowerPoint\Options", @"DisableAutoRecover", @"dword:00000001")); // Confirms disabling of AutoRecover in PowerPoint.
        // RegImport(create_regfile(@"HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Excel\Options", @"DisableAutoRecover", @"dword:00000001")); // Confirms disabling of AutoRecover in Excel.
        // RegImport(create_regfile(@"HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Word\Options", @"DisableAutoRecover", @"dword:00000001")); // Confirms disabling of AutoRecover in Word.
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

        // RegImport(create_regfile(@"HKEY_CURRENT_USER\software\Policies\Microsoft\Edge", @"RestoreOnStartup", @"dword:00000000")); // Would configure Microsoft Edge to not restore tabs on startup.
        // RegImport(create_regfile(@"HKEY_CURRENT_USER\software\Policies\Microsoft\Edge", @"HideRestoreDialogEnabled", @"dword:00000001")); // Would hide the restore dialog in Microsoft Edge.

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

        Wait(globalWaitInSeconds);
        var MainWindow = FindWindow(title:"*Word*", processName:"WINWORD", className:"Win32 Window:OpusApp", continueOnError:false, timeout:60);
        Wait(globalWaitInSeconds);
        MainWindow.Focus();
        MainWindow.Maximize();
        Wait(globalWaitInSeconds);
        SkipFirstRunDialogs();
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

    private void SkipFirstRunDialogs()
    {
        int globalWaitInSeconds = 3; // Standard wait time between actions
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
}