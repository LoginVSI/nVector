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

public class M365PrivacyPrep_DefaultScript : ScriptBase
{
    private void Execute()
    {           
        // Delete all Microsoft Word AutoRecover, backup, and temporary files
        Log("Deleting all Microsoft Word AutoRecover, backup, and temporary files...");

        string wordFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft", "Word");
        string tempFolder = Path.GetTempPath();

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
            /* Commented out becasue it may delete other important temp files 
            foreach (var file in Directory.GetFiles(tempFolder, "*.tmp"))
            {
                File.Delete(file);
                Log("Deleted file: " + file);
            } */
        }

        // Define environement variables to use with Workload
        var temp = GetEnvironmentVariable("TEMP");

        // Set registry values; this should be a run-once preparation
        Wait(seconds:2, showOnScreen:true, onScreenText:"Setting Reg Values #1");
        RegImport(create_regfile(@"HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\General",@"ShownFirstRunOptin",@"dword:00000001"));
        RegImport(create_regfile(@"HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\Licensing",@"DisableActivationUI",@"dword:00000001"));
        RegImport(create_regfile(@"HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Registration",@"AcceptAllEulas",@"dword:00000001"));
        
        RegImport(create_regfile(@"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\excel\Security\ProtectedView",@"DisableAttachmentsInPV",@"dword:00000001"));
        RegImport(create_regfile(@"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\excel\Security\ProtectedView",@"DisableInternetFilesInPV",@"dword:00000001"));
        RegImport(create_regfile(@"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\excel\Security\ProtectedView",@"DisableUnsafeLocationsInPV",@"dword:00000001"));
        RegImport(create_regfile(@"HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\excel\Options", @"SaveAutoRecoverInfoEvery", @"dword:00000000"));
        RegImport(create_regfile(@"HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Excel\Resiliency", @"DisableAutoRecover", @"dword:00000001"));
        
        RegImport(create_regfile(@"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Word\Security\ProtectedView",@"DisableAttachmentsInPV",@"dword:00000001"));
        RegImport(create_regfile(@"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Word\Security\ProtectedView",@"DisableInternetFilesInPV",@"dword:00000001"));
        RegImport(create_regfile(@"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Word\Security\ProtectedView",@"DisableUnsafeLocationsInPV",@"dword:00000001"));
        
        
        RegImport(create_regfile(@"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Powerpoint\Security\ProtectedView",@"DisableAttachmentsInPV",@"dword:00000001"));
        RegImport(create_regfile(@"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Powerpoint\Security\ProtectedView",@"DisableInternetFilesInPV",@"dword:00000001"));
        RegImport(create_regfile(@"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Powerpoint\Security\ProtectedView",@"DisableUnsafeLocationsInPV",@"dword:00000001"));
        
        RegImport(create_regfile(@"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Powerpoint\options", @"DisableHardwareNotification",@"dword:00000001"));

        RegImport(create_regfile(@"HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Powerpoint\Options", @"SaveAutoRecoverInfoEvery", @"dword:00000000"));
        RegImport(create_regfile(@"HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\PowerPoint", @"AutoRecover", @"dword:00000000"));
        

        RegImport(create_regfile(@"HKEY_CURRENT_USER\software\microsoft\office\16.0\common\sharepointintegration", @"hidelearnmorelink", @"dword:00000001"));
        RegImport(create_regfile(@"HKEY_CURRENT_USER\software\microsoft\office\16.0\common\graphics", @"disablehardwareacceleration", @"dword:00000001"));
        RegImport(create_regfile(@"HKEY_CURRENT_USER\software\microsoft\office\16.0\common\graphics", @"disableanimations", @"dword:00000001"));
        RegImport(create_regfile(@"HKEY_CURRENT_USER\software\microsoft\office\16.0\common\general",@"skydrivesigninoption", @"dword:00000000"));
        RegImport(create_regfile(@"HKEY_CURRENT_USER\software\microsoft\office\16.0\common\general", @"disableboottoofficestart", @"dword:00000001"));
        RegImport(create_regfile(@"HKEY_CURRENT_USER\software\microsoft\office\16.0\firstrun", @"disablemovie", @"dword:00000001"));
        RegImport(create_regfile(@"HKEY_CURRENT_USER\software\microsoft\office\16.0\firstrun", @"bootedrtm", @"dword:00000001"));
        RegImport(create_regfile(@"HKEY_CURRENT_USER\software\microsoft\office\16.0\excel\options", @"defaultformat", @"dword:00000051"));
        RegImport(create_regfile(@"HKEY_CURRENT_USER\software\microsoft\office\16.0\powerpoint\options", @"defaultformat", @"dword:00000027"));
        RegImport(create_regfile(@"HKEY_CURRENT_USER\software\microsoft\office\16.0\word\options", @"defaultformat",@""));
        RegImport(create_regfile(@"HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Word\Options", @"SaveAutoRecoverInfoEvery", @"dword:00000000"));
        RegImport(create_regfile(@"HKEY_CURRENT_USER\software\microsoft\office\16.0\common\options", @"PrivacyNoticeShown", @"dword:00000002"));

        // RegImport(create_regfile(@"HKEY_CURRENT_USER\software\Policies\Microsoft\Edge", @"RestoreOnStartup", @"dword:00000000"));
        // RegImport(create_regfile(@"HKEY_CURRENT_USER\software\Policies\Microsoft\Edge", @"HideRestoreDialogEnabled", @"dword:00000001"));
        
        // Start Application
        Log("Starting Word");
        Wait(seconds:2, showOnScreen:true, onScreenText:"Starting Word; finding first run dialogs, if any, then stopping App");
        START(mainWindowTitle:"*Word*", processName:"WINWORD", timeout:600);
        Wait(1);
        FindWindow(className : "Win32 Window:OpusApp", title : "*Word*", processName : "WINWORD", continueOnError:true).Focus();
        FindWindow(className : "Win32 Window:OpusApp", title : "*Word*", processName : "WINWORD", continueOnError:true).Maximize();      
        SkipFirstRunDialogs();        

        STOP();
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
                timeout: 5);
            while (dialog != null)
            {
                dialog.Close();
                dialog = FindWindow(
                    className: "Win32 Window:NUIDialog",
                    processName: "WINWORD",
                    continueOnError: true,
                    timeout: 5);
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