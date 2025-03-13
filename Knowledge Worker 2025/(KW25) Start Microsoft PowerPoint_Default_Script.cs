// TARGET:powerpnt.exe %temp%\LoginEnterprise\loginvsi.pptx
// START_IN:

/////////////
// PowerPoint Start
// Workload: KnowledgeWorker 2025
// Version: 1.0
/////////////

using LoginPI.Engine.ScriptBase;
using System;
using System.IO;

public class Start_PowerPoint_DefaultScript : ScriptBase
{
    void Execute()
    {
        int globalWaitInSeconds = 3; // Standard wait time between actions
        int waitMessageboxInSeconds = 2; // Duration for onscreen wait messages

        // DeleteTempFiles();
        DownloadPowerPointPresentation();

        Wait(seconds: waitMessageboxInSeconds, showOnScreen: true, onScreenText: "Starting PowerPoint");
        Log("Starting PowerPoint");
        START(mainWindowTitle: "*loginvsi*PowerPoint*", mainWindowClass: "*PPTFrameClass*", timeout: 60);
        Wait(globalWaitInSeconds);
        SkipFirstRunDialogs();
        MainWindow.Maximize();
        MainWindow.Focus();
    }
    
    /*
    private void DeleteTempFiles()
    {
        string pptUnsavedFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft", "Office", "UnsavedFiles");
        string tempFolder = Path.GetTempPath();

        if (Directory.Exists(pptUnsavedFolder))
        {
            foreach (var file in Directory.GetFiles(pptUnsavedFolder, "*.pptx"))
            {
                File.Delete(file);
                Log("Deleted file: " + file);
            }
            foreach (var file in Directory.GetFiles(pptUnsavedFolder, "*.tmp"))
            {
                File.Delete(file);
                Log("Deleted file: " + file);
            }
            foreach (var file in Directory.GetFiles(pptUnsavedFolder, "*.asd"))
            {
                File.Delete(file);
                Log("Deleted file: " + file);
            }
        }
        if (Directory.Exists(tempFolder))
        {
            foreach (var file in Directory.GetFiles(tempFolder, "ppt*.tmp"))
            {
                File.Delete(file);
                Log("Deleted file: " + file);
            }
            /* Commented out because it may delete other important temp files 
            foreach (var file in Directory.GetFiles(tempFolder, "*.tmp"))
        }
    }
    */
    
    private void DownloadPowerPointPresentation()
    {
        int waitMessageboxInSeconds = 2;
        Wait(seconds: waitMessageboxInSeconds, showOnScreen: true, onScreenText: "Downloading PowerPoint presentation file if it doesn't exist");
        var temp = GetEnvironmentVariable("TEMP");
        string loginEnterpriseDir = $"{temp}\\LoginEnterprise";

        if (!Directory.Exists(loginEnterpriseDir))
        {
            Directory.CreateDirectory(loginEnterpriseDir);
            Log("Created directory: " + loginEnterpriseDir);
        }

        string pptxFile = $"{loginEnterpriseDir}\\loginvsi.pptx";
        Wait(waitMessageboxInSeconds);
        Log("Downloading PowerPoint presentation file if it doesn't exist");
        CopyFile(KnownFiles.PowerPointPresentation, pptxFile, overwrite: false, continueOnError: true);
    }
    
    private void SkipFirstRunDialogs()
    {
        int loopCount = 2; // configurable number of loops
        for (int i = 0; i < loopCount; i++)
        {
            var dialog = FindWindow(
                className: "Win32 Window:NUIDialog",
                processName: "POWERPNT",
                continueOnError: true,
                timeout: 3);
            while (dialog != null)
            {
                Wait(seconds: 2, showOnScreen: true, onScreenText: "Closing first run dialog if it exists");
                dialog.Close();
                dialog = FindWindow(
                    className: "Win32 Window:NUIDialog",
                    processName: "POWERPNT",
                    continueOnError: true,
                    timeout: 3);
            }
        }
    }
}
