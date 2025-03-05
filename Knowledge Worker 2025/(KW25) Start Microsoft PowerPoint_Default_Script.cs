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

        DeleteTempFiles();
        DownloadPowerPointPresentation();

        Wait(seconds: waitMessageboxInSeconds, showOnScreen: true, onScreenText: "Starting PowerPoint");
        Log("Starting PowerPoint");
        START(mainWindowTitle: "*loginvsi*PowerPoint*", mainWindowClass: "*PPTFrameClass*", timeout: 60);
        Wait(globalWaitInSeconds);
        SkipFirstRunDialogs();
        MainWindow.Maximize();
        MainWindow.Focus();
    }
    
    private void DeleteTempFiles()
    {
        // Define relevant folders
        string pptFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft", "PowerPoint");
        string unsavedFilesFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft", "Office", "UnsavedFiles");
        string tempFolder = Path.GetTempPath();

        // Delete from PowerPoint AutoRecover folder
        if (Directory.Exists(pptFolder))
        {
            foreach (var file in Directory.GetFiles(pptFolder, "*.pptx")) // Standard PowerPoint presentations
            {
                File.Delete(file);
                Log("Deleted file: " + file);
            }
            foreach (var file in Directory.GetFiles(pptFolder, "*.pptm")) // Macro-enabled PowerPoint files
            {
                File.Delete(file);
                Log("Deleted file: " + file);
            }
            foreach (var file in Directory.GetFiles(pptFolder, "*.ppsx")) // PowerPoint Show files
            {
                File.Delete(file);
                Log("Deleted file: " + file);
            }
            foreach (var file in Directory.GetFiles(pptFolder, "*.ppsm")) // PowerPoint Show macro-enabled
            {
                File.Delete(file);
                Log("Deleted file: " + file);
            }
            foreach (var file in Directory.GetFiles(pptFolder, "*.tmp")) // Temporary PowerPoint files
            {
                File.Delete(file);
                Log("Deleted file: " + file);
            }
            foreach (var file in Directory.GetFiles(pptFolder, "*.asd")) // AutoRecover files (rare for PowerPoint)
            {
                File.Delete(file);
                Log("Deleted file: " + file);
            }
            foreach (var file in Directory.GetFiles(pptFolder, "~$*.ppt*")) // PowerPoint lock files
            {
                File.Delete(file);
                Log("Deleted file: " + file);
            }
        }

        // Delete from Office Unsaved Files folder (NEW ADDITION)
        if (Directory.Exists(unsavedFilesFolder))
        {
            foreach (var file in Directory.GetFiles(unsavedFilesFolder, "*.ppt*")) // Unsaved PowerPoint files
            {
                File.Delete(file);
                Log("Deleted file: " + file);
            }
        }

        // Delete from Temp folder
        if (Directory.Exists(tempFolder))
        {
            foreach (var file in Directory.GetFiles(tempFolder, "~$*.ppt*")) // PowerPoint lock files
            {
                File.Delete(file);
                Log("Deleted file: " + file);
            }
            foreach (var file in Directory.GetFiles(tempFolder, "ppt*.tmp")) // PowerPoint-related temp files
            {
                File.Delete(file);
                Log("Deleted file: " + file);
            }
        }
    }
    
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
        string editedPptxFile = $"{loginEnterpriseDir}\\edited.pptx";

        if (File.Exists(pptxFile))
        {
            File.Delete(pptxFile);
            Log("Deleted existing file: " + pptxFile);
        }

        if (File.Exists(editedPptxFile))
        {
            File.Delete(editedPptxFile);
            Log("Deleted existing file: " + editedPptxFile);
        }

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
