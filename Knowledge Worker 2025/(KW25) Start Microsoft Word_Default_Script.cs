// TARGET:winword.exe %temp%\LoginEnterprise\loginvsi.docx
// START_IN:

/////////////
// Word Start
// Workload: KnowledgeWorker 2025
// Version: 1.0
/////////////

using LoginPI.Engine.ScriptBase;
using System;
using System.IO;

public class Start_Word_DefaultScript : ScriptBase
{
    void Execute()
    {
        int globalWaitInSeconds = 3; // Standard wait time between actions
        int waitMessageboxInSeconds = 2; // Duration for onscreen wait messages

        DeleteTempFiles();
        DownloadWordDocument();

        // Running Word
        Wait(seconds: waitMessageboxInSeconds, showOnScreen: true, onScreenText: "Starting Word");
        Log("Starting Word");
        START(mainWindowTitle: "*loginvsi*Word*", mainWindowClass: "Win32 Window:OpusApp", processName: "WINWORD", timeout: 60);
        Wait(globalWaitInSeconds);
        SkipFirstRunDialogs();
        MainWindow.Maximize();
        MainWindow.Focus();
        Wait(globalWaitInSeconds);
    }
    
    private void DeleteTempFiles()
    {
        // Define relevant folders
        string wordFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft", "Word");
        string unsavedFilesFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft", "Office", "UnsavedFiles");
        string tempFolder = Path.GetTempPath();

        // Delete from Word AutoRecover folder
        if (Directory.Exists(wordFolder))
        {
            foreach (var file in Directory.GetFiles(wordFolder, "*.asd")) // Word AutoRecover files
            {
                File.Delete(file);
                Log("Deleted file: " + file);
            }
            foreach (var file in Directory.GetFiles(wordFolder, "*.wbk")) // Word backup files
            {
                File.Delete(file);
                Log("Deleted file: " + file);
            }
            foreach (var file in Directory.GetFiles(wordFolder, "*.docx")) // Word documents
            {
                File.Delete(file);
                Log("Deleted file: " + file);
            }
            foreach (var file in Directory.GetFiles(wordFolder, "*.docm")) // Macro-enabled Word documents
            {
                File.Delete(file);
                Log("Deleted file: " + file);
            }
            foreach (var file in Directory.GetFiles(wordFolder, "*.dotx")) // Word templates
            {
                File.Delete(file);
                Log("Deleted file: " + file);
            }
            foreach (var file in Directory.GetFiles(wordFolder, "*.dotm")) // Macro-enabled Word templates
            {
                File.Delete(file);
                Log("Deleted file: " + file);
            }
        }

        // Delete from Office Unsaved Files folder (NEW ADDITION)
        if (Directory.Exists(unsavedFilesFolder))
        {
            foreach (var file in Directory.GetFiles(unsavedFilesFolder, "*.doc*")) // Unsaved Word documents
            {
                File.Delete(file);
                Log("Deleted file: " + file);
            }
        }

        // Delete from Temp folder
        if (Directory.Exists(tempFolder))
        {
            foreach (var file in Directory.GetFiles(tempFolder, "~WRD*.tmp")) // Word temp files
            {
                File.Delete(file);
                Log("Deleted file: " + file);
            }
            foreach (var file in Directory.GetFiles(tempFolder, "~$*.doc*")) // Word lock files
            {
                File.Delete(file);
                Log("Deleted file: " + file);
            }
            foreach (var file in Directory.GetFiles(tempFolder, "Word*.tmp")) // More Word-related temp files
            {
                File.Delete(file);
                Log("Deleted file: " + file);
            }
        }
    }
    
    private void DownloadWordDocument()
    {
        int waitMessageboxInSeconds = 2;
        Wait(seconds: waitMessageboxInSeconds, showOnScreen: true, onScreenText: "Downloading Word document file if it doesn't exist");
        var temp = GetEnvironmentVariable("TEMP");
        string loginEnterpriseDir = $"{temp}\\LoginEnterprise";

        if (!Directory.Exists(loginEnterpriseDir))
        {
            Directory.CreateDirectory(loginEnterpriseDir);
            Log("Created directory: " + loginEnterpriseDir);
        }

        string docxFile = $"{loginEnterpriseDir}\\loginvsi.docx";
        string editedDocxFile = $"{loginEnterpriseDir}\\edited.docx";

        if (File.Exists(docxFile))
        {
            File.Delete(docxFile);
            Log("Deleted existing file: " + docxFile);
        }

        if (File.Exists(editedDocxFile))
        {
            File.Delete(editedDocxFile);
            Log("Deleted existing file: " + editedDocxFile);
        }

        Log("Downloading Word document file if it doesn't exist");
        CopyFile(KnownFiles.WordDocument, docxFile, overwrite: false, continueOnError: true);
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
                Wait(seconds: 2, showOnScreen: true, onScreenText: "Closing first run dialog if it exists");
                dialog.Close();
                dialog = FindWindow(
                    className: "Win32 Window:NUIDialog",
                    processName: "WINWORD",
                    continueOnError: true,
                    timeout: 3);
            }
        }
    }
}
