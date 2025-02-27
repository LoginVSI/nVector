// TARGET:winword.exe
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

        Wait(seconds:3, showOnScreen:true, onScreenText:"Starting Word");
        Log("Starting Word");
        START(mainWindowTitle: "*Word*", mainWindowClass: "Win32 Window:OpusApp", processName: "WINWORD", timeout: 60);
        SkipFirstRunDialogs();
        MainWindow.Maximize();
        MainWindow.Focus();
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
}