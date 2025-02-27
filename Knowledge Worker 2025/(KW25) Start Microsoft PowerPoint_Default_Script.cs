// TARGET:powerpnt.exe
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
        // Delete all Microsoft PowerPoint AutoRecover, backup, and temporary files
        Log("Deleting all Microsoft PowerPoint AutoRecover, backup, and temporary files...");
        
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
            /* Commented out becasue it may delete other important temp files 
            foreach (var file in Directory.GetFiles(tempFolder, "*.tmp"))
            {
                File.Delete(file);
                Log("Deleted file: " + file);
            } */
        }

        Wait(seconds:2, showOnScreen:true, onScreenText:"Starting PowerPoint");
        Log("Starting PowerPoint");
        START(mainWindowTitle:"*PowerPoint*", mainWindowClass:"*PPTFrameClass*", timeout:60);
        Wait(3);
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
                processName: "POWERPNT", 
                continueOnError: true, 
                timeout: 5);
            while (dialog != null)
            {
                dialog.Close();
                dialog = FindWindow(
                    className: "Win32 Window:NUIDialog", 
                    processName: "POWERPNT", 
                    continueOnError: true, 
                    timeout: 5);
            }
        }
    }
}