// TARGET:excel.exe
// START_IN:

/////////////
// Excel Start
// Workload: KnowledgeWorker 2025
// Version: 1.0
/////////////

using LoginPI.Engine.ScriptBase;
using System;
using System.IO;

public class Start_Excel_DefaultScript : ScriptBase
{
    const string ProcessName = "EXCEL";
    void Execute()
    {      
        // Delete all Microsoft Excel AutoRecover, backup, and temporary files
        Log("Deleting all Microsoft Excel AutoRecover, backup, and temporary files...");

        string excelFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft", "Excel");
        string tempFolder = Path.GetTempPath();

        if (Directory.Exists(excelFolder))
        {
            foreach (var file in Directory.GetFiles(excelFolder, "*.xlsb"))
            {
                File.Delete(file);
                Log("Deleted file: " + file);
            }
            foreach (var file in Directory.GetFiles(excelFolder, "*.xar"))
            {
                File.Delete(file);
                Log("Deleted file: " + file);
            }
            foreach (var file in Directory.GetFiles(excelFolder, "*.xls*"))
            {
                File.Delete(file);
                Log("Deleted file: " + file);
            }
            foreach (var file in Directory.GetFiles(excelFolder, "*.tmp"))
            {
                File.Delete(file);
                Log("Deleted file: " + file);
            }
        }

        if (Directory.Exists(tempFolder))
        {
            foreach (var file in Directory.GetFiles(tempFolder, "~$*.xls*"))
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
          
        Wait(seconds:2, showOnScreen:true, onScreenText:"Starting Excel");
        Log("Starting Excel");
        START(mainWindowTitle: "*Excel*", mainWindowClass: "*XLMAIN*", timeout: 60);
        Wait(3);
        SkipFirstRunDialogs();
        MainWindow.Maximize();
        MainWindow.Focus();
    }
    void SkipFirstRunDialogs()
    {
        int loopCount = 2; // configurable number of loops
        for (int i = 0; i < loopCount; i++)
        {
            var dialog = FindWindow(
                className: "Win32 Window:NUIDialog", 
                processName: "EXCEL", 
                continueOnError: true, 
                timeout: 5);
            while (dialog != null)
            {
                dialog.Close();
                dialog = FindWindow(
                    className: "Win32 Window:NUIDialog", 
                    processName: "EXCEL", 
                    continueOnError: true, 
                    timeout: 5);
            }
        }
    }
}