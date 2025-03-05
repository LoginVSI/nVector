// TARGET:excel.exe %temp%\LoginEnterprise\loginvsi.xlsx
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
        int globalWaitInSeconds = 3; // Standard wait time between actions
        int waitMessageboxInSeconds = 2; // Duration for onscreen wait messages

        DeleteTempFiles();
        DownloadExcelFile();

        Wait(seconds: waitMessageboxInSeconds, showOnScreen: true, onScreenText: "Starting Excel");
        Log("Starting Excel");
        START(mainWindowTitle: "*loginvsi*Excel*", mainWindowClass: "*XLMAIN*", timeout: 60);
        Wait(globalWaitInSeconds);
        SkipFirstRunDialogs();
        MainWindow.Maximize();
        MainWindow.Focus();
        Wait(globalWaitInSeconds);
    }
    
     private void DeleteTempFiles()
    {
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
            /* Commented out because it may delete other important temp files 
            foreach (var file in Directory.GetFiles(tempFolder, "*.tmp"))
            {
                File.Delete(file);
                Log("Deleted file: " + file);
            } */
        }
    }
    
    private void DownloadExcelFile()
    {
        int waitMessageboxInSeconds = 2;
        Wait(seconds: waitMessageboxInSeconds, showOnScreen: true, onScreenText: "Downloading Excel file if it doesn't exist");
        var temp = GetEnvironmentVariable("TEMP");
        string loginEnterpriseDir = $"{temp}\\LoginEnterprise";

        if (!Directory.Exists(loginEnterpriseDir))
        {
            Directory.CreateDirectory(loginEnterpriseDir);
            Log("Created directory: " + loginEnterpriseDir);
        }

        string excelFile = $"{loginEnterpriseDir}\\loginvsi.xlsx";
        Wait(waitMessageboxInSeconds);
        Log("Downloading Excel file if it doesn't exist");
        CopyFile(KnownFiles.ExcelSheet, excelFile, overwrite: false, continueOnError: true);
    }
    private void SkipFirstRunDialogs()
    {
        int loopCount = 2; // configurable number of loops
        for (int i = 0; i < loopCount; i++)
        {
            var dialog = FindWindow(
                className: "Win32 Window:NUIDialog",
                processName: "EXCEL",
                continueOnError: true,
                timeout: 3);
            while (dialog != null)
            {
                Wait(seconds: 2, showOnScreen: true, onScreenText: "Closing first run dialog if it exists");
                dialog.Close();
                dialog = FindWindow(
                    className: "Win32 Window:NUIDialog",
                    processName: "EXCEL",
                    continueOnError: true,
                    timeout: 3);
            }
        }
    }
}
