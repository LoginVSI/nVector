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
        // Define relevant folders
        string excelFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft", "Excel");
        string unsavedFilesFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft", "Office", "UnsavedFiles");
        string tempFolder = Path.GetTempPath();

        // Delete from Excel AutoRecover folder
        if (Directory.Exists(excelFolder))
        {
            foreach (var file in Directory.GetFiles(excelFolder, "*.xlsb")) // Binary workbooks
            {
                File.Delete(file);
                Log("Deleted file: " + file);
            }
            foreach (var file in Directory.GetFiles(excelFolder, "*.xar")) // Excel archive files (rare, but keeping your logic)
            {
                File.Delete(file);
                Log("Deleted file: " + file);
            }
            foreach (var file in Directory.GetFiles(excelFolder, "*.xls*")) // All Excel formats (xlsx, xlsm, etc.)
            {
                File.Delete(file);
                Log("Deleted file: " + file);
            }
            foreach (var file in Directory.GetFiles(excelFolder, "*.tmp")) // Temporary Excel files
            {
                File.Delete(file);
                Log("Deleted file: " + file);
            }
            foreach (var file in Directory.GetFiles(excelFolder, "~$*.xls*")) // Excel lock files
            {
                File.Delete(file);
                Log("Deleted file: " + file);
            }
        }

        // Delete from Office Unsaved Files folder (NEW ADDITION)
        if (Directory.Exists(unsavedFilesFolder))
        {
            foreach (var file in Directory.GetFiles(unsavedFilesFolder, "*.xls*")) // Unsaved Excel files
            {
                File.Delete(file);
                Log("Deleted file: " + file);
            }
        }

        // Delete from Temp folder
        if (Directory.Exists(tempFolder))
        {
            foreach (var file in Directory.GetFiles(tempFolder, "~$*.xls*")) // Excel lock files
            {
                File.Delete(file);
                Log("Deleted file: " + file);
            }
            foreach (var file in Directory.GetFiles(tempFolder, "Excel*.tmp")) // Excel-related temp files
            {
                File.Delete(file);
                Log("Deleted file: " + file);
            }
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
        string editedExcelFile = $"{loginEnterpriseDir}\\edited.xlsx";

        if (File.Exists(excelFile))
        {
            File.Delete(excelFile);
            Log("Deleted existing file: " + excelFile);
        }

        if (File.Exists(editedExcelFile))
        {
            File.Delete(editedExcelFile);
            Log("Deleted existing file: " + editedExcelFile);
        }

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
