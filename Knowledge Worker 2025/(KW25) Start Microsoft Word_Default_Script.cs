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
        Wait(waitMessageboxInSeconds);
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
