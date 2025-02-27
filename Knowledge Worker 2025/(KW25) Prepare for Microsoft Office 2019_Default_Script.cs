// TARGET:winword.exe /t
// START_IN:

/////////////
// Office 2019 Prepare
// Workload: KnowledgeWorker 2025
// Version: 1.0
/////////////

using LoginPI.Engine.ScriptBase;
using LoginPI.Engine.ScriptBase.Components;
using System;
using System.IO;

public class PrepareOffice2019_DefaultScript : ScriptBase
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

        Log("Dismissing first run Word dialogs");
        START(mainWindowTitle: "*Word*", mainWindowClass: "Win32 Window:OpusApp", processName: "WINWORD", timeout: 60, continueOnError: true);
        int loopCount = 2; // configurable number of loops
        for (int i = 0; i < loopCount; i++)
        {
            var openDialog = MainWindow.FindControlWithXPath(
                xPath: "*:NUIDialog", 
                timeout: 5, 
                continueOnError: true);
            
            if (openDialog is object)
            {
                if (openDialog.GetTitle().StartsWith("First things", StringComparison.CurrentCultureIgnoreCase))
                {
                    Wait(seconds: 2, showOnScreen: true, onScreenText: "Closing first things first dialog if it exists");
                    
                    openDialog.FindControl(
                        className: "RadioButton:NetUIRadioButton", 
                        title: "Install updates only", 
                        continueOnError: true)?.Click();
                        
                    openDialog.FindControl(
                        className: "Button:NetUIButton", 
                        title: "Accept", 
                        continueOnError: true)?.Click();
                    
                    openDialog = MainWindow.FindControlWithXPath(
                        xPath: "Pane:NUIDialog", 
                        timeout: 5, 
                        continueOnError: true);
                        
                    if (openDialog is object)
                    {
                        openDialog.Type("{ALT+i}", hideInLogging: false);
                        Wait(1);
                        openDialog.Type("{ALT+a}", hideInLogging: false);
                    }
                    
                    openDialog = MainWindow.FindControlWithXPath(
                        xPath: "Pane:NUIDialog", 
                        timeout: 5, 
                        continueOnError: true);
                        
                    if (openDialog is object)
                    {
                        ABORT("Could not close outlooks First things first dialog");
                    }
                }
                else
                {
                    openDialog.Type("{ESC}");
                }
            }
        }

        Wait(2);
        STOP();
    }
}
