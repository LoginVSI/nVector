// TARGET:winword.exe
// START_IN:

/////////////
// Office 2019 Prepare
// Workload: KnowledgeWorker 2025
// Version: 1.0
/////////////

using LoginPI.Engine.ScriptBase;
using LoginPI.Engine.ScriptBase.Components;
using System;

public class PrepareOffice2019_DefaultScript : ScriptBase
{
    private void Execute()
    {
        START(mainWindowTitle: "*Word*", mainWindowClass: "Win32 Window:OpusApp", processName: "WINWORD", timeout: 30, continueOnError: true);
        Wait(seconds: 3, showOnScreen: true, onScreenText: "Skip first run dialog");
        var openDialog = MainWindow.FindControlWithXPath(xPath: "*:NUIDialog", timeout: 5, continueOnError: true);
        if (openDialog is object)
        {
            if (openDialog.GetTitle().StartsWith("First things", StringComparison.CurrentCultureIgnoreCase))
            {
                Wait(seconds: 3, showOnScreen: true, onScreenText: "Closing first things first dialog");
                openDialog.FindControl(className: "RadioButton:NetUIRadioButton", title: "Install updates only", continueOnError: true)?.Click();
                openDialog.FindControl(className: "Button:NetUIButton", title: "Accept", continueOnError: true)?.Click();
                openDialog = MainWindow.FindControlWithXPath(xPath: "Pane:NUIDialog", timeout: 5, continueOnError: true);
                if (openDialog is object)
                {
                    openDialog.Type("{ALT+i}");
                    Wait(1);
                    openDialog.Type("{ALT+a}");
                }
                openDialog = MainWindow.FindControlWithXPath(xPath: "Pane:NUIDialog", timeout: 5, continueOnError: true);
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

        Wait(2);
        STOP();
    }
}
