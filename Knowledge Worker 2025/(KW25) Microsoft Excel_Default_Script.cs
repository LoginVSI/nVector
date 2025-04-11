// TARGET:excel.exe /s
// START_IN:

/////////////
// Excel Application
// Workload: KnowledgeWorker 2025
// Version: 1.0
/////////////

using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Diagnostics;
using LoginPI.Engine.ScriptBase;
using LoginPI.Engine.ScriptBase.Components;

public class Excel_DefaultScript : ScriptBase
{
    // =====================================================
    // Global Timings and Speeds
    // =====================================================
    int globalTimeoutInSeconds = 60;              // Timeout for actions
    int globalWaitInSeconds = 3;                  // Standard wait time between actions
    int waitMessageboxInSeconds = 2;              // Duration for onscreen wait messages
    int keyboardShortcutsCPM = 15;                // Typing speed for keyboard shortcuts
    int charactersPerMinuteToType = 600;          // Typing speed for file dialogs and text input
    int startMenuWaitInSeconds = 5;               // Duration for Start Menu wait between interactions
    int waitAfterScrolling = 8;                   // Duration after scrolling before inserting chart

    // Additional timing variables (chart insertion, etc.)
    int waitForGraphToShow = 15;
    int waitForGraphFullscreenToLoad = 15;
    int waitForQuickLayout = 10;
    int waitAfterQuickLayoutPreviews = 15;
    int waitCyclingQuickLayouts = 8;

    // =====================================================
    // Shared Information (set during runtime)
    // =====================================================
    string _tempFolder;
    IWindow _activeDocument;
    string _newDocName;

    // Import mouse scroll functionality
    [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
    public static extern void mouse_event(uint dwFlags, uint dx, uint dy, int dwData, UIntPtr dwExtraInfo);
    public const uint MOUSEEVENTF_WHEEL = 0x0800;

    private void Execute()
    {
        // (The Excel file is already downloaded by the start script.)
        _tempFolder = GetEnvironmentVariable("TEMP");
        string downloadDir = $"{_tempFolder}\\LoginEnterprise";
        string excelFilePath = $"{downloadDir}\\loginvsi.xlsx";

        // =====================================================
        // Simulate Start Menu Interaction
        // =====================================================
        Log("Simulating Start Menu interaction.");
        Wait(startMenuWaitInSeconds);
        Type("{LWIN}", hideInLogging:false);
        Wait(startMenuWaitInSeconds);
        Type("{LWIN}", hideInLogging:false);
        Wait(1);
        Type("{ESC}", hideInLogging:false);
        Wait(startMenuWaitInSeconds);

        // =====================================================
        // Launch new blank Excel workbook
        // =====================================================
        try
        {
            ShellExecute("excel /s", waitForProcessEnd: false, timeout: globalTimeoutInSeconds, continueOnError: true, forceKillOnExit: false);
            /* Alternate start blank Excel document function:
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = "excel.exe",
                Arguments = "/s",
                UseShellExecute = true
            };
            Process.Start(startInfo); */
        }
        catch (Exception ex)
        {
            ABORT("Error starting process: " + ex.Message);
        } 

        var newExcelWindow = FindWindow(title:"*Book*Excel*", processName:"EXCEL", continueOnError:false, timeout:globalTimeoutInSeconds);
        Wait(globalWaitInSeconds);

        // =====================================================
        // Close any extraneous Excel windows ("*loginvsi*" or "*edited*")
        // =====================================================
        CloseExtraWindows("EXCEL", "*loginvsi*");
        CloseExtraWindows("EXCEL", "*edited*");

        // =====================================================
        // Skip First-Run Dialogs
        // =====================================================
        Wait(globalWaitInSeconds);
        SkipFirstRunDialogs();

        // =====================================================
        // Bring new Excel instance into focus and open file via dialog
        // =====================================================
        Wait(startMenuWaitInSeconds);
        newExcelWindow.Focus();
        newExcelWindow.Maximize();
        Log("Opening Excel file via open file dialog");
        Wait(waitMessageboxInSeconds, true, "Open Excel file");
        newExcelWindow.Type("{CTRL+O}", cpm: keyboardShortcutsCPM, hideInLogging:false);
        Wait(globalWaitInSeconds);
        newExcelWindow.Type("{ALT+O+O}", cpm: keyboardShortcutsCPM, hideInLogging:false);
        StartTimer("Open_Excel_Dialog");
        var openWindow = FindWindow(className:"Win32 Window:#32770", processName:"EXCEL", continueOnError:false, timeout:globalTimeoutInSeconds);
        StopTimer("Open_Excel_Dialog");
        Wait(globalWaitInSeconds);
        var fileNameBox = openWindow.FindControl(className:"Edit:Edit", title:"File name:", timeout:globalTimeoutInSeconds);
        Wait(globalWaitInSeconds);
        fileNameBox.Click();
        fileNameBox.Type("{ALT+N}", hideInLogging:false);
        Wait(globalWaitInSeconds);
        SetTextBoxText(fileNameBox, excelFilePath, cpm: charactersPerMinuteToType);
        openWindow.Type("{ENTER}", hideInLogging:false);
        StartTimer("Open_Excel_Document");
        _activeDocument = FindWindow(className:"*XLMAIN*", title:"loginvsi*", processName:"EXCEL", timeout:globalTimeoutInSeconds);
        StopTimer("Open_Excel_Document");

        // Close any stray "Book" windows
        CloseExtraWindows("EXCEL", "*Book*");

        // =====================================================
        // Minimize and Maximize Excel
        // =====================================================
        Wait(globalWaitInSeconds);
        _activeDocument.Minimize();
        Wait(globalWaitInSeconds);
        _activeDocument.Maximize();
        _activeDocument.Focus();
        Wait(globalWaitInSeconds);

        // Scroll through the document
        Wait(waitMessageboxInSeconds, true, "Scroll in the spreadsheet");
        Log("Scrolling through Excel document using mouse scroll");
        Scroll("Down", 30, 1, 0.1);
        Scroll("Up", 30, 1, 0.1);

        // Insert Chart Workflow 
        Wait(waitAfterScrolling);
        Wait(waitMessageboxInSeconds, true, "Add a graph, make it fullscreen, and change the graph layout");
        _activeDocument.Type("{ALT}NR", cpm: keyboardShortcutsCPM, hideInLogging:false);
        var insertChartDialog = FindWindow(className:"Win32 Window:NUIDialog", processName:"EXCEL", timeout:globalTimeoutInSeconds);
        var allChartsTab = insertChartDialog.FindControl(className:"TabItem:NetUITabHeader", title:"All Charts", timeout:globalTimeoutInSeconds);
        Wait(globalWaitInSeconds);
        allChartsTab.Click();
        Wait(globalWaitInSeconds);
        var columnButton = insertChartDialog.FindControl(className:"ListItem:NetUIGalleryButton", title:"Column", timeout:globalTimeoutInSeconds);
        Wait(globalWaitInSeconds);
        columnButton.Click();
        Wait(globalWaitInSeconds);
        insertChartDialog.FindControl(className:"Text:NetUILabel", title:"Clustered Column", timeout:globalTimeoutInSeconds);
        Wait(globalWaitInSeconds);
        insertChartDialog.Focus();
        insertChartDialog.Type("{ENTER}", hideInLogging:false);
        Wait(waitForGraphToShow, true, "Waiting for graph to show");

        _activeDocument.Type("{F11}", hideInLogging:false);
        Wait(waitForGraphFullscreenToLoad, true, "Waiting for full screen graph");
        _activeDocument.Type("{ALT}JCL", cpm: keyboardShortcutsCPM, hideInLogging:false);
        Wait(waitForQuickLayout, true, "Quick Layout drop-down");
        _activeDocument.Type("{DOWN}", hideInLogging:false); Wait(waitCyclingQuickLayouts);
        _activeDocument.Type("{DOWN}", hideInLogging:false); Wait(waitCyclingQuickLayouts);
        _activeDocument.Type("{DOWN}", hideInLogging:false); Wait(waitCyclingQuickLayouts);
        Wait(waitAfterQuickLayoutPreviews);
        _activeDocument.Type("{ESC}", hideInLogging:false);
        Wait(globalWaitInSeconds);

        // Save As workflow
        Wait(waitMessageboxInSeconds, true, "Save the Excel doc");
        SaveAs(downloadDir);
        Log("Script complete. Excel remains open.");
    }

    void CloseExtraWindows(string processName, string titleMask)
    {
        // Maximum number of attempts to close extra windows including handling confirm dialogs
        int maxAttempts = 1;
        
        for (int attempt = 0; attempt < maxAttempts; attempt++)
        {
            var extraWindow = FindWindow(title: titleMask, processName: processName, timeout: 3, continueOnError: true);
            if (extraWindow == null)
            {
                // The window is already closed
                break;
            }
    
            // Give the window some focus and send initial close command
            Wait(globalWaitInSeconds);
            extraWindow.Focus();
            extraWindow.Maximize();
            Wait(globalWaitInSeconds);
            extraWindow.Type("{ESC}", hideInLogging: false);
            Wait(globalWaitInSeconds);
            extraWindow.Type("{ALT+F4}", hideInLogging: false);
            Wait(globalWaitInSeconds);
    
            // Check if the window still exists (a confirmation might have appeared)
            extraWindow = FindWindow(title: titleMask, processName: processName, timeout: 3, continueOnError: true);
            if (extraWindow != null)
            {
                // Wait a little longer and then send {ALT+N} for any confirmation popup
                Wait(globalWaitInSeconds);
                extraWindow.Type("{ALT+N}", hideInLogging: false);
                Wait(globalWaitInSeconds);
            }
        }
    }

    private void SkipFirstRunDialogs()
    {
        int loopCount = 2;
        for (int i = 0; i < loopCount; i++)
        {
            var dialog = FindWindow(className:"Win32 Window:NUIDialog", processName:"EXCEL", continueOnError:true, timeout:3);
            while (dialog != null)
            {
                Wait(globalWaitInSeconds);
                dialog.Close();
                dialog = FindWindow(className:"Win32 Window:NUIDialog", processName:"EXCEL", continueOnError:true, timeout:3);
            }
        }
    }

    // =====================================================
    // Helper: Scroll Function
    // =====================================================
    // Usage of Scroll():
    //   - direction: "Down" to scroll down or "Up" to scroll up.
    //   - scrollCount: Number of scroll events to send.
    //   - notches: Number of notches per event (1 notch is typically 120).
    //   - waitTime: Time in seconds to wait between each scroll event.
    // Example:
    //   Scroll("Down", 20, 1, 0.2);
    //   Scroll("Up", 10, 2, 0.3);
    void Scroll(string direction, int scrollCount, int notches, double waitTime)
    {
        if (waitTime <= 0)
            throw new ArgumentException("Scroll waitTime must be greater than 0 seconds.");
        int sign = direction.Equals("Down", StringComparison.OrdinalIgnoreCase) ? -1 : 1;
        int delta = sign * 120 * notches;
        Log($"Scrolling mouse {direction} {scrollCount} times, {notches} notch(es) each, waiting {waitTime}s between scrolls.");
        for (int i = 0; i < scrollCount; i++)
        {
            mouse_event(MOUSEEVENTF_WHEEL, 0, 0, delta, UIntPtr.Zero);
            Wait(seconds: waitTime);
        }
    }

    void SaveAs(string downloadDir)
    {
        Wait(waitMessageboxInSeconds, true, "Saving File");
        _activeDocument.Type("{F12}", cpm: charactersPerMinuteToType, hideInLogging:false);
        Wait(globalWaitInSeconds);
        var filename = $"{downloadDir}\\edited.xlsx";
        if (FileExists(filename))
        {
            Log("Removing file: " + filename);
            RemoveFile(path: filename);
        }
        else
        {
            Log("No existing file to remove.");
        }
        var saveAsDialog = GetFileDialog();
        var fileNameBox = saveAsDialog.FindControl(className:"Edit:Edit", title:"File name:", timeout:globalTimeoutInSeconds);
        fileNameBox.Click();
        fileNameBox.Type("{ALT+N}", hideInLogging:false);
        Wait(globalWaitInSeconds);
        SetTextBoxText(fileNameBox, filename, cpm: charactersPerMinuteToType);
        saveAsDialog.Type("{ENTER}", hideInLogging:false);
        StartTimer("Saving_Excel_file");
        FindWindow(title: "edited*", processName:"EXCEL", timeout:globalTimeoutInSeconds);
        StopTimer("Saving_Excel_file");
    }

    IWindow GetFileDialog()
    {
        var dialog = FindWindow(className:"Win32 Window:#32770", processName:"EXCEL", continueOnError:true, timeout:globalTimeoutInSeconds);
        if (dialog is null)
            ABORT("File dialog could not be found");
        return dialog;
    }

    void SetTextBoxText(IWindow textBox, string text, int cpm = 600)
    {
        int localWait = 3;
        int numTries = 1;
        string currentText = null;
        do
        {
            textBox.Type("{CTRL+a}", hideInLogging:false);
            Wait(localWait);
            textBox.Type(text, cpm: cpm);
            Wait(localWait);
            currentText = textBox.GetText();
            if (currentText != text)
                CreateEvent($"Typing error in attempt {numTries}", $"Expected '{text}', got '{currentText}'");
        }
        while (++numTries < 5 && currentText != text);
        if (currentText != text)
            ABORT($"Unable to set the correct text '{text}', got '{currentText}'");
    }
}

// =====================================================
// Helper Class for TextBox Operations
// =====================================================
public static class ScriptHelpers
{
    public static void SetTextBoxText(ScriptBase script, IWindow textBox, string text, int cpm = 600)
    {
        int localWait = 3;
        int numTries = 1;
        string currentText = null;
        do
        {
            textBox.Type("{CTRL+a}", hideInLogging:false);
            script.Wait(localWait);
            textBox.Type(text, cpm: cpm, hideInLogging:false);
            script.Wait(localWait);
            currentText = textBox.GetText();
            if (currentText != text)
                script.CreateEvent($"Typing error in attempt {numTries}", $"Expected '{text}', got '{currentText}'");
        }
        while (++numTries < 5 && currentText != text);
        if (currentText != text)
            script.ABORT($"Unable to set the correct text '{text}', got '{currentText}'");
    }
}
