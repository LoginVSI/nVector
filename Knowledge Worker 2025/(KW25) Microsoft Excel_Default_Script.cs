// TARGET:excel.exe /e
// START_IN:

/////////////
// Excel Application
// Workload: KnowledgeWorker 2025
// Version: 1.0
/////////////

using System;
using System.IO;
using System.Runtime.InteropServices;
using LoginPI.Engine.ScriptBase;
using LoginPI.Engine.ScriptBase.Components;

public class Excel_DefaultScript : ScriptBase
{
    // =====================================================
    // Global Timings and Speeds
    // =====================================================
    int globalTimeoutInSeconds = 60;              // Timeout for actions
    int globalWaitInSeconds = 8;                  // Standard wait time between actions
    int waitMessageboxInSeconds = 8;              // Duration for onscreen wait messages
    int keyboardShortcutsCPM = 30;                // Typing speed for keyboard shortcuts
    int charactersPerMinuteToType = 300;          // Typing speed for file dialogs and text input

    // New timing variables for chart insertion
    int waitForGraphToShow = 15;                   // Wait time for the graph to show
    int waitForGraphFullscreenToLoad = 15;         // Wait time for the fullscreen graph to load
    int waitForQuickLayout = 10;                   // Wait time for the quick layout drop down to show
    int waitAfterQuickLayoutPreviews = 15;         // Wait time after the quick layout previews

    // =====================================================
    // Shared Information (set during runtime)
    // =====================================================
    string _tempFolder;
    IWindow _activeDocument;
    string _newDocName;

    // Import mouse scroll functionality
    [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
    public static extern void mouse_event(uint dwFlags, uint dx, uint dy, int dwData, UIntPtr dwExtraInfo);
    public const uint MOUSEEVENTF_WHEEL = 0x0800; // Constant for a mouse wheel event

    // =====================================================
    // Execute Method
    // =====================================================
    private void Execute()
    {
        // -------------------------------
        // Environment Setup and Forced File Download
        // -------------------------------
        _tempFolder = GetEnvironmentVariable("TEMP");
        Log("TEMP folder: " + _tempFolder);
        string downloadDir = $"{_tempFolder}\\LoginEnterprise";
        if (!Directory.Exists(downloadDir))
        {
            Directory.CreateDirectory(downloadDir);
            Log("Created directory: " + downloadDir);
        }
        string excelFilePath = $"{downloadDir}\\loginvsi.xlsx";
        // Force download by always overwriting
        Log("Downloading Excel file (overwrite).");
        CopyFile(KnownFiles.ExcelSheet, excelFilePath, continueOnError: false, overwrite: true);
        
        // -------------------------------
        // Simulate Start Menu Interaction
        // -------------------------------
        Log("Opening Start Menu");
        Wait(seconds: waitMessageboxInSeconds, showOnScreen: true, onScreenText: "Start Menu");
        Type("{LWIN}", hideInLogging: false);
        Wait(seconds: 3);
        Type("{LWIN}", hideInLogging: false);
        Wait(seconds: 1);
        Type("{ESC}", hideInLogging: false);
        
        // -------------------------------
        // Skip First-Run Dialogs (Pre-Excel)
        // -------------------------------
        SkipFirstRunDialogs();
        
        // -------------------------------
        // Open Excel File via File Dialog
        // -------------------------------
        var mainWindow = FindWindow(className: "*XLMAIN*", title: "*Excel*", processName: "EXCEL", timeout: globalTimeoutInSeconds);
        Wait(seconds: globalWaitInSeconds);
        mainWindow.Focus();
        mainWindow.Maximize();
        
        Log("Opening Excel file via open file dialog");
        Wait(seconds: waitMessageboxInSeconds, showOnScreen: true, onScreenText: "Open Excel file");
        mainWindow.Type("{CTRL+O}{ALT+O+O}", cpm: keyboardShortcutsCPM, hideInLogging: false);
        StartTimer("Open_Excel_Dialog");
        var openWindow = FindWindow(className: "Win32 Window:#32770", processName: "EXCEL", continueOnError: false, timeout: globalTimeoutInSeconds);
        StopTimer("Open_Excel_Dialog");
        Wait(seconds: globalWaitInSeconds);
        var fileNameBox = openWindow.FindControl(className: "Edit:Edit", title: "File name:", timeout: globalTimeoutInSeconds);
        fileNameBox.Click();
        Wait(seconds: globalWaitInSeconds);
        SetTextBoxText(fileNameBox, excelFilePath, cpm: 300);
        openWindow.Type("{ENTER}", hideInLogging: false);
        StartTimer("Open_Excel_Document");
        _activeDocument = FindWindow(className: "*XLMAIN*", title: "loginvsi*", processName: "EXCEL", timeout: globalTimeoutInSeconds);
        StopTimer("Open_Excel_Document");
        _activeDocument.Focus();
        
        // -------------------------------
        // Check and Close Existing "Edited" Window (if any)
        // -------------------------------
        Wait(seconds: waitMessageboxInSeconds, showOnScreen: true, onScreenText: "If it exists close the existing Excel window");
        _newDocName = "edited";
        for (int attempt = 0; attempt < 2; attempt++)
        {
            var editedWindow = FindWindow(className: "*XLMAIN*", title: $"{_newDocName}*", processName: "EXCEL", timeout: 2, continueOnError: true);
            if (editedWindow != null)
            {
                Log("Existing edited window found. Closing it.");
                editedWindow.Focus();
                editedWindow.Maximize();
                Wait(seconds: globalWaitInSeconds, showOnScreen: true, onScreenText: "Closing edited window");
                editedWindow.Type("{alt+f4}",hideInLogging:false);
                Wait(seconds: globalWaitInSeconds);
                // Type "n" to cancel any confirmation dialog
                editedWindow.Type("n", hideInLogging: false);
                Wait(seconds: globalWaitInSeconds);
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
        // -------------------------------
        // Mouse Scroll Through the Document
        // -------------------------------
        _activeDocument = FindWindow(className: "*XLMAIN*", title: "loginvsi*", processName: "EXCEL", timeout: globalTimeoutInSeconds);
        _activeDocument.Focus();
        _activeDocument.Maximize();
        Wait(seconds: waitMessageboxInSeconds, showOnScreen: true, onScreenText: "Scroll in the spreadsheet");
        Log("Scrolling through Excel document using mouse scroll");
        // For example: scroll down 5 notches with 0.2 sec between scrolls
        Scroll("Down", 30, 1, 0.1);
        Scroll("Up", 30, 1, 0.1);
        
        // -------------------------------
        // Insert Chart Workflow
        // -------------------------------
        Wait(seconds: waitMessageboxInSeconds, showOnScreen: true, onScreenText: "Add a graph, make it fullscreen, and change the graph layout");
        // Open Insert Chart window via shortcut {ALT}NR
        _activeDocument.Type("{ALT}NR", cpm: keyboardShortcutsCPM, hideInLogging: false);
        Wait(seconds: globalWaitInSeconds);
        // Find the Insert Chart dialog
        var insertChartDialog = FindWindow(className: "Win32 Window:NUIDialog", title: "Insert Chart", processName: "EXCEL", timeout: globalTimeoutInSeconds);
        // In the dialog, click the "All Charts" tab
        var allChartsTab = insertChartDialog.FindControl(className: "TabItem:NetUITabHeader", title: "All Charts", timeout: globalTimeoutInSeconds);
        Wait(seconds: globalWaitInSeconds);
        allChartsTab.Click();
        Wait(seconds: globalWaitInSeconds);
        // Find and click the "Column" gallery button
        var columnButton = insertChartDialog.FindControl(className: "ListItem:NetUIGalleryButton", title: "Column", timeout: globalTimeoutInSeconds);
        Wait(seconds: globalWaitInSeconds);
        columnButton.Click();
        Wait(seconds: globalWaitInSeconds);
        // Find the "Clustered Column" label and load the chart
        insertChartDialog.FindControl(className: "Text:NetUILabel", title: "Clustered Column", timeout: globalTimeoutInSeconds);
        Wait(seconds: globalWaitInSeconds);
        insertChartDialog.Focus();
        insertChartDialog.Type("{ENTER}", hideInLogging: false);
        // Wait for the chart to show
        Wait(seconds: waitForGraphToShow, showOnScreen: true, onScreenText: "Waiting for graph to show");
        
        // Make the new graph go full screen with {F11}
        _activeDocument.Type("{F11}", hideInLogging: false);
        Wait(seconds: waitForGraphFullscreenToLoad, showOnScreen: true, onScreenText: "Waiting for full screen graph");
        
        // Open Quick Layout drop-down via {ALT}JCL
        _activeDocument.Type("{ALT}JCL", cpm: keyboardShortcutsCPM, hideInLogging: false);
        Wait(seconds: waitForQuickLayout, showOnScreen: true, onScreenText: "Quick Layout drop-down");
        // Simulate keystrokes to navigate layout options:
        _activeDocument.Type("{DOWN}", hideInLogging: false); Wait(2);
        _activeDocument.Type("{DOWN}", hideInLogging: false); Wait(2);
        _activeDocument.Type("{DOWN}", hideInLogging: false); Wait(2);
        _activeDocument.Type("{RIGHT}", hideInLogging: false); Wait(2);
        _activeDocument.Type("{UP}", hideInLogging: false); Wait(2);
        _activeDocument.Type("{UP}", hideInLogging: false); Wait(2);
        _activeDocument.Type("{UP}", hideInLogging: false); Wait(2);
        _activeDocument.Type("{RIGHT}", hideInLogging: false); Wait(2);
        _activeDocument.Type("{DOWN}", hideInLogging: false); Wait(2);
        _activeDocument.Type("{DOWN}", hideInLogging: false); Wait(2);
        // Global wait then close the drop-down
        Wait(seconds: waitAfterQuickLayoutPreviews);
        _activeDocument.Type("{ESC}", hideInLogging: false);
        Wait(seconds: globalWaitInSeconds);
        
        // -------------------------------
        // Save the Edited Document
        // -------------------------------
        Wait(seconds: waitMessageboxInSeconds, showOnScreen: true, onScreenText: "Save the Excel doc");
        SaveAs(downloadDir);
        
        // -------------------------------
        // End of Script (Excel remains open)
        // -------------------------------
        Log("Script complete. Excel remains open.");
    }
    
    // =====================================================
    // Scroll Helper Method (Real Mouse Scroll)
    // =====================================================
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
        Log($"Completed scrolling mouse {direction} {scrollCount} times.");
    }
    
    // =====================================================
    // Save As (Save the Edited Excel File)
    // =====================================================
    void SaveAs(string downloadDir)
    {
        Wait(seconds: waitMessageboxInSeconds, showOnScreen: true, onScreenText: "Saving File");
        _activeDocument.Type("{F12}", cpm: charactersPerMinuteToType, hideInLogging: false);
        Wait(seconds: globalWaitInSeconds);
        var filename = $"{downloadDir}\\{_newDocName}.xlsx";
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
        var fileNameBox = saveAsDialog.FindControl(className: "Edit:Edit", title: "File name:", timeout: globalTimeoutInSeconds);
        fileNameBox.Click();
        Wait(seconds: globalWaitInSeconds);
        SetTextBoxText(fileNameBox, filename, cpm: 300);
        saveAsDialog.Type("{ENTER}",hideInLogging:false);
        StartTimer("Saving_Excel_file");
        FindWindow(title: $"{_newDocName}*", processName: "EXCEL", timeout: globalTimeoutInSeconds);
        StopTimer("Saving_Excel_file");
    }
    
    // =====================================================
    // Get File Dialog Helper
    // =====================================================
    IWindow GetFileDialog()
    {
        var dialog = FindWindow(className: "Win32 Window:#32770", processName: "EXCEL", continueOnError: true, timeout: globalTimeoutInSeconds);
        if (dialog is null)
        {
            ABORT("File dialog could not be found");
        }
        return dialog;
    }
    
    // =====================================================
    // Skip First-Run Dialogs Helper
    // =====================================================
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
    
    // =====================================================
    // SetTextBoxText Helper Method
    // =====================================================
    void SetTextBoxText(IWindow textBox, string text, int cpm = 300)
    {
        int numTries = 1;
        string currentText = null;
        do
        {
            textBox.Type("{CTRL+a}",hideInLogging:false);
            Wait(globalWaitInSeconds);
            textBox.Type(text, cpm: cpm);
            Wait(globalWaitInSeconds);
            currentText = textBox.GetText();
            if (currentText != text)
                CreateEvent($"Typing error in attempt {numTries}", $"Expected '{text}', got '{currentText}'");
        }
        while (++numTries < 5 && currentText != text);
        if (currentText != text)
            ABORT($"Unable to set the correct text '{text}', got '{currentText}'");
    }
}
