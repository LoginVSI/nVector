// TARGET:excel.exe "C:\Users\NVtestuser\Desktop\loginvsi.xlsx"
// START_IN:

/////////////
// Excel Application
// Workload: KnowledgeWorker 2025
// Version: 1.0
/////////////

using LoginPI.Engine.ScriptBase;
using LoginPI.Engine.ScriptBase.Components;
using System;
using System.Runtime.InteropServices;

public class M365Excel_DefaultScript : ScriptBase
{
    const string ProcessName = "EXCEL";
    
    // Information that is shared across the functions
    string _newDocName;
    string _tempFolder;
    IWindow _activeDocument;
    IWindow _dataSheetArea;
    Location _row1Location;
    bool _isOffice365;
    
    // Import the user32.dll function to simulate mouse events.
    [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
    public static extern void mouse_event(uint dwFlags, uint dx, uint dy, int dwData, UIntPtr dwExtraInfo);

    public const uint MOUSEEVENTF_WHEEL = 0x0800; // Constant for a mouse wheel event
    
    private void Execute()
    
    {
    int ctrlTabIterations = 10; // Number of iterations for scrolling interactions
int ctrlTabWaitSecondsBeforeScroll = 3; // Wait time (in seconds) before scrolling to allow the page to load
int ctrlTabWaitSecondsAfterScroll = 1;  // Wait time (in seconds) after scrolling before next iteration
        // This is a language dependent script. English is required.

        _tempFolder = GetEnvironmentVariable("TEMP");

        // Optionally you can use the MyDocuments folder for file storage by setting the temp folder as follows
        // _tempFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        // Directory.CreateDirectory($"{_tempFolder}\\LoginPI");

        Log(_tempFolder);

        // Download file from the appliance through the KnownFiles method, if it already exists: Skip Download.
        /*Wait(seconds: 1, showOnScreen: true, onScreenText: "Get .xlsx file");
        if (!(FileExists($"{_tempFolder}\\LoginPI\\loginvsi.xlsx")))
        {
            Log("Downloading File");
            CopyFile(KnownFiles.ExcelSheet, $"{_tempFolder}\\LoginPI\\loginvsi.xlsx");
        }
        else
        {
            Log("File already exists");
        }
        */

        // Click the Start Menu
        Wait(seconds: 1, showOnScreen: true, onScreenText: "Start Menu");
        /*Type("{LWIN}");
        Wait(2);
        Type("{ESC}");*/

        // Start Application
        Wait(seconds: 1, showOnScreen: true, onScreenText: "Starting Excel");
        // START(mainWindowTitle: "*Excel*", mainWindowClass: "*XLMAIN*", timeout: 30);
        ShellExecute(@"excel.exe C:\Users\NVtestuser\Desktop\loginvsi.xlsx",forceKillOnExit:true,waitForProcessEnd:false);        
        //START(mainWindowTitle: "*Excel*", timeout: 30);
        var MainWindow = FindWindow(title:"*Excel*");
        Wait(1);
        MainWindow.Maximize();
        MainWindow.Focus();
        MainWindow.MoveMouseToCenter();
        Wait(1);
        MainWindow.Click();
        // Wait(999);
        Wait(1);
        
        

        // This is going to be the name of document when we save it later on
        // InitialiseSharedInformation();
        
        //DoScrolling();
        // Usage of Scroll():
//   - direction: "Down" to scroll down or "Up" to scroll up.
//   - scrollCount: Number of scroll events to send.
//   - notches: Number of notches per event (1 notch is typically 120).
//   - waitTime: Time in seconds to wait between each scroll event.
// Example:
//   Scroll("Down", 20, 1, 0.2);
//   Scroll("Up", 10, 2, 0.3);
Type("{ctrl+a}{ctrl+c}{ctrl+v}{ctrl+v}{ctrl+v}",cpm:50);
Scroll("Down", 10, 1, 0.2);
Scroll("Up", 5, 2, 0.3);
Type("{alt}nr",cpm:50);
Wait(1);
Type("{enter}");
Wait(1);
MainWindow.FindControl(className : "Text:NetUISimpleButton", title : "Zoom 100%").Click();
Wait(1);
Type("{alt+0}{enter}",cpm:50);
Wait(1);
MainWindow.FindControl(className : "DataItem:XLSpreadsheetCell", title : "F10").Click();
Type("{ctrl+a}{ctrl+c}{ctrl+v}{ctrl+v}{ctrl+v}",cpm:50);
Scroll("Down", 10, 1, 0.2);
Scroll("Up", 5, 2, 0.3);
Wait(1);
MainWindow.Close();


/*


        _activeDocument.Minimize();
        Wait(1);
        _activeDocument.Maximize();

        DoCopyPaste();
        
        CreateChart();
        
        SaveAs();
        
        // Stop the application
        Wait(seconds: 1, showOnScreen: true, onScreenText: "Stop App");
        MainWindow.Close();
        // STOP();

    }

    void InitialiseSharedInformation()
    {
        _newDocName = "edited";
       _activeDocument = OpenLoginVsiDoc();
        _dataSheetArea = _activeDocument.FindControlWithXPath(xPath: "Pane:XLDESK");
        // Although excel in the latest versions exposes the rows as UI elements, searching them is pretty expensive.
        // So we go by offset to find them
        _row1Location = _dataSheetArea.GetBounds().LeftTop.Move(12, 30);
        var row = _dataSheetArea.FindControl("DataItem", "1", continueOnError: true, timeout: 1);
        _isOffice365 = row is object;
        if (!_isOffice365)
        {
            Log("This is not office 365");
        }
    }
    
    void SaveAs()
    {
        // Saving the file in temp
        Wait(seconds: 1, showOnScreen: true, onScreenText: "Saving File");
        _activeDocument.Type("{F12}", cpm: 0);
        Wait(1);

        var filename = $"{_tempFolder}\\LoginPI\\{_newDocName}.xlsx";
        // We use a completely different filename here, so its easier to distinguish it from the other window
        // Remove file if it already exists
        if (FileExists(filename))
        {
            Log("Removing file");
            RemoveFile(path: filename);
        }
        else
        {
            Log("File already removed");
        }

        // Saving the file in temp 
        var SaveAs = GetFileDialog();

        var fileNameBox = SaveAs.FindControl(className: "Edit:Edit", title: "File name:");
        fileNameBox.Click();
        Wait(1);
        SetTextBoxText(fileNameBox, filename, cpm: 1000);
        StartTimer("Saving_file");
        SaveAs.Type("{ENTER}");
        FindWindow(title: $"{_newDocName}*", processName: ProcessName);
        StopTimer("Saving_file");
        */
    }
    
    void CreateChart()
    {
        //Create a Chart
        Wait(seconds: 1, showOnScreen: true, onScreenText: "Create a chart");
        _row1Location.Click();
        Wait(1);
        KeyDown(KeyCode.SHIFT);
        _activeDocument.Type("{DOWN}".Repeat(10), cpm: 1000);
        KeyUp(KeyCode.SHIFT); 
        Wait(1);
        var chartShortcut = _isOffice365 ? "{ALT+N}C1" : "{ALT+N}C";
        _activeDocument.Type(chartShortcut, cpm: 120); Wait(1);
        Type("{ENTER}"); // we do not use the window to type here, because that would force a 'focus window', which breaks the chart selector focus
        KeyDown(KeyCode.SHIFT);
        _activeDocument.Type("{UP}".Repeat(5), cpm: 100);
        _activeDocument.Type("{RIGHT}".Repeat(5), cpm: 100);
        KeyUp(KeyCode.SHIFT);
        Wait(2);
    }
    
    void DoCopyPaste()
    {
        //Copy a row and paste
        Wait(seconds: 1, showOnScreen: true, onScreenText: "Copy & Paste");
        var row5Location = _dataSheetArea.GetBounds().LeftTop.Move(12, 111);
        row5Location.RightClick();
        Wait(1);
        Type("i"); // Type without window reference to avoid focus setting. Focus setting closes the context menu
        _activeDocument.Type("{CTRL+Y}".Repeat(15), cpm: 100);
        Wait(1);
        _row1Location.Click(); Wait(1);
        _activeDocument.Type("{CTRL+C}"); Wait(1);
        row5Location.Click(); Wait(1);
        KeyDown(KeyCode.SHIFT); Wait(1);
        _activeDocument.Type("{DOWN}".Repeat(15), cpm: 100);
        KeyUp(KeyCode.SHIFT); Wait(1);
        _activeDocument.Type("{CTRL+V}"); Wait(1);

        _activeDocument.Type("{ESC}"); Wait(1);
    }
    
    void DoScrolling()
    {
        //Scroll through excel document
        Wait(seconds: 1, showOnScreen: true, onScreenText: "Scroll");
        _activeDocument.MoveMouseToCenter();
        MouseDown();
        Wait(1);
        MouseUp();
        _activeDocument.Type("{PAGEDOWN}".Repeat(6), cpm: 1000);
        Wait(1);
        _activeDocument.Type("{PAGEUP}".Repeat(5), cpm: 1000);
        Wait(1);
        _activeDocument.Type("{PAGEDOWN}".Repeat(5), cpm: 1000);
        Wait(1);
        _activeDocument.Type("{PAGEUP}".Repeat(6), cpm: 1000);
        Wait(1);
    }
    
    IWindow OpenLoginVsiDoc()
    {
        var appWasLeftOpen = MainWindow.GetTitle().Contains(_newDocName);
        if (appWasLeftOpen)
        {
            Log("Excel was left open from previous run");
        }
        else
        {
            Wait(10);

            SkipFirstRunDialogs();

            // If a previous run of Excel crashed, we need to close the *recover document* frame now
            MainWindow.FindControl(className: "Button:NetUIButton", title: "Close", continueOnError: true, timeout: 5)?.Click();
        }

            // Open "Open File" window and start measurement.
        Wait(seconds: 1, showOnScreen: true, onScreenText: "Open File Window");
        MainWindow.Type("{CTRL+O}");
        MainWindow.Type("{ALT+O+O}");
        StartTimer("Open_Window");
        var openWindow = GetFileDialog();
        StopTimer("Open_Window");
        Wait(1);
        openWindow.Click();

        //Navigate to copied XLSX file and press Open, measure time to open the file.
        Wait(seconds: 1, showOnScreen: true, onScreenText: "Open File");
        var fileNameBox = openWindow.FindControl(className: "Edit:Edit", title: "File name:");
        fileNameBox.Click();
        Wait(1);
        SetTextBoxText(fileNameBox, $"{_tempFolder}\\LoginPI\\loginvsi.xlsx", cpm: 1000);
        Wait(1);
        openWindow.FindControl(className: "SplitButton:Button", title: "&Open").Click();
        StartTimer("Open_Excel_Document");
        _activeDocument = FindWindow(className: "*XLMAIN*", title: "loginvsi*");
        StopTimer("Open_Excel_Document");
        
        // We close the old window here, so our FindWindow does not find it window
        if (appWasLeftOpen)
        {
            _activeDocument.Minimize();
            Wait(3, showOnScreen: true, onScreenText: $"Closing window {MainWindow.GetTitle()}");
            MainWindow.Type("{CTRL+F4}");
            Wait(1);

            // Win32 Window:NUIDialog   => Excel 365
            // Pane:NUIDialog           => Excel 2019/2016
            var confirmDialog = FindWindow(className: "*NUIDialog", processName: ProcessName, continueOnError: true, timeout:10);
            if (confirmDialog != null)
            {
                confirmDialog.FindControl(title: "*Don*")?.Click();
            }
            _activeDocument.Maximize();
        }

        return _activeDocument;
    }

    void SkipFirstRunDialogs()
    {
        var dialog = FindWindow(className: "Win32 Window:NUIDialog", processName: ProcessName, continueOnError: true, timeout: 1);
        while (dialog != null)
        {
            dialog.Close();
            dialog = FindWindow(className: "Win32 Window:NUIDialog", processName: ProcessName, continueOnError: true, timeout: 10);
        }
    }

    IWindow GetFileDialog()
    {
        // Finding a dialog window is faster and more reliable if we use the global find window
        var dialog = FindWindow(className: "Win32 Window:#32770", processName: ProcessName, continueOnError: true, timeout:10);
        if (dialog is null)
        {
            ABORT("File dialog could not be found");
        }
        return dialog;
    }

    ///
    /// This method types the given text to the textbox (any existing text is cleared)
    /// After typing, it confirms the resulting value.
    /// If it does not match, it will clear the textbox and try again
    ///
    void SetTextBoxText(IWindow textBox, string text, int cpm = 800)
    {
        var numTries = 1;
        string currentText = null;
        do
        {
            textBox.Type("{CTRL+a}");
            Wait(0.5);
            textBox.Type(text, cpm: cpm);
            Wait(1);
            currentText = textBox.GetText();
            if (currentText != text)
                CreateEvent($"Typing error in attempt {numTries}", $"Expected '{text}', got '{currentText}'");
        }
        while (++numTries < 5 && currentText != text);
        if (currentText != text)
            ABORT($"Unable to set the correct text '{text}', got '{currentText}'");
    }
        void Scroll(string direction, int scrollCount, int notches, double waitTime)
{
if (waitTime <= 0)
{
throw new ArgumentException("Scroll waitTime must be greater than 0 seconds.");
}

int sign = direction.Equals("Down", StringComparison.OrdinalIgnoreCase) ? -1 : 1;
int delta = sign * 120 * notches;

Log($"Scrolling mouse {direction} {scrollCount} times, {notches} notch(es) per scroll, with {waitTime} sec between each scroll.");
for (int i = 0; i < scrollCount; i++)
{
mouse_event(MOUSEEVENTF_WHEEL, 0, 0, delta, UIntPtr.Zero);
Wait(waitTime);
}
}

}


