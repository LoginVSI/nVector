// TARGET:winword.exe C:\Users\NVtestuser\Desktop\loginvsi.docx
// START_IN:

/////////////
// Word Application
// Workload: KnowledgeWorker 2025
// Version: 1.0
/////////////

using LoginPI.Engine.ScriptBase;
using LoginPI.Engine.ScriptBase.Components;
using System;
using System.Runtime.InteropServices;

public class M365Word_DefaultScript : ScriptBase

{
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

        var temp = GetEnvironmentVariable("TEMP");
        
        // Optionally you can use the MyDocuments folder for file storage by setting the temp folder as follows
        // var temp = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        // Directory.CreateDirectory($"{temp}\\LoginPI");

        // Define random integer
        var waitTime = 2;

        // Download file from the appliance through the KnownFiles method, if it already exists: Skip Download.
        /*Wait(seconds: 3, showOnScreen: true, onScreenText: "Get .docx file");
        if (!(FileExists($"{temp}\\LoginPI\\loginvsi.docx")))
        {
            Log("Downloading File");
            CopyFile(KnownFiles.WordDocument, $"{temp}\\LoginPI\\loginvsi.docx");
        }
        else
        {
            Log("File already exists");
        }

        // Click the Start Menu
        Wait(seconds: 3, showOnScreen: true, onScreenText: "Start Menu");
        Type("{LWIN}");
        Wait(3);
        Type("{ESC}");

        //Start Application
        //Log("Starting Word");
        Wait(seconds: 3, showOnScreen: true, onScreenText: "Starting Word");
        START(mainWindowTitle: "*Word*", mainWindowClass: "Win32 Window:OpusApp", processName: "WINWORD", timeout: 30);
        MainWindow.Maximize();
        var newDocName = "edited";
        var appWasLeftOpen = MainWindow.GetTitle().Contains(newDocName);
        if (appWasLeftOpen)
        {
            Log("Word was left open from previous run");
        }
        else
        {
            Wait(10);

            SkipFirstRunDialogs();
        }*/
        
        Wait(seconds: 1, showOnScreen: true, onScreenText: "Starting Excel");
        // START(mainWindowTitle: "*Excel*", mainWindowClass: "*XLMAIN*", timeout: 30);
        ShellExecute(@"winword.exe C:\Users\NVtestuser\Desktop\loginvsi.docx",forceKillOnExit:true,waitForProcessEnd:false);        
        //START(mainWindowTitle: "*Excel*", timeout: 30);
        var MainWindow = FindWindow(title:"*Word*");
        Wait(1);
        MainWindow.Maximize();
        MainWindow.Focus();
        //MainWindow.MoveMouseToCenter();
        Wait(1);
        //MainWindow.Click();
        MainWindow.Type("{alt}np",cpm:50);
        Wait(1);
        MainWindow.Type("d",cpm:50);
        Wait(2);
        Type(@"C:\Users\NVtestuser\Desktop\LoginVSI_BattlingRobots.bmp{enter}", cpm:1000);
        Wait(2);
        
        // Usage of Scroll():
//   - direction: "Down" to scroll down or "Up" to scroll up.
//   - scrollCount: Number of scroll events to send.
//   - notches: Number of notches per event (1 notch is typically 120).
//   - waitTime: Time in seconds to wait between each scroll event.
// Example:
//   Scroll("Down", 20, 1, 0.2);
//   Scroll("Up", 10, 2, 0.3);

// Scroll interactions on the active tab after switching:
MainWindow.MoveMouseToCenter();
MainWindow.Click();
Wait(1);
Type("{ctrl+a}{ctrl+c}{ctrl+v}{ctrl+v}",cpm:50);
Wait(3);
//MainWindow.FindControl(className : "Button:NetUISimpleButton", title : "Zoom *").Click();
//Wait(1);
//Type("{alt+2}",cpm:50);
//Wait(2);
//Type("{enter}",cpm:50);
//Wait(2);
Scroll("Up",50, 2, 0.1);
Scroll("Down", 50, 1, 0.1);


Wait(1);
        
        /*
        //Open "Open File" window and start measurement.
        Wait(seconds: 3, showOnScreen: true, onScreenText: "Open File Window");
        MainWindow.Type("{CTRL+O}");
        MainWindow.Type("{ALT+O+O}");
        StartTimer("Open_Window");
        var OpenWindow = get_file_dialog();

        // OpenWindow.FindControl(className : "SplitButton:Button", title : "&Open").Click();
        StopTimer("Open_Window");
        OpenWindow.Click();

        //Navigate to copied DOCX file and press Open, measure time to open the file.
        Wait(seconds: 3, showOnScreen: true, onScreenText: "Open File");
        var fileNameBox = OpenWindow.FindControl(className: "Edit:Edit", title: "File name:");
        fileNameBox.Click();
        Wait(1);
        ScriptHelpers.SetTextBoxText(this, fileNameBox, $"{temp}\\LoginPI\\loginvsi.docx", cpm: 300);
        Wait(1);
        OpenWindow.FindControl(className: "SplitButton:Button", title: "&Open").Click();
        StartTimer("Open_Word_Document");
        var newWord = FindWindow(className: "Win32 Window:OpusApp", title: "loginvsi*", processName: "WINWORD"); //FindControlWithXPath doens't work here
        newWord.Focus();
        //newWord.FindControl(className : "TabItem:NetUIRibbonTab", title : "Insert"); //this failed under load. The change doesn't throw off timing
        StopTimer("Open_Word_Document");

        if (appWasLeftOpen)
        {
            MainWindow.Close();
            Wait(1);
        }

        //Scroll through Word Document
        Wait(seconds: 3, showOnScreen: true, onScreenText: "Scroll");
        newWord.MoveMouseToCenter();
        MouseDown();
        Wait(1);
        MouseUp();
        newWord.Type("{PAGEDOWN}".Repeat(waitTime));
        Wait(1);
        newWord.Type("{PAGEUP}".Repeat(waitTime));

        //Type in the document (in the future create a txt file of content and type randomly from it)
        // newWord.Type("{CTRL+END}");
        Wait(seconds: 3, showOnScreen: true, onScreenText: "Type");
        newWord.Type("The snappy guy, who was a little rough around the edges, blossomed. The old fogey sat down in order to pass the time. ", cpm: 900);
        newWord.Type("The slippery townspeople had an unshakable fear of ostriches while encountering a whirling dervish. ", cpm: 900);
        newWord.Type("The prisoner stepped in a puddle while chasing the neighbor's cat out of the yard. The gal thought about mowing the lawn during a pie fight. ", cpm: 900);
        newWord.Type("A darn good bean-counter had a pen break while chewing on it while placing one ear to the ground. ", cpm: 900);
        Wait(1);

        newWord.Type("{ENTER}");
        newWord.Type("The intelligent baby felt sick after watching a silent film. As usual, the beekeeper spoke on a cellphone in nothing flat. ", cpm: 900);
        newWord.Type("A behemoth of a horde of morons committed a small crime and then chuckled arrogantly. The typical girl frequently wore a toga. ", cpm: 900);
        newWord.Type("The meowing guy, who had a little too much confidence in himself, threw a gutter ball in a rather graceful manner. ", cpm: 900);
        newWord.Type("The wicked Bridge Club shrugged both shoulders, which was considered a sign of great wisdom. ", cpm: 900);

        newWord.Minimize();
        Wait(2);
        newWord.Maximize();

        //Copy some text and paste it
        Wait(seconds: 3, showOnScreen: true, onScreenText: "Copy & Paste");
        KeyDown(KeyCode.SHIFT);
        Type("{UP}".Repeat(10));
        KeyUp(KeyCode.SHIFT);
        Wait(1);
        newWord.Type("{CTRL+C}");
        Wait(1);
        newWord.Type("{CTRL+V}");
        Wait(1);
        newWord.Type("{CTRL+V}");
        Wait(1);
        newWord.Type("{PAGEUP}");
        Wait(1);
        newWord.Type("{CTRL+V}");
        Wait(1);
        newWord.Type("{PAGEUP}");
        Wait(1);
        newWord.Type("{CTRL+V}");
        Wait(1);
        newWord.Type("{CTRL+V}");
        Wait(1);

        // Saving the file in temp
        Wait(seconds: 3, showOnScreen: true, onScreenText: "Saving File");
        newWord.Type("{F12}", cpm: 0);
        Wait(1);

        var filename = $"{temp}\\LoginPI\\{newDocName}.docx";
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

        var SaveAs = get_file_dialog();

        fileNameBox = SaveAs.FindControl(className: "Edit:Edit", title: "File name:");
        fileNameBox.Click();
        Wait(1);
        ScriptHelpers.SetTextBoxText(this, fileNameBox, filename, cpm: 300);
        StartTimer("Saving_file");
        SaveAs.Type("{ENTER}");
        FindWindow(title: $"{newDocName}*", processName: "WINWORD");
        StopTimer("Saving_file");
        Wait(2);

        // Stop application
        Wait(seconds: 3, showOnScreen: true, onScreenText: "Stopping App");
        Wait(2);*/
        //STOP();
        MainWindow.Close();

    }

    private void SkipFirstRunDialogs()
    {
        var dialog = FindWindow(className: "Win32 Window:NUIDialog", processName: "WINWORD", continueOnError: true, timeout: 1);
        while (dialog != null)
        {
            dialog.Close();
            dialog = FindWindow(className: "Win32 Window:NUIDialog", processName: "WINWORD", continueOnError: true, timeout: 10);
        }
    }

    private IWindow get_file_dialog()
    {
        var dialog = FindWindow(className: "Win32 Window:#32770", processName: "WINWORD", continueOnError: true, timeout:10);
        if (dialog is null)
        {
            ABORT("File dialog could not be found");
        }
        return dialog;
    }

    public static class ScriptHelpers
    {
        ///
        /// This method types the given text to the textbox (any existing text is cleared)
        /// After typing, it confirms the resulting value.
        /// If it does not match, it will clear the textbox and try again
        ///
        public static void SetTextBoxText(ScriptBase script, IWindow textBox, string text, int cpm = 800)
        {
            var numTries = 1;
            string currentText = null;
            do
            {
                textBox.Type("{CTRL+a}");
                script.Wait(0.5);
                textBox.Type(text, cpm: cpm);
                script.Wait(1);
                currentText = textBox.GetText();
                if (currentText != text)
                    script.CreateEvent($"Typing error in attempt {numTries}", $"Expected '{text}', got '{currentText}'");
            }
            while (++numTries < 5 && currentText != text);
            if (currentText != text)
                script.ABORT($"Unable to set the correct text '{text}', got '{currentText}'");
        }
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

