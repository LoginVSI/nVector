// TARGET:winword.exe /t
// START_IN:

/////////////
// Word Application
// Workload: KnowledgeWorker 2025
// Version: 1.0
/////////////

using LoginPI.Engine.ScriptBase;
using LoginPI.Engine.ScriptBase.Components;
using System;
using System.IO;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;
using System.Runtime.InteropServices;

public class WordDefaultScript : ScriptBase
{
    // =====================================================
    // Import and Constants
    // =====================================================
    [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
    public static extern void mouse_event(uint dwFlags, uint dx, uint dy, int dwData, UIntPtr dwExtraInfo);
    public const uint MOUSEEVENTF_WHEEL = 0x0800; // Constant for a mouse wheel event

    // =====================================================
    // Configurable Variables
    // =====================================================
    // Global timings and speeds
    int globalTimeoutInSeconds = 60;          // Timeout for actions (e.g., finding the app window)
    int waitMessageboxInSeconds = 2;          // Duration for onscreen wait messages (in seconds)
    double globalWaitInSeconds = 3;           // General wait time between actions for human-like behavior
    int keyboardShortcutsCPM = 30;            // Typing speed for keyboard shortcuts
    int typingTextCPM = 600;                  // Typing speed for document text
    int copyPasteRepetitions = 1;             // Number of times to perform the copy-paste action
    int waitForCopyPasteInSeconds = 5;        // Wait time after copy-paste actions

    // Scrolling parameters (for document navigation)
    int scrollDownCount = 40;                 // Number of scroll events for scrolling down
    int scrollUpCount = 40;                   // Number of scroll events for scrolling up
    double scrollWaitTime = 0.1;              // Wait time between each scroll event

    // File download settings
    private string bmpUrl = "https://myAppliance.myOrg.com/contentDelivery/content/LoginVSI_BattlingRobots.bmp"; // Replace with the actual URL for the BMP file

    // Document content to be typed into Word, broken into separate lines.
    string[] documentContentLines = new string[] 
    {
        "About Login VSI",
        "The VDI and DaaS industry has transformed incredibly, and Login VSI has evolved alongside the world of remote and hybrid work.",
        "Through an innovative and dynamic culture, the Login VSI team is passionate about helping enterprises worldwide understand, build, and maintain amazing digital workspaces.",
        "Trusted globally for 360° proactive visibility of performance, cost, and capacity of virtual desktops and applications, Login Enterprise is accepted as the industry standard and used by major vendors to spot problems quicker, avoid unexpected downtime, and deliver next-level digital experiences for end-users.",
        "Our Mission",
        "The paradigm for remote computing has shifted with virtual app delivery coupled with the growth in Web and SaaS-based applications.",
        "Now more than ever, organizations rely on digital workspaces to function. We give our customers 360° insights into the entire stack of virtual desktops and applications – in production or delivery and across various settings and infrastructure.",
        "We aim to empower IT teams to take control of their virtual desktops and applications’ performance, cost, and capacity wherever they reside – traditional, hybrid, or cloud."
    };

    private void Execute()
    {
        // =====================================================
        // Setup: Directory and File Downloads
        // =====================================================
        var temp = GetEnvironmentVariable("TEMP");
        string loginEnterpriseDir = $"{temp}\\LoginEnterprise";
        if (!Directory.Exists(loginEnterpriseDir))
        {
            Directory.CreateDirectory(loginEnterpriseDir);
            Log("Created directory: " + loginEnterpriseDir);
        }

        // --- Download .docx File ---
        Wait(seconds: waitMessageboxInSeconds, showOnScreen: true, onScreenText: "Retrieving .docx file");
        string docxFile = $"{loginEnterpriseDir}\\loginvsi.docx";
        Log("Downloading Word document file (force overwrite)");
        CopyFile(KnownFiles.WordDocument, docxFile, overwrite: true);

        // --- Download the BMP File ---
        string bmpFile = $"{loginEnterpriseDir}\\LoginVSI_BattlingRobots.bmp";
        if (!FileExists(bmpFile))
        {
            Log("Downloading BMP file");
            try
            {
                // Disable SSL certificate validation.
                ServicePointManager.ServerCertificateValidationCallback =
                    delegate (object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
                    {
                        return true;
                    };

                using (WebClient client = new WebClient())
                {
                    client.DownloadFile(bmpUrl, bmpFile);
                    Log("BMP file downloaded successfully to: " + bmpFile);
                }
            }
            catch (Exception ex)
            {
                ABORT("Error downloading BMP file: " + ex.Message);
            }
        }
        else
        {
            Log("BMP file already exists");
        }

        // =====================================================
        // Open/Close Start Menu
        // =====================================================
        Log("Opening Start Menu");
        Wait(seconds: waitMessageboxInSeconds, showOnScreen: true, onScreenText: "Start Menu");
        Type("{LWIN}", hideInLogging:false);
        Wait(seconds: 3);
        Type("{LWIN}", hideInLogging:false);
        Wait(seconds: 1);
        Type("{ESC}", hideInLogging:false);

        // =====================================================
        // Skip First-Run Dialogs before Bringing Word into Focus
        // =====================================================
        Wait(globalWaitInSeconds);
        SkipFirstRunDialogs();

        // =====================================================
        // Bring Word Application into Focus
        // =====================================================
        var MainWindow = FindWindow(className: "Win32 Window:OpusApp", title: "*Word*", processName: "WINWORD", timeout: globalTimeoutInSeconds);
        Wait(seconds: waitMessageboxInSeconds, showOnScreen: true, onScreenText: "Interacting with existing Word");
        MainWindow.Maximize();
        MainWindow.Focus();

        // =====================================================
        // Open .docx File via Open File Dialog
        // =====================================================
        Log("Opening .docx file via open file dialog");
        Wait(seconds: waitMessageboxInSeconds, showOnScreen: true, onScreenText: "Open .docx file");
        MainWindow.Type("{CTRL+O}{ALT+O+O}", cpm: keyboardShortcutsCPM, hideInLogging:false);
        StartTimer("Open_DOCX_Dialog");
        var openWindow = FindWindow(className: "Win32 Window:#32770", processName: "WINWORD", continueOnError:false, timeout: globalTimeoutInSeconds);
        StopTimer("Open_DOCX_Dialog");
        Wait(seconds: globalWaitInSeconds, showOnScreen: true, onScreenText: "Waiting for dialog...");
        var fileNameBox = openWindow.FindControl(className: "Edit:Edit", title: "File name:", timeout: globalTimeoutInSeconds);
        fileNameBox.Click();
        Wait(seconds: globalWaitInSeconds, showOnScreen: true, onScreenText: "Typing file path...");
        ScriptHelpers.SetTextBoxText(this, fileNameBox, docxFile, cpm: typingTextCPM);
        Type("{enter}", hideInLogging:false);
        StartTimer("Open_Word_Document");
        var newWord = FindWindow(className: "Win32 Window:OpusApp", title: "loginvsi*", processName: "WINWORD", timeout: globalTimeoutInSeconds);
        StopTimer("Open_Word_Document");
        newWord.Focus();
        newWord.Maximize();
        Wait(seconds: globalWaitInSeconds, showOnScreen: true, onScreenText: "Document loaded");

        // =====================================================
        // Skip First-Run Dialogs before Checking for an Existing Edited Window
        // =====================================================
        Wait(globalWaitInSeconds);
        SkipFirstRunDialogs();

        // =====================================================
        // Check and Close Existing Edited Window (if any)
        // =====================================================
        string newDocName = "edited";
        for (int attempt = 0; attempt < 2; attempt++)
        {
            var editedWindow = FindWindow(title: $"{newDocName}*", processName: "WINWORD", timeout: 2, continueOnError: true);
            if (editedWindow != null)
            {
                Wait(seconds: globalWaitInSeconds, showOnScreen: true, onScreenText: "Closing existing edited window");
                editedWindow.Focus();
                editedWindow.Maximize();
                Wait(seconds: globalWaitInSeconds);
                Log("Existing edited Word window found. Closing it.");
                editedWindow.Type("{ALT+F4}", hideInLogging: false);
                Wait(seconds: globalWaitInSeconds);
                newWord = FindWindow(className: "Win32 Window:OpusApp", title: "loginvsi*", processName: "WINWORD", timeout: globalTimeoutInSeconds);
                Wait(seconds: globalWaitInSeconds);
                newWord.Focus();
                newWord.Maximize();
                Wait(seconds: globalWaitInSeconds);
            }
        }
        
        Wait(seconds: waitMessageboxInSeconds, showOnScreen: true, onScreenText: "Format text editor pane, type to it, insert picture, copy and paste, and scroll");

        // =====================================================
        // Type Initial Commands into Word
        // =====================================================
        // Switch page view to "Page Width" and then to "One Page"
        newWord.Type("{ALT}wi", cpm: keyboardShortcutsCPM, hideInLogging:false);
        Wait(seconds: globalWaitInSeconds, showOnScreen: true, onScreenText: "Switching view...");
        newWord.Type("{ALT}w1", cpm: keyboardShortcutsCPM, hideInLogging:false);
        Wait(seconds: globalWaitInSeconds);

        // =====================================================
        // Type Document Content
        // =====================================================
        // Type each line separately with an {Enter} keypress after each line.
        foreach (var line in documentContentLines)
        {
            newWord.Type(line, cpm: typingTextCPM, hideInLogging:false);
            newWord.Type("{Enter}", cpm: typingTextCPM, hideInLogging:false);
        }
        Wait(seconds: globalWaitInSeconds, showOnScreen: true, onScreenText: "Content entered");

        // =====================================================
        // Insert BMP Image into Document
        // =====================================================
        Log("Inserting BMP image");
        newWord.Type("{ALT}np", cpm: keyboardShortcutsCPM, hideInLogging:false);
        Wait(seconds: globalWaitInSeconds, showOnScreen: true, onScreenText: "Preparing picture dialog");
        // 'd' opens the insert file dialog.
        Type("d", cpm: keyboardShortcutsCPM, hideInLogging:false);
        StartTimer("Insert_Picture_Dialog");
        var addPictureDialog = FindWindow(className: "Win32 Window:#32770", processName: "WINWORD", timeout: globalTimeoutInSeconds);
        StopTimer("Insert_Picture_Dialog");
        addPictureDialog.Focus();
        var fileNameBoxPic = addPictureDialog.FindControl(className: "Edit:Edit", title: "File name:", timeout: globalTimeoutInSeconds);
        Wait(seconds: globalWaitInSeconds, showOnScreen: true, onScreenText: "Typing BMP file path");
        fileNameBoxPic.Click();
        ScriptHelpers.SetTextBoxText(this, fileNameBoxPic, bmpFile, cpm: typingTextCPM);
        fileNameBoxPic.Type("{ENTER}", cpm: keyboardShortcutsCPM, hideInLogging:false);
        Wait(seconds: globalWaitInSeconds, showOnScreen: true, onScreenText: "Picture inserted");

        // =====================================================
        // Copy & Paste Operations
        // =====================================================
        Log("Performing copy and paste operations");
        newWord.Type("{CTRL+a}", cpm: keyboardShortcutsCPM, hideInLogging:false);
        Wait(seconds: globalWaitInSeconds, showOnScreen: true, onScreenText: "Selecting content");
        newWord.Type("{CTRL+c}", cpm: keyboardShortcutsCPM, hideInLogging:false);
        Wait(seconds: globalWaitInSeconds, showOnScreen: true, onScreenText: "Content copied");
        for (int i = 0; i < copyPasteRepetitions; i++)
        {
            newWord.Type("{CTRL+V}", cpm: keyboardShortcutsCPM, hideInLogging:false);
            Wait(seconds: waitForCopyPasteInSeconds, showOnScreen: true, onScreenText: "Pasting content");
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
        Log("Scrolling the document");
        newWord.Type("{CTRL+HOME}", hideInLogging:false); // Go to the top of the document
        Wait(seconds: waitForCopyPasteInSeconds, showOnScreen: true, onScreenText: "Scrolling...");
        Scroll("Down", scrollDownCount, 1, scrollWaitTime);
        Scroll("Up", scrollUpCount, 1, scrollWaitTime);
        newWord.Type("{CTRL+HOME}", hideInLogging:false); // Return to top
        Wait(seconds: waitForCopyPasteInSeconds, showOnScreen: true, onScreenText: "Scrolling...");
        Scroll("Down", scrollDownCount, 1, scrollWaitTime);
        Scroll("Up", scrollUpCount, 1, scrollWaitTime);
        newWord.Type("{CTRL+HOME}", hideInLogging:false); // Return to top
        Wait(seconds: globalWaitInSeconds);

        // =====================================================
        // Minimize and Maximize Word Window
        // =====================================================
        Log("Minimizing and maximizing Word window");
        newWord.Minimize();
        Wait(seconds: globalWaitInSeconds, showOnScreen: true, onScreenText: "Minimizing...");
        newWord.Maximize();
        newWord.Focus();
        Wait(seconds: globalWaitInSeconds, showOnScreen: true, onScreenText: "Restoring window");

        // =====================================================
        // Save the Edited Document
        // =====================================================
        Log("Saving the edited Word document");
        Wait(seconds: waitMessageboxInSeconds, showOnScreen: true, onScreenText: "Save the document");
        string saveFilename = $"{loginEnterpriseDir}\\{newDocName}.docx";
        if (FileExists(saveFilename))
        {
            Log("Removing existing file: " + saveFilename);
            RemoveFile(saveFilename);
        }
        else
        {
            Log("No existing file to remove at: " + saveFilename);
        }
        newWord.Type("{F12}", cpm: keyboardShortcutsCPM, hideInLogging:false);
        StartTimer("Save_As_Dialog");
        var saveAs = FindWindow(className: "Win32 Window:#32770", processName: "WINWORD", continueOnError:true, timeout: globalTimeoutInSeconds);
        var saveFileNameBox = saveAs.FindControl(className: "Edit:Edit", title: "File name:", timeout: globalTimeoutInSeconds);
        StopTimer("Save_As_Dialog");
        saveFileNameBox.Click();
        Wait(seconds: globalWaitInSeconds, showOnScreen: true, onScreenText: "Typing save path");
        ScriptHelpers.SetTextBoxText(this, saveFileNameBox, saveFilename, cpm: typingTextCPM);
        saveAs.Type("{ENTER}", hideInLogging:false);
        StartTimer("Saving_file");
        FindWindow(title: $"{newDocName}*", processName: "WINWORD", timeout: globalTimeoutInSeconds);
        StopTimer("Saving_file");
        Wait(seconds: globalWaitInSeconds, showOnScreen: true, onScreenText: "Finalizing...");

        Log("Script complete. Word remains open.");
    }

    // =====================================================
    // Helper: Scroll Function
    // =====================================================
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
            Wait(seconds: waitTime);
        }
    }

    // =====================================================
    // Helper: Skip First-Run Dialogs
    // =====================================================
    private void SkipFirstRunDialogs()
    {
        int loopCount = 2; // configurable number of loops
        for (int i = 0; i < loopCount; i++)
        {
            var dialog = FindWindow(
                className: "Win32 Window:NUIDialog", 
                processName: "WINWORD", 
                continueOnError: true, 
                timeout: 5);
            while (dialog != null)
            {
                dialog.Close();
                dialog = FindWindow(
                    className: "Win32 Window:NUIDialog", 
                    processName: "WINWORD", 
                    continueOnError: true, 
                    timeout: 5);
            }
        }
    }
}

// =====================================================
// Helper Class for TextBox Operations
// =====================================================
public static class ScriptHelpers
{
    public static void SetTextBoxText(ScriptBase script, IWindow textBox, string text, int cpm = 600)
    {
        double globalWaitInSeconds = 3;           // General wait time between actions for human-like behavior
        var numTries = 1;
        string currentText = null;
        do
        {
            textBox.Type("{CTRL+a}", hideInLogging:false);
            script.Wait(globalWaitInSeconds);
            textBox.Type(text, cpm: cpm, hideInLogging:false);
            script.Wait(globalWaitInSeconds);
            currentText = textBox.GetText();
            if (currentText != text)
                script.CreateEvent($"Typing error in attempt {numTries}", $"Expected '{text}', got '{currentText}'");
        }
        while (++numTries < 5 && currentText != text);
        if (currentText != text)
            script.ABORT($"Unable to set the correct text '{text}', got '{currentText}'");
    }
}
