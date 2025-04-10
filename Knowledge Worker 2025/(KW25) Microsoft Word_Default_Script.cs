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
using System.Diagnostics;

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
    int globalTimeoutInSeconds = 60;          // Timeout for actions (e.g., finding the app window)
    int waitMessageboxInSeconds = 2;          // Duration for onscreen wait messages (in seconds)
    double globalWaitInSeconds = 3;           // Standard wait time between actions
    int keyboardShortcutsCPM = 15;            // Typing speed for keyboard shortcuts
    int waitInBetweenKeyboardShortcuts = 3;   // Wait time between keyboard shortcuts
    int typingTextCPM = 600;                  // Typing speed for document text
    int copyPasteRepetitions = 1;             // Number of times to perform the copy-paste action
    int waitForCopyPasteInSeconds = 5;        // Wait time after copy-paste actions
    int startMenuWaitInSeconds = 5;           // Duration for Start Menu wait

    // File download for BMP remains in run script (if needed)
    private string bmpUrl = "https://myAppliance.myOrg.com/contentDelivery/content/LoginVSI_BattlingRobots.bmp"; // Replace with your actual URL

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
        // ----- (File download for docx is handled by the Start script) -----
        var temp = GetEnvironmentVariable("TEMP");
        string loginEnterpriseDir = $"{temp}\\LoginEnterprise";
        string docxFile = $"{loginEnterpriseDir}\\loginvsi.docx";

        // ----- Download BMP if needed -----
        string bmpFile = $"{loginEnterpriseDir}\\LoginVSI_BattlingRobots.bmp";
        if (!FileExists(bmpFile))
        {
            Log("Downloading BMP file");
            try
            {
                ServicePointManager.ServerCertificateValidationCallback =
                    delegate(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) 
                    { return true; };
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
        // Launch new blank Word document
        // =====================================================
        ShellExecute("cmd /c start \"\" winword /t", waitForProcessEnd: true, timeout: 3, continueOnError: false);
        /* Alternate start blank word document function:
        try
        {
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = "winword.exe",
                Arguments = "/t",
                UseShellExecute = true
            };
            Process.Start(startInfo);
        }
        catch (Exception ex)
        {
            ABORT("Error starting process: " + ex.Message);
        }
        */

        var newExcelWindow = FindWindow(title:"*Document*Word*", processName:"WINWORD", continueOnError:false, timeout:globalTimeoutInSeconds);
        Wait(globalWaitInSeconds);

        // =====================================================
        // Close any extraneous Word windows ("*loginvsi*" or "*edited*")
        // =====================================================
        CloseExtraWindows("WINWORD", "*loginvsi*");
        CloseExtraWindows("WINWORD", "*edited*");

        // =====================================================
        // Call new method to close a specific Microsoft Word confirmation window
        // =====================================================
        CloseMicrosoftWordDialog();

        // =====================================================
        // Skip First-Run Dialogs before Bringing Word into Focus
        // =====================================================
        Wait(globalWaitInSeconds);
        SkipFirstRunDialogs();

        // =====================================================
        // Bring Word Application into Focus
        // =====================================================
        var MainWindow = FindWindow(className:"Win32 Window:OpusApp", title:"*Document*", processName:"WINWORD", timeout:globalTimeoutInSeconds);
        Wait(waitMessageboxInSeconds, true, "Interacting with existing Word");
        MainWindow.Maximize();
        MainWindow.Focus();
        Wait(globalWaitInSeconds);

        // =====================================================
        // Open .docx File via Open File Dialog using MainWindow
        // =====================================================
        Log("Opening .docx file via open file dialog");
        Wait(waitMessageboxInSeconds, true, "Open .docx file");
        MainWindow.Type("{CTRL+O}", cpm: keyboardShortcutsCPM, hideInLogging:false);
        Wait(waitInBetweenKeyboardShortcuts);
        MainWindow.Type("{ALT+O+O}", cpm: keyboardShortcutsCPM, hideInLogging:false);
        StartTimer("Open_DOCX_Dialog");
        var openWindow = FindWindow(className:"Win32 Window:#32770", processName:"WINWORD", continueOnError:false, timeout:globalTimeoutInSeconds);
        StopTimer("Open_DOCX_Dialog");
        Wait(globalWaitInSeconds);
        var fileNameBox = openWindow.FindControl(className:"Edit:Edit", title:"File name:", timeout:globalTimeoutInSeconds);
        Wait(globalWaitInSeconds);
        openWindow.Type("{ALT+N}", hideInLogging:false);
        fileNameBox.Click();
        Wait(globalWaitInSeconds);
        ScriptHelpers.SetTextBoxText(this, fileNameBox, docxFile, cpm: typingTextCPM);
        Type("{ENTER}", hideInLogging:false);
        StartTimer("Open_Word_Document");
        var newWord = FindWindow(className:"Win32 Window:OpusApp", title:"loginvsi*", processName:"WINWORD", timeout:globalTimeoutInSeconds);
        StopTimer("Open_Word_Document");

        // Close any stray "Document" windows
        CloseExtraWindows("WINWORD", "*Document*");
        Wait(globalWaitInSeconds);
        newWord.Focus();
        newWord.Maximize();

        // =====================================================
        // Skip First-Run Dialogs before Checking for an Existing Edited Window
        // =====================================================
        Wait(globalWaitInSeconds);
        SkipFirstRunDialogs();

        // =====================================================
        // Typing content, inserting BMP, scrolling, Save As, etc.
        // =====================================================
        // Type initial commands
        newWord.Type("{ALT}wi", cpm: keyboardShortcutsCPM, hideInLogging:false);
        Wait(waitInBetweenKeyboardShortcuts);
        newWord.Type("{ALT}w1", cpm: keyboardShortcutsCPM, hideInLogging:false);
        Wait(waitInBetweenKeyboardShortcuts);
        foreach (var line in documentContentLines)
        {
            newWord.Type(line, cpm: typingTextCPM, hideInLogging:false);
            newWord.Type("{Enter}", cpm: typingTextCPM, hideInLogging:false);
        }
        Wait(globalWaitInSeconds);

        // Insert BMP image into document
        Log("Inserting BMP image");
        newWord.Type("{ALT}np", cpm: keyboardShortcutsCPM, hideInLogging:false);
        Wait(waitInBetweenKeyboardShortcuts);
        Type("d", cpm: keyboardShortcutsCPM, hideInLogging:false); // Opens insert file dialog
        StartTimer("Insert_Picture_Dialog");
        var addPictureDialog = FindWindow(className:"Win32 Window:#32770", processName:"WINWORD", timeout:globalTimeoutInSeconds);
        StopTimer("Insert_Picture_Dialog");
        addPictureDialog.Focus();
        var fileNameBoxPic = addPictureDialog.FindControl(className:"Edit:Edit", title:"File name:", timeout:globalTimeoutInSeconds);
        Wait(globalWaitInSeconds);
        addPictureDialog.Type("{ALT+N}", hideInLogging:false);
        fileNameBoxPic.Click();
        ScriptHelpers.SetTextBoxText(this, fileNameBoxPic, bmpFile, cpm: typingTextCPM);
        fileNameBoxPic.Type("{ENTER}", cpm: keyboardShortcutsCPM, hideInLogging:false);
        Wait(waitInBetweenKeyboardShortcuts, true, "Picture inserted");

        // Copy & paste operations
        Log("Performing copy and paste operations");
        Wait(globalWaitInSeconds);
        newWord.Type("{CTRL+a}{CTRL+c}", cpm: keyboardShortcutsCPM, hideInLogging:false);
        Wait(waitInBetweenKeyboardShortcuts);
        for (int i = 0; i < copyPasteRepetitions; i++)
        {
            newWord.Type("{CTRL+V}", cpm: keyboardShortcutsCPM, hideInLogging:false);
            Wait(waitForCopyPasteInSeconds, true, "Pasting content");
        }

        // Scroll function
        Log("Scrolling the document");
        newWord.Type("{CTRL+HOME}", hideInLogging:false);
        Wait(waitForCopyPasteInSeconds, true, "Scrolling...");
        Scroll("Down", 40, 1, 0.1);
        Scroll("Up", 40, 1, 0.1);
        newWord.Type("{CTRL+HOME}", hideInLogging:false);
        Wait(waitForCopyPasteInSeconds, true, "Scrolling...");
        Scroll("Down", 40, 1, 0.1);
        Scroll("Up", 40, 1, 0.1);
        newWord.Type("{CTRL+HOME}", hideInLogging:false);
        Wait(globalWaitInSeconds);

        // Minimize/Maximize and Save As
        Log("Minimizing and maximizing Word window");
        newWord.Minimize();
        Wait(globalWaitInSeconds);
        newWord.Maximize();
        newWord.Focus();
        Wait(globalWaitInSeconds);

        Log("Saving the edited Word document");
        Wait(waitMessageboxInSeconds, true, "Save the document");
        string saveFilename = $"{loginEnterpriseDir}\\edited.docx";
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
        var saveAs = FindWindow(className:"Win32 Window:#32770", processName:"WINWORD", continueOnError:true, timeout:globalTimeoutInSeconds);
        var saveFileNameBox = saveAs.FindControl(className:"Edit:Edit", title:"File name:", timeout:globalTimeoutInSeconds);
        Wait(globalWaitInSeconds);
        StopTimer("Save_As_Dialog");
        saveFileNameBox.Click();
        saveAs.Type("{ALT+N}", hideInLogging:false);
        ScriptHelpers.SetTextBoxText(this, saveFileNameBox, saveFilename, cpm: typingTextCPM);
        saveAs.Type("{ENTER}", hideInLogging:false);
        StartTimer("Saving_file");
        FindWindow(title:"edited*", processName:"WINWORD", timeout:globalTimeoutInSeconds);
        StopTimer("Saving_file");
        Wait(globalWaitInSeconds);

        Log("Script complete. Word remains open.");
    }

    void CloseExtraWindows(string processName, string titleMask)
    {
        int timeoutSeconds = 2;      // Timeout for re-checking the window
        int maxAttempts = 1;         // Maximum number of attempts to close the window

        for (int attempt = 0; attempt < maxAttempts; attempt++)
        {
            var extraWindow = FindWindow(title: titleMask, processName: processName, timeout: 3, continueOnError: true);
            if (extraWindow == null)
            {
                // Window is already closed
                break;
            }

            Wait(globalWaitInSeconds);
            extraWindow.Focus();
            extraWindow.Maximize();
            Wait(globalWaitInSeconds);
            extraWindow.Type("{ESC}", hideInLogging: false);
            Wait(globalWaitInSeconds);
            extraWindow.Type("{ALT+F4}", hideInLogging: false);
            Wait(globalWaitInSeconds);

            // Check if the window is still present
            var checkWindow = FindWindow(title: titleMask, processName: processName, timeout: timeoutSeconds, continueOnError: true);
            if (checkWindow != null)
            {
                Wait(globalWaitInSeconds);
                checkWindow.Type("{ALT+N}", hideInLogging: false);
                Wait(globalWaitInSeconds);
            }
        }
    }

    // Handling a specific Microsoft Word confirmation dialog: "Do you want to keep the last item you copied?"
    void CloseMicrosoftWordDialog()
    {
        int timeoutSeconds = 2; 
        var msWordWindow = FindWindow(className: "Win32 Window:#32770", title: "Microsoft Word", processName: "WINWORD", timeout: timeoutSeconds, continueOnError: true);
        if (msWordWindow != null)
        {
            msWordWindow.Focus();
            msWordWindow.Maximize();
            Wait(globalWaitInSeconds);
            msWordWindow.Type("{ALT+N}", hideInLogging: false);
            Wait(globalWaitInSeconds);
        }
    }

    private void SkipFirstRunDialogs()
    {
        int loopCount = 2;
        for (int i = 0; i < loopCount; i++)
        {
            var dialog = FindWindow(className:"Win32 Window:NUIDialog", processName:"WINWORD", continueOnError:true, timeout:3);
            while (dialog != null)
            {
                Wait(globalWaitInSeconds);
                dialog.Close();
                dialog = FindWindow(className:"Win32 Window:NUIDialog", processName:"WINWORD", continueOnError:true, timeout:3);
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
