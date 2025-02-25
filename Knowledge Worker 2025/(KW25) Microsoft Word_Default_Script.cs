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

public class M365Word_RefactoredScript : ScriptBase
{
    // Import the user32.dll function to simulate mouse events.
    [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
    public static extern void mouse_event(uint dwFlags, uint dx, uint dy, int dwData, UIntPtr dwExtraInfo);
    public const uint MOUSEEVENTF_WHEEL = 0x0800; // Constant for a mouse wheel event

    // BMP URL and file download settings
    private string bmpUrl = "<your URL here>"; // Replace with the actual URL for the BMP file, such as "https://myAppliance.myOrg.com/contentDelivery/content/LoginVSI_BattlingRobots.bmp";
    
    private void Execute()
    {
        // Global settings
        int globalTimeoutInSeconds = 60;                // Timeout for actions (e.g., finding the app window)
        int waitMessageboxInSeconds = 3;                // Duration for wait message boxes (in seconds)
        double globalWaitInSeconds = 3;                 // General wait time between actions for human-like behavior
        int keyboardShortcutsCharactersPerMinuteToType = 50; // Typing speed for keyboard shortcuts
        int CharactersPerMinuteToType = 600;           // Typing speed for text on the document canvas
        int copyPasteRepetitions = 1;                   // Number of times to perform the copy-paste action
        int waitForCopyPasteInSeconds = 5;              // After the copy and paste keyboard shortcuts how long to wait for the paste to finalize
    
        // Get the TEMP directory and ensure the LoginEnterprise folder exists.
        var temp = GetEnvironmentVariable("TEMP");
        string loginEnterpriseDir = $"{temp}\\LoginEnterprise";
        if (!Directory.Exists(loginEnterpriseDir))
        {
            Directory.CreateDirectory(loginEnterpriseDir);
            Log("Created directory: " + loginEnterpriseDir);
        }

        // --- Download .docx file ---
        Wait(seconds: waitMessageboxInSeconds, showOnScreen: true, onScreenText: "Get .docx file");
        string docxFile = $"{loginEnterpriseDir}\\loginvsi.docx";
        if (!FileExists(docxFile))
        {
            Log("Downloading Word document file");
            CopyFile(KnownFiles.WordDocument, docxFile);
        }
        else
        {
            Log("Word document file already exists");
        }

        // --- Download the BMP file ---
        string bmpFile = $"{loginEnterpriseDir}\\LoginVSI_BattlingRobots.bmp";
        if (!FileExists(bmpFile))
        {
            Log("Downloading BMP file");
            try
            {
                // Disable SSL certificate validation.
                ServicePointManager.ServerCertificateValidationCallback =
                    delegate(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
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

        /*
        // --- Open/Close Start Menu ---
        Log("Opening Start Menu");
        Wait(seconds: 2, showOnScreen: true, onScreenText: "Start Menu");
        Type("{LWIN}");
        Wait(3);
        Type("{LWIN}");
        Wait(1);
        Type("{ESC}");
        */

        // --- Locate the already open Word Application ---        
        // Find the main Word window (assuming it's already open)
        var MainWindow = FindWindow(className:"Win32 Window:OpusApp", title:"*Word*", processName:"WINWORD", timeout:globalTimeoutInSeconds);
        Wait(seconds: waitMessageboxInSeconds, showOnScreen: true, onScreenText: "Interacting with existing Word");
        Wait(globalWaitInSeconds);
        MainWindow.Maximize();
        MainWindow.Focus();

        // --- Open File Dialog to open the .docx (if needed) ---
        Log("Opening .docx file via open file dialog");
        Wait(seconds: waitMessageboxInSeconds, showOnScreen: true, onScreenText: "Open .docx file");
        MainWindow.Type("{CTRL+O}{ALT+O+O}", cpm: keyboardShortcutsCharactersPerMinuteToType);
        StartTimer("Open_DOCX_Dialog");
        var openWindow = FindWindow(className:"Win32 Window:#32770", processName:"WINWORD", continueOnError:false, timeout:globalTimeoutInSeconds);
        StopTimer("Open_DOCX_Dialog");
        Wait(globalWaitInSeconds);
        var fileNameBox = openWindow.FindControl(className:"Edit:Edit", title:"File name:", timeout:globalTimeoutInSeconds);
        fileNameBox.Click();
        Wait(globalWaitInSeconds);
        ScriptHelpers.SetTextBoxText(this, fileNameBox, docxFile, cpm:1000);
        Type("{enter}");
        StartTimer("Open_Word_Document");
        var newWord = FindWindow(className:"Win32 Window:OpusApp", title:"loginvsi*", processName:"WINWORD", timeout:globalTimeoutInSeconds);        
        StopTimer("Open_Word_Document");
        newWord.Focus();
        Wait(globalWaitInSeconds);

        // --- Check for an existing "edited" Word window ---
        string newDocName = "edited";
        var editedWindow = FindWindow(title:$"{newDocName}*", processName:"WINWORD", timeout:5, continueOnError:true);
        if (editedWindow != null)
        {
            Wait(globalWaitInSeconds);
            editedWindow.Focus();
            Wait(globalWaitInSeconds);
            Log("Existing edited Word window found. Closing it.");
            editedWindow.Close();
            Wait(globalWaitInSeconds);
            newWord = FindWindow(className:"Win32 Window:OpusApp", title:"loginvsi*", processName:"WINWORD", timeout:globalTimeoutInSeconds);
            Wait(globalWaitInSeconds);
            newWord.Focus();
            Wait(globalWaitInSeconds);
        }

        // --- Type initial commands into Word ---
        // Switch page view to "Page Width" using keyboard shortcuts.
        newWord.Type("{ALT}wi", cpm: keyboardShortcutsCharactersPerMinuteToType);
        Wait(globalWaitInSeconds);
        // Switch page view to "One Page"
        newWord.Type("{ALT}w1", cpm: keyboardShortcutsCharactersPerMinuteToType);
        Wait(globalWaitInSeconds);

        // --- Type the document content ---
        string content =
@"About Login VSI The VDI and DaaS industry has transformed incredibly, and Login VSI has evolved alongside the world of remote and hybrid work. Through an innovative and dynamic culture, the Login VSI team is passionate about helping enterprises worldwide understand, build, and maintain amazing digital workspaces. Trusted globally for 360° proactive visibility of performance, cost, and capacity of virtual desktops and applications, Login Enterprise is accepted as the industry standard and used by major vendors to spot problems quicker, avoid unexpected downtime, and deliver next-level digital experiences for end-users. Our Mission The paradigm for remote computing has shifted with virtual app delivery coupled with the growth in Web and SaaS-based applications. Now more than ever, organizations rely on digital workspaces to function. We give our customers 360° insights into the entire stack of virtual desktops and applications – in production or delivery and across various settings and infrastructure. We aim to empower IT teams to take control of their virtual desktops and applications’ performance, cost, and capacity wherever they reside – traditional, hybrid, or cloud.";
        // Type text at the specified text speed.
        newWord.Type(content, cpm: CharactersPerMinuteToType);
        Wait(globalWaitInSeconds);

        // --- Insert BMP image into the document ---
        Log("Inserting BMP image");
        newWord.Type("{ALT}np", cpm: keyboardShortcutsCharactersPerMinuteToType);
        Wait(globalWaitInSeconds);
        // 'd' opens the insert file dialog.
        Type("d", cpm: keyboardShortcutsCharactersPerMinuteToType);
        StartTimer("Insert_Picture_Dialog");
        var addPictureDialog = FindWindow(className:"Win32 Window:#32770", processName:"WINWORD", timeout:globalTimeoutInSeconds);
        StopTimer("Insert_Picture_Dialog");
        addPictureDialog.Focus();
        var fileNameBoxPic = addPictureDialog.FindControl(className:"Edit:Edit", title:"File name:", timeout:globalTimeoutInSeconds);
        Wait(globalWaitInSeconds);
        fileNameBoxPic.Click();
        ScriptHelpers.SetTextBoxText(this, fileNameBoxPic, bmpFile, cpm:1000);
        fileNameBoxPic.Type("{ENTER}", cpm: keyboardShortcutsCharactersPerMinuteToType);
        Wait(globalWaitInSeconds);

        // --- Copy & Paste operations ---
        Log("Performing copy and paste operations");
        newWord.Type("{CTRL+a}", cpm: keyboardShortcutsCharactersPerMinuteToType);
        Wait(globalWaitInSeconds);
        newWord.Type("{CTRL+c}", cpm: keyboardShortcutsCharactersPerMinuteToType);
        Wait(globalWaitInSeconds);
        for (int i = 0; i < copyPasteRepetitions; i++)
        {
            newWord.Type("{CTRL+V}", cpm: keyboardShortcutsCharactersPerMinuteToType);
            Wait(waitForCopyPasteInSeconds);
            newWord.Type("{ctrl+home}"); // Go to top of the doc
        }
        
        // Scroll interactions on the active tab:
        // Usage of Scroll():
        //   - direction: "Down" to scroll down or "Up" to scroll up.
        //   - scrollCount: Number of scroll events to send.
        //   - notches: Number of notches per event (1 notch is typically 120).
        //   - waitTime: Time in seconds to wait between each scroll event.
        // Example:
        //   Scroll("Down", 20, 1, 0.2);
        //   Scroll("Up", 10, 2, 0.3);
        // --- Scroll interactions using the helper Scroll() function ---
        Log("Scrolling the document");
        Wait(waitForCopyPasteInSeconds);
        newWord.Type("{ctrl+home}"); // Go to top of the doc
        Scroll("Down", 20, 1, 0.2);  
        newWord.Type("{ctrl+home}"); // Go to top of the doc
        Scroll("Up", 20, 1, 0.2);
        newWord.Type("{ctrl+home}"); // Go to top of the doc
        Scroll("Down", 20, 1, 0.2);
        Scroll("Up", 20, 1, 0.2);  
        newWord.Type("{ctrl+home}"); // Go to top of the doc
        Wait(globalWaitInSeconds);

        // --- Simulate minimizing and maximizing the Word window ---
        Log("Minimizing and maximizing Word window");
        newWord.Minimize();
        Wait(globalWaitInSeconds);
        newWord.Maximize();
        newWord.Focus();
        Wait(globalWaitInSeconds);

        // --- Save the Edited Document ---
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
        newWord.Type("{F12}", cpm: keyboardShortcutsCharactersPerMinuteToType);
        StartTimer("Save_As_Dialog");
        var saveAs = FindWindow(className:"Win32 Window:#32770", processName:"WINWORD", continueOnError:true, timeout:globalTimeoutInSeconds);
        var saveFileNameBox = saveAs.FindControl(className:"Edit:Edit", title:"File name:", timeout:globalTimeoutInSeconds);
        StopTimer("Save_As_Dialog");
        saveFileNameBox.Click();
        Wait(globalWaitInSeconds);
        ScriptHelpers.SetTextBoxText(this, saveFileNameBox, saveFilename, cpm:1000);
        saveAs.Type("{ENTER}");
        StartTimer("Saving_file");
        FindWindow(title:$"{newDocName}*", processName:"WINWORD", timeout:globalTimeoutInSeconds);
        StopTimer("Saving_file");
        Wait(globalWaitInSeconds);

        Log("Script complete. Word remains open.");        
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

    // Helper class for setting text in text boxes.
    public static class ScriptHelpers
    {
        public static void SetTextBoxText(ScriptBase script, IWindow textBox, string text, int cpm = 1000)
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
}
