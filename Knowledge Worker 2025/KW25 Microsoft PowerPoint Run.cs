// TARGET:powerpnt.exe /n
// START_IN:

/////////////
// PowerPoint Run
// Workload: Knowledge Worker 2025
// Version: 0.1.0
/////////////

using LoginPI.Engine.ScriptBase;
using LoginPI.Engine.ScriptBase.Components;
using System;
using System.IO;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;
using System.Diagnostics;

public class PowerPoint_Run : ScriptBase
{
    // =====================================================
    // Configurable Variables
    // =====================================================
    int globalTimeoutInSeconds = 60;
    int globalWaitInSeconds = 3;
    int waitMessageboxInSeconds = 2;
    int charactersPerMinuteToType = 15;
    int waitInBetweenKeyboardShortcuts = 4;
    int slideshowCharactersPerMinuteToType = 12;
    int pageScrollCpm = 60;
    int transitionPopupCharactersPerMinuteToType = 60;
    int waitForPopupShowingInSeconds = 10;
    int waitSlideshowStart = 10;
    int typingTextCharacterPerMinute = 600;
    int startMenuWaitInSeconds = 5;

    string bmpUrl = "https://myAppliance.myOrg.com/contentDelivery/content/LoginVSI_BattlingRobots.bmp"; // Replace with your actual URL

    private void Execute()
    {   
        // ----- (File download for pptx is handled by the Start script) -----
        var temp = GetEnvironmentVariable("TEMP");
        string loginEnterpriseDir = $"{temp}\\LoginEnterprise";
        string pptxFile = $"{loginEnterpriseDir}\\loginvsi.pptx";

        // Download BMP file if needed
        string bmpFile = Path.Combine(loginEnterpriseDir, "LoginVSI_BattlingRobots.bmp");
        if (!FileExists(bmpFile))
        {
            Log("Downloading BMP file");
            try
            {
                ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
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
        // Launch new blank PowerPoint presentation 
        // =====================================================
        try
        {
            // ShellExecute("powerpnt /n", waitForProcessEnd: false, timeout: globalTimeoutInSeconds, continueOnError: true, forceKillOnExit: false);
            // Alternate start blank PowerPoint document function:
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = "powerpnt.exe",
                Arguments = "/n",
                UseShellExecute = true
            };
            Process.Start(startInfo);
        }
        catch (Exception ex)
        {
            ABORT("Error starting process: " + ex.Message);
        }
        
        Wait(globalWaitInSeconds);
        var newPowerpointWindow = FindWindow(title:"*Presentation*PowerPoint*", processName:"POWERPNT", continueOnError:false, timeout:globalTimeoutInSeconds);
        Wait(globalWaitInSeconds);

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
        // Close any extraneous PowerPoint windows ("*loginvsi*" or "*edited*")
        // =====================================================
        CloseExtraWindows("POWERPNT", "*loginvsi*");
        CloseExtraWindows("POWERPNT", "*edited*");

        // =====================================================
        // Skip First-Run Dialogs before Bringing PowerPoint into Focus
        // =====================================================
        Wait(globalWaitInSeconds);
        SkipFirstRunDialogs();

        // =====================================================
        // Bring new PowerPoint instance into focus and open file via dialog
        // =====================================================
        Wait(startMenuWaitInSeconds);
        newPowerpointWindow.Focus();
        newPowerpointWindow.Maximize();
        Log("Opening PPTX file via open file dialog");
        Wait(waitMessageboxInSeconds, true, "Open pptx file");
        newPowerpointWindow.Type("{CTRL+O}", cpm: charactersPerMinuteToType, hideInLogging:false);
        Wait(waitInBetweenKeyboardShortcuts);
        newPowerpointWindow.Type("{ALT+O+O}", cpm: charactersPerMinuteToType, hideInLogging:false);
        StartTimer("Open_Window");
        var openWindow = FindWindow(className:"Win32 Window:#32770", processName:"POWERPNT", continueOnError:false, timeout:globalTimeoutInSeconds);
        StopTimer("Open_Window");
        Wait(globalWaitInSeconds);
        var fileNameBox = openWindow.FindControl(className:"Edit:Edit", title:"File name:", timeout:globalTimeoutInSeconds);
        Wait(globalWaitInSeconds);
        openWindow.Type("{ALT+N}", hideInLogging:false);
        fileNameBox.Click();
        Wait(globalWaitInSeconds);
        SetTextBoxText(fileNameBox, pptxFile, cpm: typingTextCharacterPerMinute);
        Type("{ENTER}", hideInLogging:false);
        StartTimer("Open_Powerpoint_Document");
        var newPowerpoint = FindWindow(className:"Win32 Window:PPTFrameClass", title:"loginvsi*", processName:"POWERPNT", timeout:globalTimeoutInSeconds);
        StopTimer("Open_Powerpoint_Document");
        Wait(globalWaitInSeconds);
        newPowerpoint.Focus();
        newPowerpoint.Maximize();
        newPowerpoint.FindControl(className:"TabItem:NetUIRibbonTab", title:"Insert", timeout:globalTimeoutInSeconds);
        Wait(globalWaitInSeconds);

        // Close any stray "Presentation" windows
        CloseExtraWindows("POWERPNT", "*Presentation*");

        // =====================================================
        // Continue with the rest of the workflow (transitions, slideshow, Save As, etc.)
        // =====================================================
        // --- Add new slide ---
        newPowerpoint.Focus();
        newPowerpoint.Maximize();
        Wait(globalWaitInSeconds);
        newPowerpoint.Type("{CTRL+M}", cpm: charactersPerMinuteToType, hideInLogging:false);
        Wait(waitInBetweenKeyboardShortcuts);

        // --- Insert BMP into New Slide ---
        Log("Inserting BMP into new slide");
        newPowerpoint.Type("{ALT}NP1", cpm: charactersPerMinuteToType, hideInLogging:false);
        Wait(waitForPopupShowingInSeconds);
        Type("d", cpm: charactersPerMinuteToType, hideInLogging:false);
        StartTimer("Insert_Picture_Dialog");
        var addPictureDialog = FindWindow(className:"Win32 Window:#32770", processName:"POWERPNT", title:"Insert Picture", timeout:globalTimeoutInSeconds);
        StopTimer("Insert_Picture_Dialog");
        addPictureDialog.Focus();
        var fileNameBox2 = addPictureDialog.FindControl(className:"Edit:Edit", title:"File name:", timeout:globalTimeoutInSeconds);
        Wait(globalWaitInSeconds);
        fileNameBox2.Click();
        newPowerpoint.Type("{ALT+N}", cpm: charactersPerMinuteToType, hideInLogging:false);
        Wait(waitInBetweenKeyboardShortcuts);
        ScriptHelpers.SetTextBoxText(this, fileNameBox2, bmpFile, cpm: typingTextCharacterPerMinute);
        fileNameBox2.Type("{ENTER}", cpm: charactersPerMinuteToType, hideInLogging:false);
        Wait(globalWaitInSeconds);
        var stillExists = FindWindow(className:"Win32 Window:#32770", title:"Insert Picture", processName:"POWERPNT", timeout:2, continueOnError:true);
        if (stillExists != null)
        {
            newPowerpoint.Type("{ESC}", cpm: charactersPerMinuteToType, hideInLogging:false);
        }

        // --- Add 'Honeycomb' Transition ---
        Log("Adding 'Honeycomb' transition");
        newPowerpoint.Type("{ALT+K}{ALT+T}", cpm: charactersPerMinuteToType, hideInLogging:false);
        newPowerpoint.FindControl(title:"Honeycomb", timeout:globalTimeoutInSeconds);
        Wait(globalWaitInSeconds);
        newPowerpoint.Type("{DOWN}", cpm: transitionPopupCharactersPerMinuteToType, hideInLogging:false);
        newPowerpoint.Type("{RIGHT}".Repeat(16), cpm: transitionPopupCharactersPerMinuteToType, hideInLogging:false);
        Wait(waitInBetweenKeyboardShortcuts);
        newPowerpoint.Type("{ENTER}", hideInLogging:false);
        Wait(globalWaitInSeconds);

        // --- Add transition to all slides ---
        Log("Adding transition to all slides");
        newPowerpoint.Type("{ALT+K}{ALT+L}", hideInLogging:false);
        Wait(globalWaitInSeconds);
        
        // --- Scroll, Minimize, and Maximize Presentation ---
        Wait(waitMessageboxInSeconds, true, "Scroll, minimize, and maximize");
        Log("Scrolling through presentation");
        newPowerpoint.Type("{PAGEDOWN}".Repeat(10), cpm: pageScrollCpm, hideInLogging:false);
        Wait(globalWaitInSeconds);
        newPowerpoint.Type("{PAGEUP}".Repeat(10), cpm: pageScrollCpm, hideInLogging:false);
        Wait(globalWaitInSeconds);
        newPowerpoint.Minimize();
        Wait(globalWaitInSeconds);
        newPowerpoint.Maximize();
        newPowerpoint.Focus();
        Wait(globalWaitInSeconds);

        // --- Run the Slideshow ---
        Wait(waitMessageboxInSeconds, true, "Run the slideshow");
        Log("Starting slideshow");
        newPowerpoint.Type("{F5}", cpm: charactersPerMinuteToType, hideInLogging:false);
        Wait(waitSlideshowStart);
        newPowerpoint.Type("{DOWN}".Repeat(5), cpm: slideshowCharactersPerMinuteToType, hideInLogging:false);
        Wait(waitInBetweenKeyboardShortcuts);
        Type("{ESC}", hideInLogging:false);
        Wait(globalWaitInSeconds);
        Type("{HOME}", hideInLogging:false);
        Wait(globalWaitInSeconds);

        // --- Save the Edited Slideshow ---
        Log("Saving the edited slideshow");
        Wait(waitMessageboxInSeconds, true, "Save the slideshow");
        string saveFilename = $"{loginEnterpriseDir}\\edited.pptx";
        if (FileExists(saveFilename))
        {
            Log("Removing existing file: " + saveFilename);
            RemoveFile(saveFilename);
        }
        else
        {
            Log("No existing file to remove at: " + saveFilename);
        }
        newPowerpoint.Type("{F12}", hideInLogging:false);
        StartTimer("Save_As_Dialog");
        var saveAs = FindWindow(className:"Win32 Window:#32770", processName:"POWERPNT", timeout:globalTimeoutInSeconds);
        StopTimer("Save_As_Dialog");
        var saveFileNameBox = saveAs.FindControl(className:"Edit:Edit", title:"File name:", timeout:globalTimeoutInSeconds);
        Wait(globalWaitInSeconds);
        Type("{ALT+N}", hideInLogging:false);
        saveFileNameBox.Click();
        Wait(globalWaitInSeconds);
        ScriptHelpers.SetTextBoxText(this, saveFileNameBox, saveFilename, cpm: typingTextCharacterPerMinute);
        saveAs.Type("{ENTER}", hideInLogging:false);
        StartTimer("Saving_file");
        FindWindow(title:"edited*", processName:"POWERPNT", timeout:globalTimeoutInSeconds);
        StopTimer("Saving_file");
        Wait(globalWaitInSeconds);

        Log("Script complete. PowerPoint remains open.");
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
    
    private void SkipFirstRunDialogs()
    {
        int loopCount = 2;
        for (int i = 0; i < loopCount; i++)
        {
            var dialog = FindWindow(className:"Win32 Window:NUIDialog", processName:"POWERPNT", continueOnError:true, timeout:3);
            while (dialog != null)
            {
                Wait(globalWaitInSeconds);
                dialog.Close();
                dialog = FindWindow(className:"Win32 Window:NUIDialog", processName:"POWERPNT", continueOnError:true, timeout:3);
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
