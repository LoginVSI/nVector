// TARGET:powerpnt.exe /n
// START_IN:

/////////////
// PowerPoint Application
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

public class PowerPoint_DefaultScript : ScriptBase
{
    // =====================================================
    // Configurable Variables
    // =====================================================
    // Global timings and speeds
    int globalTimeoutInSeconds = 60;              // How long to wait for actions (e.g., opening the app)
    int globalWaitInSeconds = 3;                  // Standard wait time between actions
    int waitMessageboxInSeconds = 2;              // Duration for onscreen wait messages
    int charactersPerMinuteToType = 15;           // Typing speed for keyboard shortcuts
    int waitInBetweenKeyboardShortcuts = 4;       // Wait time between keyboard shortcuts
    int slideshowCharactersPerMinuteToType = 12;  // Typing speed for slideshow navigation
    int pageScrollCpm = 60;                      // Typing speed for page scrolling actions
    int transitionPopupCharactersPerMinuteToType = 60; // Typing speed for navigating transitions popup
    int waitForPopupShowingInSeconds = 10;         // Wait time for popups to show (e.g., ribbon popups)
    int waitSlideshowStart = 10;                   // Wait time for slideshow to start
    int typingTextCharacterPerMinute = 600;         // Typing speed for saving and opening the file
    int startMenuWaitInSeconds = 5;                // Duration for Start Menu wait between interactions

    // File download settings
    string bmpUrl = "https://myAppliance.myOrg.com/contentDelivery/content/LoginVSI_BattlingRobots.bmp"; // URL for BMP file

    // =====================================================
    // Setup: Directories and File Downloads
    // =====================================================
    private void Execute()
    {   
        // Retrieve TEMP environment variable and ensure the LoginEnterprise directory exists.
        var temp = GetEnvironmentVariable("TEMP");
        string loginEnterpriseDir = $"{temp}\\LoginEnterprise";
        if (!Directory.Exists(loginEnterpriseDir))
        {
            Directory.CreateDirectory(loginEnterpriseDir);
            Log("Created directory: " + loginEnterpriseDir);
        }

        // =====================================================
        // Simulate Start Menu Interaction
        // =====================================================
        Log("Opening Start Menu");
        Wait(startMenuWaitInSeconds);
        Wait(seconds: waitMessageboxInSeconds, showOnScreen: true, onScreenText: "Start Menu");
        Type("{LWIN}", hideInLogging: false);
        Wait(startMenuWaitInSeconds);
        Type("{LWIN}", hideInLogging: false);
        Wait(seconds: 1);
        Type("{ESC}", hideInLogging: false);
        Wait(startMenuWaitInSeconds);

        // =====================================================
        // Download Files (Always overwrite PPTX)
        // =====================================================
        Wait(seconds: waitMessageboxInSeconds, showOnScreen: true, onScreenText: "Get .pptx and .bmp file");

        // -- Download PowerPoint (.pptx) file (overwrite without checking existence)
        string pptxFile = $"{loginEnterpriseDir}\\loginvsi.pptx";
        Log("Downloading PowerPoint presentation file.");
        CopyFile(KnownFiles.PowerPointPresentation, pptxFile, overwrite: false, continueOnError: true);

        // -- Download the BMP file if it doesn't exist.
        string bmpFile = $"{loginEnterpriseDir}\\LoginVSI_BattlingRobots.bmp";
        if (!FileExists(bmpFile))
        {
            Log("Downloading BMP file for slideshow");
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

        // =====================================================
        // Skip First-Run Dialogs before Bringing PowerPoint into Focus
        // =====================================================
        Wait(seconds: globalWaitInSeconds);
        SkipFirstRunDialogs();

        // =====================================================
        // Bring PowerPoint into Focus and Open PPTX File
        // =====================================================
        var MainWindow = FindWindow(className: "Win32 Window:PPTFrameClass", title: "*PowerPoint*", processName: "POWERPNT", timeout: globalTimeoutInSeconds);
        Wait(seconds: globalWaitInSeconds);
        MainWindow.Focus();
        MainWindow.Maximize();

        Log("Opening PPTX file via open file dialog");
        Wait(seconds: waitMessageboxInSeconds, showOnScreen: true, onScreenText: "Open pptx file");
        MainWindow.Type("{CTRL+O}", cpm: charactersPerMinuteToType, hideInLogging: false);
        Wait(seconds: waitInBetweenKeyboardShortcuts);
        MainWindow.Type("{ALT+O+O}", cpm: charactersPerMinuteToType, hideInLogging: false);
        StartTimer("Open_PPTX_Dialog");
        var openWindow = FindWindow(className: "Win32 Window:#32770", processName: "POWERPNT", continueOnError: false, timeout: globalTimeoutInSeconds);
        StopTimer("Open_PPTX_Dialog");
        Wait(seconds: globalWaitInSeconds);
        var fileNameBox = openWindow.FindControl(className: "Edit:Edit", title: "File name:", timeout: globalTimeoutInSeconds);
        fileNameBox.Click();
        Wait(seconds: globalWaitInSeconds);
        ScriptHelpers.SetTextBoxText(this, fileNameBox, pptxFile, cpm: typingTextCharacterPerMinute);
        Type("{enter}", hideInLogging: false);
        StartTimer("Open_Powerpoint_Document");
        var newPowerpoint = FindWindow(className: "Win32 Window:PPTFrameClass", title: "loginvsi*", processName: "POWERPNT", timeout: globalTimeoutInSeconds);
        newPowerpoint.Focus();
        newPowerpoint.Maximize();
        newPowerpoint.FindControl(className: "TabItem:NetUIRibbonTab", title: "Insert", timeout: globalTimeoutInSeconds);
        StopTimer("Open_Powerpoint_Document");
        Wait(seconds: globalWaitInSeconds);

        // =====================================================
        // Skip First-Run Dialogs before Checking for an Existing Edited Window
        // =====================================================
        Wait(seconds: globalWaitInSeconds);
        SkipFirstRunDialogs();

        // =====================================================
        // Check and Close Existing "Edited" Window (if any)
        // =====================================================
        string newDocName = "edited";
        for (int attempt = 0; attempt < 3; attempt++)
        {
            var editedWindow = FindWindow(className: "Win32 Window:PPTFrameClass", title: $"{newDocName}*", processName: "POWERPNT", timeout: 2, continueOnError: true);
            if (editedWindow != null)
            {
                Log("Existing edited window found. Closing it.");
                Wait(seconds: globalWaitInSeconds);
                editedWindow.Focus();
                editedWindow.Maximize();
                Wait(seconds: globalWaitInSeconds, showOnScreen: true, onScreenText: "Closing edited window");
                editedWindow.Type("{ALT+F4}", hideInLogging: false);
                Wait(seconds: globalWaitInSeconds);
                
                newPowerpoint = FindWindow(className: "Win32 Window:PPTFrameClass", title: "loginvsi*", processName: "POWERPNT", timeout: globalTimeoutInSeconds);
                newPowerpoint.Focus();
                newPowerpoint.Maximize();
                Wait(seconds: globalWaitInSeconds);
            }
        }

        // =====================================================
        // Add Transitions, Duplicate Slides, and Insert BMP
        // =====================================================
        Wait(seconds: waitMessageboxInSeconds, showOnScreen: true, onScreenText: "Adding bmp into a slide, duplicating it, and adding transitions");

        // --- Add new slide ---
        newPowerpoint.Focus();
        newPowerpoint.Maximize();
        Wait(seconds: globalWaitInSeconds);
        newPowerpoint.Type("{CTRL+M}", cpm: charactersPerMinuteToType, hideInLogging: false);
        Wait(seconds: waitInBetweenKeyboardShortcuts);

        // --- Insert BMP into New Slide ---
        Log("Inserting BMP into new slide");
        newPowerpoint.Type("{ALT}NP1", cpm: charactersPerMinuteToType, hideInLogging: false);
        Wait(seconds: waitForPopupShowingInSeconds);
        Type("d", cpm: charactersPerMinuteToType, hideInLogging: false);
        StartTimer("Insert_Picture_Dialog");
        var addPictureDialog = FindWindow(className: "Win32 Window:#32770", title: "Insert Picture", processName: "POWERPNT", timeout: globalTimeoutInSeconds);
        StopTimer("Insert_Picture_Dialog");
        addPictureDialog.Focus();
        var fileNameBox2 = addPictureDialog.FindControl(className: "Edit:Edit", title: "File name:", timeout: globalTimeoutInSeconds);
        Wait(seconds: globalWaitInSeconds);
        fileNameBox2.Click();
        newPowerpoint.Type("{ALT+N}", cpm: charactersPerMinuteToType, hideInLogging: false);
        Wait(seconds: waitInBetweenKeyboardShortcuts);
        ScriptHelpers.SetTextBoxText(this, fileNameBox2, bmpFile, cpm: typingTextCharacterPerMinute);
        fileNameBox2.Type("{ENTER}", cpm: charactersPerMinuteToType, hideInLogging: false);
        Wait(seconds: globalWaitInSeconds);
        var stillExists = FindWindow(className: "Win32 Window:#32770", title: "Insert Picture", processName: "POWERPNT", timeout: 2, continueOnError: true);
        if (stillExists != null)
        {
            newPowerpoint.Type("{ESC}", cpm: charactersPerMinuteToType, hideInLogging: false);
        }

        // --- Add 'Honeycomb' Transition ---
        Log("Adding 'Honeycomb' transition");
        newPowerpoint.Type("{ALT+K}{ALT+T}", cpm: charactersPerMinuteToType, hideInLogging: false);        
        newPowerpoint.FindControl(title: "Honeycomb", timeout: globalTimeoutInSeconds);
        Wait(seconds: globalWaitInSeconds);
        newPowerpoint.Type("{DOWN}", cpm: transitionPopupCharactersPerMinuteToType, hideInLogging: false);
        newPowerpoint.Type("{RIGHT}".Repeat(16), cpm: transitionPopupCharactersPerMinuteToType, hideInLogging: false);
        Wait(seconds: waitInBetweenKeyboardShortcuts);
        newPowerpoint.Type("{ENTER}", hideInLogging: false);
        Wait(seconds: globalWaitInSeconds);

        // --- Add the transition to all slides ---
        Log("Adding transition to all slides");
        newPowerpoint.Type("{ALT+K}{ALT+L}", hideInLogging: false);
        Wait(seconds: globalWaitInSeconds);
        
        /* // --- Add 'Curtains' Transition and Duplicate Slide ---
        Log("Adding 'Curtains' transition");
        newPowerpoint.Type("{ALT+K}{ALT+T}", cpm: charactersPerMinuteToType, hideInLogging: false);
        var transitionCurtains = newPowerpoint.FindControl(title: "Curtains", timeout: globalTimeoutInSeconds);
        Wait(seconds: globalWaitInSeconds);
        transitionCurtains.Click();
        Wait(seconds: globalWaitInSeconds);
        newPowerpoint.Type("{ALT}NSI", cpm: charactersPerMinuteToType, hideInLogging: false);
        Wait(seconds: waitForPopupShowingInSeconds);
        Type("d", cpm: charactersPerMinuteToType, hideInLogging: false);
        Wait(seconds: globalWaitInSeconds);

        // --- Add 'Origami' Transition and Duplicate Slide ---
        Log("Adding 'Origami' transition");
        newPowerpoint.Type("{ALT+K}{ALT+T}", cpm: charactersPerMinuteToType, hideInLogging: false);        
        newPowerpoint.FindControl(title: "Origami", timeout: globalTimeoutInSeconds);
        Wait(seconds: globalWaitInSeconds);
        newPowerpoint.Type("{DOWN}", cpm: transitionPopupCharactersPerMinuteToType, hideInLogging: false);
        newPowerpoint.Type("{RIGHT}".Repeat(10), cpm: transitionPopupCharactersPerMinuteToType, hideInLogging: false);
        newPowerpoint.Type("{ENTER}", hideInLogging: false);
        Wait(seconds: globalWaitInSeconds);
        newPowerpoint.Type("{ALT}NSI", cpm: charactersPerMinuteToType, hideInLogging: false);
        Wait(seconds: waitForPopupShowingInSeconds);
        Type("d", cpm: charactersPerMinuteToType, hideInLogging: false);
        Wait(seconds: globalWaitInSeconds);

        // --- Add 'Curtains' Transition and Duplicate Slide ---
        Log("Adding 'Curtains' transition");
        newPowerpoint.Type("{ALT+K}{ALT+T}", cpm: charactersPerMinuteToType, hideInLogging: false);        
        newPowerpoint.FindControl(title: "Curtains", timeout: globalTimeoutInSeconds);
        Wait(seconds: globalWaitInSeconds);
        newPowerpoint.Type("{DOWN}", cpm: transitionPopupCharactersPerMinuteToType, hideInLogging: false);
        newPowerpoint.Type("{RIGHT}".Repeat(2), cpm: transitionPopupCharactersPerMinuteToType, hideInLogging: false);
        newPowerpoint.Type("{ENTER}", hideInLogging: false);
        Wait(seconds: globalWaitInSeconds);
        newPowerpoint.Type("{ALT}NSI", cpm: charactersPerMinuteToType, hideInLogging: false);
        Wait(seconds: waitForPopupShowingInSeconds);
        Type("d", cpm: charactersPerMinuteToType, hideInLogging: false);
        Wait(seconds: globalWaitInSeconds);

        // --- Add 'Ripple' Transition and Duplicate Slide ---
        Log("Adding 'Ripple' transition");
        newPowerpoint.Type("{ALT+K}{ALT+T}", cpm: charactersPerMinuteToType, hideInLogging: false);        
        newPowerpoint.FindControl(title: "Ripple", timeout: globalTimeoutInSeconds);
        Wait(seconds: globalWaitInSeconds);
        newPowerpoint.Type("{DOWN}", cpm: transitionPopupCharactersPerMinuteToType, hideInLogging: false);
        newPowerpoint.Type("{RIGHT}".Repeat(15), cpm: transitionPopupCharactersPerMinuteToType, hideInLogging: false);
        newPowerpoint.Type("{ENTER}", hideInLogging: false);
        Wait(seconds: globalWaitInSeconds);
        newPowerpoint.Type("{ALT}NSI", cpm: charactersPerMinuteToType, hideInLogging: false);
        Wait(seconds: waitForPopupShowingInSeconds);
        Type("d", cpm: charactersPerMinuteToType, hideInLogging: false);
        Wait(seconds: globalWaitInSeconds);

        // --- Add 'Vortex' Transition and Duplicate Slide ---
        Log("Adding 'Vortex' transition");
        newPowerpoint.Type("{ALT+K}{ALT+T}", cpm: charactersPerMinuteToType, hideInLogging: false);        
        newPowerpoint.FindControl(title: "Vortex", timeout: globalTimeoutInSeconds);
        Wait(seconds: globalWaitInSeconds);
        newPowerpoint.Type("{DOWN}".Repeat(2), cpm: transitionPopupCharactersPerMinuteToType, hideInLogging: false);
        newPowerpoint.Type("{RIGHT}", cpm: transitionPopupCharactersPerMinuteToType, hideInLogging: false);
        newPowerpoint.Type("{ENTER}", hideInLogging: false);
        Wait(seconds: globalWaitInSeconds);
        */
        
        // =====================================================
        // Scroll, Minimize, and Maximize Presentation
        // =====================================================
        Wait(seconds: waitMessageboxInSeconds, showOnScreen: true, onScreenText: "Scroll, minimize, and maximize");
        Log("Scrolling through presentation");
        newPowerpoint.Type("{PAGEDOWN}".Repeat(10), cpm: pageScrollCpm, hideInLogging: false);
        Wait(seconds: globalWaitInSeconds);
        newPowerpoint.Type("{PAGEUP}".Repeat(10), cpm: pageScrollCpm, hideInLogging: false);
        Wait(seconds: globalWaitInSeconds);  

        newPowerpoint.Minimize();
        Wait(seconds: globalWaitInSeconds);
        newPowerpoint.Maximize();
        newPowerpoint.Focus();
        Wait(seconds: globalWaitInSeconds);
        
        // =====================================================
        // Run the Slideshow
        // =====================================================
        Wait(seconds: waitMessageboxInSeconds, showOnScreen: true, onScreenText: "Run the slideshow");
        Log("Starting slideshow");
        newPowerpoint.Type("{F5}", cpm: charactersPerMinuteToType, hideInLogging: false);
        Wait(seconds: waitSlideshowStart);
        newPowerpoint.Type("{DOWN}".Repeat(5), cpm: slideshowCharactersPerMinuteToType, hideInLogging: false);
        Wait(seconds: waitInBetweenKeyboardShortcuts);
        Type("{ESC}", hideInLogging: false);
        Wait(seconds: globalWaitInSeconds);
        Type("{HOME}", hideInLogging: false); // Go to the first slide
        Wait(seconds: globalWaitInSeconds);

        // =====================================================
        // Save the Edited Slideshow
        // =====================================================
        Log("Saving the edited slideshow");        
        Wait(seconds: waitMessageboxInSeconds, showOnScreen: true, onScreenText: "Save the slideshow");
        string saveFilename = $"{loginEnterpriseDir}\\{newDocName}.pptx";
        if (FileExists(saveFilename))
        {
            Log("Removing existing file: " + saveFilename);
            RemoveFile(saveFilename);
        }
        else
        {
            Log("No existing file to remove at: " + saveFilename);
        }
        newPowerpoint.Type("{F12}", hideInLogging: false);
        StartTimer("Save_As_Dialog");
        var saveAs = FindWindow(className: "Win32 Window:#32770", processName: "POWERPNT", timeout: globalTimeoutInSeconds);
        StopTimer("Save_As_Dialog");
        var saveFileNameBox = saveAs.FindControl(className: "Edit:Edit", title: "File name:", timeout: globalTimeoutInSeconds);
        saveFileNameBox.Click();
        Wait(seconds: globalWaitInSeconds);
        ScriptHelpers.SetTextBoxText(this, saveFileNameBox, saveFilename, cpm: typingTextCharacterPerMinute);
        saveAs.Type("{ENTER}", hideInLogging: false);
        StartTimer("Saving_file");        
        FindWindow(title: $"{newDocName}*", processName: "POWERPNT", timeout: globalTimeoutInSeconds);
        StopTimer("Saving_file");
        Wait(seconds: globalWaitInSeconds);

        Log("Script complete. PowerPoint remains open.");
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
                processName: "POWERPNT", 
                continueOnError: true, 
                timeout: 5);
            while (dialog != null)
            {
                dialog.Close();
                dialog = FindWindow(
                    className: "Win32 Window:NUIDialog", 
                    processName: "POWERPNT", 
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
        int globalWaitInSeconds = 3;                  // Standard wait time between actions
        var numTries = 1;
        string currentText = null;
        do
        {
            textBox.Type("{CTRL+a}", hideInLogging: false);
            script.Wait(globalWaitInSeconds);
            textBox.Type(text, cpm: cpm, hideInLogging: false);
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
