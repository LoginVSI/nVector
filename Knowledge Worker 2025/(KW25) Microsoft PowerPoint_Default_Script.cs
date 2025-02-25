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
    private void Execute()
    {   
        // Global settings
        int globalTimeoutInSeconds = 60;              // How long to wait for actions (e.g., opening the app)
        int waitMessageboxInSeconds = 3;              // Duration for wait message boxes
        double globalWaitInSeconds = 3;               // Wait time between actions for human-like behavior (double value)
        int charactersPerMinuteToType = 50;           // Typing speed for keyboard shortcuts
        int slideshowCharactersPerMinuteToType = 50;  // How many down arrows to press per minute to go through the slideshow presentation (60 = 1 per second)
        int waitSlideshowStart = 4;                   // Allow time for slideshow to start
        int pageScrollCpm = 120;                      // Typing speed for page scrolling actions (60 = 1 per second)
        int transitionPopupCharactersPerMinuteToType = 200; // How quickly to navigate around the insert transitions popup (60 = 1 per second)
        int waitForPopupShowingInSeconds = 3;         // Wait time for popups to show, such as ribbon popups
        var temp = GetEnvironmentVariable("TEMP");

        // Define the BMP URL as a variable.
        string bmpUrl = "<your URL here>"; // Replace with the actual URL for the BMP file, such as "https://myAppliance.myOrg.com/contentDelivery/content/LoginVSI_BattlingRobots.bmp";

        // Ensure the LoginEnterprise directory exists.
        string loginEnterpriseDir = $"{temp}\\LoginEnterprise";
        if (!Directory.Exists(loginEnterpriseDir))
        {
            Directory.CreateDirectory(loginEnterpriseDir);
            Log("Created directory: " + loginEnterpriseDir);
        }
        
        // Simulate user interaction to open the Start Menu.
        Log("Opening Start Menu");
        Wait(seconds:2, showOnScreen:true, onScreenText:"Start Menu");
        Type("{LWIN}");
        Wait(3);
        Type("{LWIN}");
        Wait(1);
        Type("{ESC}");

        // Download the PowerPoint (.pptx) file and bmp if they dont already exist.
        Wait(seconds:waitMessageboxInSeconds, showOnScreen:true, onScreenText:"Get .pptx and .bmp file");
        string pptxFile = $"{loginEnterpriseDir}\\loginvsi.pptx";
        if (!FileExists(pptxFile))
        {
            Log("Downloading PowerPoint presentation file");
            CopyFile(KnownFiles.PowerPointPresentation, pptxFile);
        }
        else
        {
            Log("PowerPoint file already exists");
        }

        // Download the BMP file if it doesn't exist.
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
                    // Use the defined bmpUrl variable.
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

        // Use the already open PowerPoint instance.
        var MainWindow = FindWindow(className:"Win32 Window:PPTFrameClass", title:"*PowerPoint*", processName:"POWERPNT", timeout:globalTimeoutInSeconds);
        Wait(globalWaitInSeconds);
        MainWindow.Focus();

        // --- Open File Dialog to Open PPTX ---
        Log("Opening PPTX file via open file dialog");
        Wait(seconds:waitMessageboxInSeconds, showOnScreen:true, onScreenText:"Open pptx file");
        MainWindow.Type("{CTRL+O}{ALT+O+O}", cpm: charactersPerMinuteToType);
        StartTimer("Open_PPTX_Dialog");
        var openWindow = FindWindow(className:"Win32 Window:#32770", processName:"POWERPNT", continueOnError:false, timeout:globalTimeoutInSeconds);
        StopTimer("Open_PPTX_Dialog");
        
        Wait(globalWaitInSeconds);
        var fileNameBox = openWindow.FindControl(className:"Edit:Edit", title:"File name:", timeout:globalTimeoutInSeconds);
        fileNameBox.Click();
        Wait(globalWaitInSeconds);
        ScriptHelpers.SetTextBoxText(this, fileNameBox, pptxFile, cpm:1000);
        Type("{enter}");
        StartTimer("Open_Powerpoint_Document");
        var newPowerpoint = FindWindow(className:"Win32 Window:PPTFrameClass", title:"loginvsi*", processName:"POWERPNT", timeout:globalTimeoutInSeconds);
        newPowerpoint.Focus();
        newPowerpoint.FindControl(className:"TabItem:NetUIRibbonTab", title:"Insert", timeout:globalTimeoutInSeconds);
        StopTimer("Open_Powerpoint_Document");
        Wait(globalWaitInSeconds);
        
        // --- Check for an existing "edited" PowerPoint window ---
        string newDocName = "edited";
        var editedWindow = FindWindow(className:"Win32 Window:PPTFrameClass", title:$"{newDocName}*", processName:"POWERPNT", timeout:2, continueOnError:true);
        if (editedWindow != null)
        {
            Log("Existing edited window found. Closing it.");
            editedWindow.Close();
            Wait(globalWaitInSeconds);
            newPowerpoint = FindWindow(className:"Win32 Window:PPTFrameClass", title:"loginvsi*", processName:"POWERPNT", timeout:globalTimeoutInSeconds);
            newPowerpoint.Focus();
            Wait(globalWaitInSeconds);
        }
        
        Wait(seconds:waitMessageboxInSeconds, showOnScreen:true, onScreenText:"Adding bmp into a slide, duplicating it, and adding transitions");

        // --- Add 'Curtains' Transition and Duplicate Slide ---
        Log("Adding 'Curtains' transition");
        newPowerpoint.Type("{ALT+K}{ALT+T}", cpm: charactersPerMinuteToType);
        var transitionCurtains = newPowerpoint.FindControl(title:"Curtains", timeout:globalTimeoutInSeconds);
        Wait(globalWaitInSeconds);
        transitionCurtains.Click();
        Wait(globalWaitInSeconds);
        newPowerpoint.Type("{CTRL+M}", cpm: charactersPerMinuteToType);
        Wait(globalWaitInSeconds);

        // --- Insert BMP into New Slide ---
        Log("Inserting BMP into new slide");
        newPowerpoint.Type("{ALT}NP1", cpm: charactersPerMinuteToType);
        Wait(waitForPopupShowingInSeconds);  // Wait before typing the "d" for the BMP add popup
        Type("d", cpm: charactersPerMinuteToType);
        StartTimer("Insert_Picture_Dialog");
        var addPictureDialog = FindWindow(className:"Win32 Window:#32770", title:"Insert Picture", processName:"POWERPNT", timeout:globalTimeoutInSeconds);
        StopTimer("Insert_Picture_Dialog");
        addPictureDialog.Focus();
        var fileNameBox2 = addPictureDialog.FindControl(className:"Edit:Edit", title:"File name:", timeout:globalTimeoutInSeconds);
        Wait(globalWaitInSeconds);
        fileNameBox2.Click();
        newPowerpoint.Type("{ALT+N}", cpm: charactersPerMinuteToType);
        Wait(globalWaitInSeconds);
        ScriptHelpers.SetTextBoxText(this, fileNameBox2, bmpFile, cpm:1000);
        fileNameBox2.Type("{ENTER}", cpm: charactersPerMinuteToType);
        Wait(globalWaitInSeconds);
        var stillExists = FindWindow(className:"Win32 Window:#32770", title:"Insert Picture", processName:"POWERPNT", timeout:2, continueOnError:true);
        if(stillExists != null)
        {
            newPowerpoint.Type("{ESC}", cpm: charactersPerMinuteToType);
        }

        // --- Add 'Origami' Transition and Duplicate Slide ---
        Log("Adding 'Origami' transition");
        newPowerpoint.Type("{ALT+K}{ALT+T}", cpm: charactersPerMinuteToType);        
        newPowerpoint.FindControl(title:"Origami", timeout:globalTimeoutInSeconds);
        Wait(globalWaitInSeconds);
        newPowerpoint.Type("{DOWN}", cpm: transitionPopupCharactersPerMinuteToType);
        newPowerpoint.Type("{RIGHT}".Repeat(10), cpm: transitionPopupCharactersPerMinuteToType);
        newPowerpoint.Type("{ENTER}");
        Wait(globalWaitInSeconds);
        newPowerpoint.Type("{ALT}NSI", cpm: charactersPerMinuteToType);
        Wait(waitForPopupShowingInSeconds);  // Wait for the duplicate slide popup to show
        Type("d", cpm: charactersPerMinuteToType);
        Wait(globalWaitInSeconds);

        // --- Add 'Curtains' Transition and Duplicate Slide ---
        Log("Adding 'Curtains' transition");
        newPowerpoint.Type("{ALT+K}{ALT+T}", cpm: charactersPerMinuteToType);        
        newPowerpoint.FindControl(title:"Curtains", timeout:globalTimeoutInSeconds);
        Wait(globalWaitInSeconds);
        newPowerpoint.Type("{DOWN}", cpm: transitionPopupCharactersPerMinuteToType);
        newPowerpoint.Type("{RIGHT}".Repeat(2), cpm: transitionPopupCharactersPerMinuteToType);
        newPowerpoint.Type("{ENTER}");
        Wait(globalWaitInSeconds);
        newPowerpoint.Type("{ALT}NSI", cpm: charactersPerMinuteToType);
        Wait(waitForPopupShowingInSeconds);  // Wait for the duplicate slide popup to show
        Type("d", cpm: charactersPerMinuteToType);
        Wait(globalWaitInSeconds);
        
        // --- Add 'Ripple' Transition and Duplicate Slide ---
        Log("Adding 'Ripple' transition");
        newPowerpoint.Type("{ALT+K}{ALT+T}", cpm: charactersPerMinuteToType);        
        newPowerpoint.FindControl(title:"Ripple", timeout:globalTimeoutInSeconds);
        Wait(globalWaitInSeconds);
        newPowerpoint.Type("{DOWN}", cpm: transitionPopupCharactersPerMinuteToType);
        newPowerpoint.Type("{RIGHT}".Repeat(15), cpm: transitionPopupCharactersPerMinuteToType);
        newPowerpoint.Type("{ENTER}");
        Wait(globalWaitInSeconds);
        newPowerpoint.Type("{ALT}NSI", cpm: charactersPerMinuteToType);
        Wait(waitForPopupShowingInSeconds);  // Wait for the duplicate slide popup to show
        Type("d", cpm: charactersPerMinuteToType);
        Wait(globalWaitInSeconds);
        
        // --- Add 'Honeycomb' Transition and Duplicate Slide ---
        Log("Adding 'Honeycomb' transition");
        newPowerpoint.Type("{ALT+K}{ALT+T}", cpm: charactersPerMinuteToType);        
        newPowerpoint.FindControl(title:"Honeycomb", timeout:globalTimeoutInSeconds);
        Wait(globalWaitInSeconds);
        newPowerpoint.Type("{DOWN}", cpm: transitionPopupCharactersPerMinuteToType);
        newPowerpoint.Type("{RIGHT}".Repeat(16), cpm: transitionPopupCharactersPerMinuteToType);
        newPowerpoint.Type("{ENTER}");
        Wait(globalWaitInSeconds);
        newPowerpoint.Type("{ALT}NSI", cpm: charactersPerMinuteToType);
        Wait(waitForPopupShowingInSeconds);  // Wait for the duplicate slide popup to show
        Type("d", cpm: charactersPerMinuteToType);
        Wait(globalWaitInSeconds);
        
        // --- Add 'Vortex' Transition and Duplicate Slide ---
        Log("Adding 'Vortex' transition");
        newPowerpoint.Type("{ALT+K}{ALT+T}", cpm: charactersPerMinuteToType);        
        newPowerpoint.FindControl(title:"Vortex", timeout:globalTimeoutInSeconds);
        Wait(globalWaitInSeconds);
        newPowerpoint.Type("{DOWN}".Repeat(2), cpm: transitionPopupCharactersPerMinuteToType);
        newPowerpoint.Type("{RIGHT}", cpm: transitionPopupCharactersPerMinuteToType);
        newPowerpoint.Type("{ENTER}");
        Wait(globalWaitInSeconds);
        
        Wait(seconds:waitMessageboxInSeconds, showOnScreen:true, onScreenText:"Scroll, minimize, and maximize");
        
        // --- Scroll Through the Presentation ---
        Log("Scrolling through presentation");
        newPowerpoint.Type("{PAGEDOWN}".Repeat(20), cpm: pageScrollCpm);
        Wait(globalWaitInSeconds);
        newPowerpoint.Type("{PAGEUP}".Repeat(20), cpm: pageScrollCpm);
        Wait(globalWaitInSeconds);  

        // --- Minimize and Maximize the Presentation Window ---
        newPowerpoint.Minimize();
        Wait(globalWaitInSeconds);
        newPowerpoint.Maximize();
        newPowerpoint.Focus();
        Wait(globalWaitInSeconds);
        
        Wait(seconds:waitMessageboxInSeconds, showOnScreen:true, onScreenText:"Run the slideshow");
        
        // --- Run the Slideshow ---
        Log("Starting slideshow");
        newPowerpoint.Type("{F5}", cpm: charactersPerMinuteToType);
        Wait(waitSlideshowStart);
        newPowerpoint.Type("{DOWN}".Repeat(20), cpm: slideshowCharactersPerMinuteToType);
        Wait(globalWaitInSeconds);
        Type("{ESC}");
        Wait(globalWaitInSeconds);
        Type("{HOME}"); // Go to the first slide
        Wait(globalWaitInSeconds);

        // --- Save the Edited Slideshow ---
        Log("Saving the edited slideshow");        
        Wait(seconds:waitMessageboxInSeconds, showOnScreen:true, onScreenText:"Save the slideshow");
        // newDocName already defined above.
        string saveFilename = $"{loginEnterpriseDir}\\{newDocName}.pptx";
        if(FileExists(saveFilename))
        {
            Log("Removing existing file: " + saveFilename);
            RemoveFile(saveFilename);
        }
        else
        {
            Log("No existing file to remove at: " + saveFilename);
        }
        newPowerpoint.Type("{F12}", cpm: charactersPerMinuteToType);
        StartTimer("Save_As_Dialog");
        var saveAs = FindWindow(className:"Win32 Window:#32770", processName:"POWERPNT", continueOnError:true, timeout:globalTimeoutInSeconds);
        if (saveAs == null)
        {
            ABORT("Save file dialog could not be found");
        }
        var saveFileNameBox = saveAs.FindControl(className:"Edit:Edit", title:"File name:", timeout:globalTimeoutInSeconds);
        saveFileNameBox.Click();
        Wait(globalWaitInSeconds);
        ScriptHelpers.SetTextBoxText(this, saveFileNameBox, saveFilename, cpm:1000);
        saveAs.Type("{ENTER}", cpm: charactersPerMinuteToType);
        StartTimer("Saving_file");        
        FindWindow(title:$"{newDocName}*", processName:"POWERPNT", timeout:globalTimeoutInSeconds);
        StopTimer("Saving_file");
        Wait(globalWaitInSeconds);

        Log("Script complete. PowerPoint remains open.");        
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
