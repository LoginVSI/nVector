// TARGET:outlook.exe /importprf %TEMP%\LoginEnterprise\Outlook.prf
// START_IN:

/////////////
// Outlook Running Script
// Workload: KnowledgeWorker 2025
// Version: 1.0
/////////////

using LoginPI.Engine.ScriptBase;
using System.IO;
using System;
using System.Net;
using System.Net.Security;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;

public class Outlook_DefaultScript : ScriptBase
{
    // Import the user32.dll function to simulate mouse events.
    [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
    public static extern void mouse_event(uint dwFlags, uint dx, uint dy, int dwData, UIntPtr dwExtraInfo);
    public const uint MOUSEEVENTF_WHEEL = 0x0800; // Constant for a mouse wheel event

    private void Execute()
    {
        int globalTimeoutInSeconds = 60; // Timeout for functions
        int globalWaitInSeconds = 3; // Wait in between functions
        int waitMessageboxInSeconds = 3; // Waits with information message        
        int keyboardShortcutsCPM = 50; // typing speed for keyboard shortcuts
        int typingTextCharacterPerMinute = 600; // typing speed for email body text
        int copyPasteRepetitions = 2; // Copy and pasting iterations in the email body
        string bmpUrl = "https://myAppliance.myOrg.com/contentDelivery/content/LoginVSI_BattlingRobots.bmp"; // Replace with the actual URL for the BMP file, such as "https://myAppliance.myOrg.com/contentDelivery/content/LoginVSI_BattlingRobots.bmp"

        // Download image to add to an email draft
        var temp = GetEnvironmentVariable("TEMP");
        string outlookDir = Path.Combine(temp, "LoginEnterprise");
        if (!Directory.Exists(outlookDir))
        {
            Directory.CreateDirectory(outlookDir);
            Log("Created directory: " + outlookDir);
        }

        string bmpFile = Path.Combine(outlookDir, "LoginVSI_BattlingRobots.bmp");
        if (!FileExists(bmpFile))
        {
            Log("Downloading BMP file");
            try
            {
                // Disable SSL certificate validation for the download
                ServicePointManager.ServerCertificateValidationCallback =
                    delegate { return true; };

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
        
        // Open/Close Start Menu
        Log("Opening Start Menu");
        Wait(seconds: waitMessageboxInSeconds, showOnScreen: true, onScreenText: "Start Menu");
        Type("{LWIN}");
        Wait(3);
        Type("{LWIN}");
        Wait(1);
        Type("{ESC}");

        // Verifying Outlook is already running
        var mainWindow = FindWindow(title:"Inbox -*", processName: "OUTLOOK", timeout: globalTimeoutInSeconds);
        mainWindow.Focus();
        mainWindow.Maximize();
        Wait(globalWaitInSeconds);
        
        // Dismiss Activate/first run Office popup, if it appears
        var signinWindow = mainWindow.FindControlWithXPath("Win32 Window:NUIDialog", timeout: 2,continueOnError:true);
        signinWindow?.Type("{ESC}");
        
        // Dismiss Reminder popup, if exists
        var reminderWindow = FindWindow(className: "Win32 Window:#32770", title: "*Reminder(s)",processName: "OUTLOOK", timeout: 2, continueOnError: true);
        if (reminderWindow != null)
        {
            Wait(globalWaitInSeconds);                
            // Try clicking "Dismiss &All"
            var dismissAllBtn = reminderWindow.FindControl(className: "Button:Button", title: "Dismiss &All", continueOnError: true,timeout:2);
            if (dismissAllBtn != null)
            {
                reminderWindow.Focus();
                Wait(globalWaitInSeconds);    
                dismissAllBtn.Click();
            }                
            // Try "Yes" to confirm
            var yesBtn = reminderWindow.FindControl(className: "Button:Button", title: "&Yes", continueOnError: true,timeout:2);
            if (yesBtn != null)
            {
                reminderWindow.Focus();
                Wait(globalWaitInSeconds);
                yesBtn.Click();
                mainWindow.Focus();
            }
        }
        
        // Open and scroll on an email
        var inboxWindow = mainWindow.FindControlWithXPath("Table:SuperGrid",timeout:globalTimeoutInSeconds);
        Log("Reading an existing email");
        inboxWindow.Focus();
        Wait(seconds: globalWaitInSeconds, showOnScreen: true, onScreenText: "Open and Read an Email");            
        inboxWindow.Click();
        Wait(globalWaitInSeconds);
        inboxWindow.Type("{DOWN}".Repeat(5), cpm: keyboardShortcutsCPM);
        inboxWindow.Type("{UP}".Repeat(5), cpm: keyboardShortcutsCPM);
        inboxWindow.Type("{ENTER}", cpm: keyboardShortcutsCPM);
        StartTimer("Open_Existing_Email");
        var openEmail = FindWindow(className : "Win32 Window:rctrl_renwnd32", title : "Login Enterprise *", processName : "OUTLOOK",timeout: globalTimeoutInSeconds);
        StopTimer("Open_Existing_Email");
        Wait(globalWaitInSeconds);
        openEmail.Focus();
        openEmail.Maximize();
        Wait(globalWaitInSeconds);
        openEmail.MoveMouseToCenter();
        openEmail.Click();
        
        {
            void InlineScroll(string direction, int scrollCount, int notches, double waitTime)
            {
                if (waitTime <= 0)
                    throw new ArgumentException("Scroll waitTime must be > 0 seconds.");

                int sign = direction.Equals("Down", StringComparison.OrdinalIgnoreCase) ? -1 : 1;
                int delta = sign * 120 * notches;

                Log($"Scrolling mouse {direction} {scrollCount} times, {notches} notch(es) each, {waitTime}s between.");
                for (int i = 0; i < scrollCount; i++)
                {
                    mouse_event(MOUSEEVENTF_WHEEL, 0, 0, delta, UIntPtr.Zero);
                    Wait(waitTime);
                }
            }
            InlineScroll("Down", 10, 1, 0.1);
            InlineScroll("Up", 10, 1, 0.1);
        }        
        openEmail.Close();

        // Minimize and maximize Outlook
        mainWindow.Minimize();
        Wait(globalWaitInSeconds);
        mainWindow.Maximize();
        mainWindow.Focus();
        Wait(globalWaitInSeconds);

        // Compose a new email
        Log("Initiating new email composition");
        mainWindow.Type("{CTRL+N}", cpm: keyboardShortcutsCPM);
        StartTimer("New_Email_Open");
        var newEmail = FindWindow(className: "Win32 Window:rctrl_renwnd32", title: "Untitled *", processName: "OUTLOOK",timeout: globalTimeoutInSeconds);
        StopTimer("New_Email_Open");
        Wait(globalWaitInSeconds);
        newEmail.Focus();
        newEmail.Maximize();
        Wait(globalWaitInSeconds);

        // Add text to the email
        Log("Populating email fields");
        var toField = newEmail.FindControl(className: "*RichEdit20WPT", title: "To",timeout:globalTimeoutInSeconds);
        Wait(globalWaitInSeconds);
        toField.Type("LoginEnterpriseVirtualUser1@LoginVSI.com;LoginEnterpriseVirtualUser2@LoginVSI.com;", cpm: typingTextCharacterPerMinute);
        newEmail.Type("{ALT+U}", cpm: 50);
        Wait(globalWaitInSeconds);
        newEmail.Type("About Login VSI", cpm: typingTextCharacterPerMinute);
        newEmail.Type("{TAB}", cpm: 50);
        Wait(globalWaitInSeconds);
        string emailBody = @"About Login VSI{Enter}The VDI and DaaS industry has transformed incredibly, and Login VSI has evolved alongside the world of remote and hybrid work.{Enter}Through an innovative and dynamic culture, the Login VSI team is passionate about helping enterprises worldwide understand, build, and maintain amazing digital workspaces.{Enter}Trusted globally for 360° proactive visibility of performance, cost, and capacity of virtual desktops and applications, Login Enterprise is accepted as the industry standard and used by major vendors to spot problems quicker, avoid unexpected downtime, and deliver next-level digital experiences for end-users.{Enter}Our Mission{Enter}The paradigm for remote computing has shifted with virtual app delivery coupled with the growth in Web and SaaS-based applications.{Enter}Now more than ever, organizations rely on digital workspaces to function. We give our customers 360° insights into the entire stack of virtual desktops and applications – in production or delivery and across various settings and infrastructure.{Enter}We aim to empower IT teams to take control of their virtual desktops and applications’ performance, cost, and capacity wherever they reside – traditional, hybrid, or cloud.";
        newEmail.Type(emailBody, cpm: typingTextCharacterPerMinute);
        Wait(globalWaitInSeconds);
        
        // Add picture file to the email body
        Log("Inserting BMP image into email body");
        newEmail.Type("{ALT}np", cpm: keyboardShortcutsCPM); // e.g., Insert -> Picture
        StartTimer("Insert_Picture_Dialog");
        var insertPictureDialog = FindWindow(className : "Win32 Window:#32770", title : "Insert Picture", processName : "OUTLOOK",timeout:globalTimeoutInSeconds);
        StopTimer("Insert_Picture_Dialog");
        var insertPictureDialogFileName = insertPictureDialog.FindControl(className : "Edit:Edit", title : "File name:",timeout:globalTimeoutInSeconds);
        Wait(globalWaitInSeconds);
        insertPictureDialogFileName.Click();
        Wait(globalWaitInSeconds);
        Type(bmpFile + "{ENTER}", cpm: 1000);
        Wait(globalWaitInSeconds);

        // Copy & Paste the email body to expand its content
        Log("Copying and pasting email body content");
        newEmail.Type("{CTRL+A}", cpm: keyboardShortcutsCPM);
        Wait(globalWaitInSeconds);
        newEmail.Type("{CTRL+C}", cpm: keyboardShortcutsCPM);
        Wait(globalWaitInSeconds);
        for (int i = 0; i < copyPasteRepetitions; i++)
        {
            newEmail.Type("{CTRL+V}", cpm: keyboardShortcutsCPM);
            Wait(globalWaitInSeconds);
        }
        
        // Scroll in the new email window
        Log("Scrolling within the new email window");
        newEmail.Focus();
        newEmail.MoveMouseToCenter();
        Wait(globalWaitInSeconds);
        newEmail.Click();
        Wait(globalWaitInSeconds);
        {
            void InlineScroll(string direction, int scrollCount, int notches, double waitTime)
            {
                if (waitTime <= 0)
                    throw new ArgumentException("Scroll waitTime must be > 0 seconds.");

                int sign = direction.Equals("Down", StringComparison.OrdinalIgnoreCase) ? -1 : 1;
                int delta = sign * 120 * notches;

                Log($"Scrolling mouse {direction} {scrollCount} times, {notches} notch(es) each, {waitTime}s between.");
                for (int i = 0; i < scrollCount; i++)
                {
                    mouse_event(MOUSEEVENTF_WHEEL, 0, 0, delta, UIntPtr.Zero);
                    Wait(waitTime);
                }
            }
            InlineScroll("Up", 40, 1, 0.1);
            InlineScroll("Down", 40, 1, 0.1);
        }
        Wait(globalWaitInSeconds);

        // Add picture file as email attachment
        Log("Attaching BMP file to email");
        newEmail.Type("{ALT}naf", cpm: keyboardShortcutsCPM); // Insert and Attach
        Wait(globalWaitInSeconds);
        newEmail.Type("b", cpm: keyboardShortcutsCPM);      // 'b' for browse
        StartTimer("Add_Attachment_Dialog");
        var addAttachmentDialog = FindWindow(className : "Win32 Window:#32770", title : "Insert File", processName : "OUTLOOK",timeout:globalTimeoutInSeconds);
        StopTimer("Add_Attachment_Dialog");
        Wait(globalWaitInSeconds);
        addAttachmentDialog.Focus();
        StopTimer("Add_Attachment_Dialog");
        Type("{alt+n}");
        Wait(globalWaitInSeconds);
        Type(bmpFile + "{ENTER}", cpm: 1000);
        Wait(globalWaitInSeconds);
        
        // Close the new email
        newEmail.Close();
        Wait(globalWaitInSeconds);
        Type("n");
        
        Wait(globalWaitInSeconds);
        Log("Outlook prepared for next iteration. Main window persists; all other windows closed.");
    }
}
