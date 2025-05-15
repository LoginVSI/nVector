// TARGET:outlook.exe /importprf %TEMP%\LoginEnterprise\Outlook.prf
// START_IN:

/////////////
// Outlook Run
// Workload: Knowledge Worker 2025
// Version: 0.1.0
/////////////

using LoginPI.Engine.ScriptBase;
using LoginPI.Engine.ScriptBase.Components; 
using System.IO;
using System;
using System.Net;
using System.Net.Security;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;
using LoginPI.Engine.ScriptBase.Constants;

public class Run_Outlook : ScriptBase
{
    // =====================================================
    // Import and Constants
    // =====================================================
    // Import the user32.dll function to simulate mouse events.
    [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
    public static extern void mouse_event(uint dwFlags, uint dx, uint dy, int dwData, UIntPtr dwExtraInfo);
    public const uint MOUSEEVENTF_WHEEL = 0x0800; // Constant for a mouse wheel event

    // =====================================================
    // Configurable Variables
    // =====================================================
    // Global timings and speeds
    int globalTimeoutInSeconds = 60;                // How long to wait for actions
    int globalWaitInSeconds = 3;                    // Wait time between actions
    int waitMessageboxInSeconds = 2;                // Duration for onscreen wait messages
    int keyboardShortcutsCPM = 15;                  // Typing speed for keyboard shortcuts
    int waitInBetweenKeyboardShortcuts = 4;         // Wait time between keyboard shortcuts
    int typingTextCharacterPerMinute = 600;         // Typing speed for email body text
    int copyPasteRepetitions = 2;                   // Number of times to copy-paste email body content
    int startMenuWaitInSeconds = 5;                 // Duration for Start Menu wait between interactions
    int waitOpenExistingEmail = 8;                  // Wait time for opening an existing email after inbox scrolling
    int waitBeforePictureInsert = 8;                // Wait time before inserting a picture in the email

    // Scrolling parameters for navigating emails
    int inboxDownRepeat = 5;                        // How many times to press DOWN in the inbox list
    int inboxUpRepeat = 5;                          // How many times to press UP in the inbox list
    int existingEmailScrollDownCount = 10;          // Mouse scroll count for an open email
    int existingEmailScrollUpCount = 10;
    int newEmailScrollDownCount = 40;               // Mouse scroll count for a composing email
    int newEmailScrollUpCount = 40;

    // Email composition parameters
    string toFieldText = "LoginEnterpriseVirtualUser1@LoginVSI.com;LoginEnterpriseVirtualUser2@LoginVSI.com;";
    string emailSubject = "About Login VSI";
    // Define the email body content as an array of lines.
    string[] emailBodyLines = new string[]
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

    // =====================================================
    // Helper Methods
    // =====================================================
    // Performs an inline mouse scroll action.
    private void InlineScroll(string direction, int scrollCount, int notches, double waitTime)
    {
        if (waitTime <= 0)
            throw new ArgumentException("Scroll waitTime must be > 0 seconds.");

        int sign = direction.Equals("Down", StringComparison.OrdinalIgnoreCase) ? -1 : 1;
        int delta = sign * 120 * notches;

        Log($"Scrolling mouse {direction} {scrollCount} times, {notches} notch(es) each, {waitTime}s between.");
        for (int i = 0; i < scrollCount; i++)
        {
            mouse_event(MOUSEEVENTF_WHEEL, 0, 0, delta, UIntPtr.Zero);
            Wait(seconds: waitTime, showOnScreen: false);
        }
    }

    // Dismisses the Reminder popup if it appears.
    private void DismissReminderPopup(IWindow mainWindow, int waitTime)
    {
        var reminderWindow = FindWindow(className: "Win32 Window:#32770", title: "*Reminder(s)", processName: "OUTLOOK", timeout: 3, continueOnError: true);
        if (reminderWindow != null)
        {
            Wait(seconds: waitTime, showOnScreen: true, onScreenText: "Dismissing Reminder");
            // Try clicking "Dismiss &All"
            var dismissAllBtn = reminderWindow.FindControl(className: "Button:Button", title: "Dismiss &All", continueOnError: true, timeout: 3);
            if (dismissAllBtn != null)
            {
                reminderWindow.Focus();
                Wait(seconds: waitTime, showOnScreen: true, onScreenText: "Clicking Dismiss &All");
                dismissAllBtn.Click();
            }
            // Try clicking "Yes"
            var yesBtn = reminderWindow.FindControl(className: "Button:Button", title: "&Yes", continueOnError: true, timeout: 2);
            if (yesBtn != null)
            {
                reminderWindow.Focus();
                Wait(seconds: waitTime, showOnScreen: true, onScreenText: "Clicking Yes");
                yesBtn.Click();
                mainWindow.Focus();
                mainWindow.Maximize();
            }
        }
    }

    // Dismiss the modal Outlook dialog that asks "Outlook closed while you had items open. Reopen those items from your last session?".
    void DismissOutlookClosedItemsDialog()
    {
        int timeoutSeconds = 2;
        // Look for the dialog with the given class and title.
        var outlookDialog = FindWindow(className: "Win32 Window:NUIDialog", title: "Microsoft Outlook", processName: "OUTLOOK", timeout: timeoutSeconds, continueOnError: true);
        if (outlookDialog != null)
        {
            // Look for the No button in that dialog.
            var noButton = outlookDialog.FindControl(className: "Button:NetUIButton", title: "No", timeout: timeoutSeconds, continueOnError: true);
            if (noButton != null)
            {
                Wait(globalWaitInSeconds);
                outlookDialog.Focus();
                outlookDialog.Maximize();
                Wait(globalWaitInSeconds);
                // Send the Escape key to dismiss the dialog.
                outlookDialog.Type("{ESC}", hideInLogging: false);
                Wait(globalWaitInSeconds);
            }
        }
    }

    // =====================================================
    // Execute Method
    // =====================================================
    private void Execute()
    {
        // =====================================================
        // Setup and File Download
        // =====================================================
        var temp = GetEnvironmentVariable("TEMP");
        string outlookDir = Path.Combine(temp, "LoginEnterprise");
        if (!Directory.Exists(outlookDir))
        {
            Directory.CreateDirectory(outlookDir);
            Log("Created directory: " + outlookDir);
        }

        // ----- Copy BMP from Appliance scriptcontent if needed -----
        string bmpFile = Path.Combine(outlookDir, "LoginVSI_BattlingRobots.bmp");
        if (!FileExists(bmpFile))
        {
            Log("Copying BMP from Appliance ScriptContent");
            CopyFile(
            sourcePath      : UrnBaseForFiles.UrnBase + "LoginVSI_BattlingRobots.bmp",
            destinationPath : bmpFile
            );
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
        Type("{LWIN}", hideInLogging: false);
        Wait(startMenuWaitInSeconds);
        Type("{LWIN}", hideInLogging: false);
        Wait(1);
        Type("{ESC}", hideInLogging: false);
        Wait(startMenuWaitInSeconds);

        // =====================================================
        // Bring Outlook to Focus and Dismiss Popups
        // =====================================================
        Wait(seconds: waitMessageboxInSeconds, showOnScreen: true, onScreenText: "Dismiss any Outlook popups");
        var mainWindow = FindWindow(title: "Inbox -*", processName: "OUTLOOK", timeout: globalTimeoutInSeconds);
        mainWindow.Focus();
        mainWindow.Maximize();
        Wait(seconds: globalWaitInSeconds, showOnScreen: true, onScreenText: "Focusing Outlook");

        // =====================================================
        // Refresh Outlook Main Window
        // =====================================================
        mainWindow.Minimize();
        Wait(seconds: globalWaitInSeconds, showOnScreen: true, onScreenText: "Minimizing Outlook");
        mainWindow.Maximize();
        mainWindow.Focus();
        Wait(seconds: globalWaitInSeconds, showOnScreen: true, onScreenText: "Maximizing Outlook");

        // Dismiss the Outlook Closed Items dialog, if it appears.
        DismissOutlookClosedItemsDialog();

        // Dismiss the Reminder popup if it exists.
        DismissReminderPopup(mainWindow, globalWaitInSeconds);

        // =====================================================
        // Open and Process an Existing Email
        // =====================================================
        Wait(seconds: waitMessageboxInSeconds, showOnScreen: true, onScreenText: "Scroll through existing emails; open one and scroll in it");
        var inboxWindow = mainWindow.FindControlWithXPath("Table:SuperGrid", timeout: globalTimeoutInSeconds);
        Log("Reading an existing email");
        inboxWindow.Focus();
        inboxWindow.Maximize();
        Wait(seconds: globalWaitInSeconds, showOnScreen: true, onScreenText: "Opening email");
        inboxWindow.Click();
        Wait(seconds: globalWaitInSeconds, showOnScreen: true, onScreenText: "Navigating inbox");
        inboxWindow.Type("{DOWN}".Repeat(inboxDownRepeat), cpm: keyboardShortcutsCPM, hideInLogging: false);
        inboxWindow.Type("{UP}".Repeat(inboxUpRepeat), cpm: keyboardShortcutsCPM, hideInLogging: false);
        Wait(seconds: waitOpenExistingEmail);
        inboxWindow.Type("{ENTER}", cpm: keyboardShortcutsCPM, hideInLogging: false);

        StartTimer("Open_Existing_Email");
        var openEmail = FindWindow(className: "Win32 Window:rctrl_renwnd32", title: "Login Enterprise *", processName: "OUTLOOK", timeout: globalTimeoutInSeconds);
        StopTimer("Open_Existing_Email");
        Wait(seconds: globalWaitInSeconds, showOnScreen: true, onScreenText: "Opening email");
        openEmail.Focus();
        openEmail.Maximize();
        Wait(seconds: globalWaitInSeconds);
        openEmail.MoveMouseToCenter();
        openEmail.Click();

        // Scroll within the opened email.
        InlineScroll("Down", existingEmailScrollDownCount, 1, 0.1);
        InlineScroll("Up", existingEmailScrollUpCount, 1, 0.1);
        InlineScroll("Down", existingEmailScrollDownCount, 1, 0.1);
        InlineScroll("Up", existingEmailScrollUpCount, 1, 0.1);
        openEmail.Close();

        // =====================================================
        // Compose a New Email
        // =====================================================
        Wait(seconds: waitMessageboxInSeconds, showOnScreen: true, onScreenText: "Compose new email with text, attachment, image; do scrolling");
        Log("Initiating new email composition");
        mainWindow.Type("{CTRL+N}", cpm: keyboardShortcutsCPM, hideInLogging: false);
        StartTimer("New_Email_Open");
        var newEmail = FindWindow(className: "Win32 Window:rctrl_renwnd32", title: "Untitled *", processName: "OUTLOOK", timeout: globalTimeoutInSeconds);
        StopTimer("New_Email_Open");
        Wait(seconds: globalWaitInSeconds, showOnScreen: true, onScreenText: "Opening new email");
        newEmail.Focus();
        newEmail.Maximize();
        Wait(seconds: globalWaitInSeconds, showOnScreen: true, onScreenText: "New email window ready");

        // Populate email fields (To, Subject, and Email Body).
        Log("Populating email fields");
        var toField = newEmail.FindControl(className: "*RichEdit20WPT", title: "To", timeout: globalTimeoutInSeconds);
        Wait(seconds: globalWaitInSeconds, showOnScreen: true, onScreenText: "Setting To field");
        toField.Type(toFieldText, cpm: typingTextCharacterPerMinute, hideInLogging: false);
        newEmail.Type("{ALT+U}", cpm: keyboardShortcutsCPM, hideInLogging: false);
        Wait(seconds: waitInBetweenKeyboardShortcuts);
        newEmail.Type(emailSubject, cpm: typingTextCharacterPerMinute, hideInLogging: false);
        newEmail.Type("{TAB}", cpm: keyboardShortcutsCPM, hideInLogging: false);
        Wait(seconds: globalWaitInSeconds, showOnScreen: true, onScreenText: "Entering email body");

        // Type each line of the email body.
        foreach (var line in emailBodyLines)
        {
            newEmail.Type(line, cpm: typingTextCharacterPerMinute, hideInLogging: false);
            newEmail.Type("{Enter}", cpm: typingTextCharacterPerMinute, hideInLogging: false);
        }
        Wait(seconds: globalWaitInSeconds, showOnScreen: true, onScreenText: "Email body entered");

        // =====================================================
        // Insert BMP Image into the Email Body
        // =====================================================
        Log("Inserting BMP image into email body");
        newEmail.Type("{ALT}np", cpm: keyboardShortcutsCPM, hideInLogging: false); // Insert -> Picture command.
        StartTimer("Insert_Picture_Dialog");
        var insertPictureDialog = FindWindow(className: "Win32 Window:#32770", title: "Insert Picture", processName: "OUTLOOK", timeout: globalTimeoutInSeconds);
        StopTimer("Insert_Picture_Dialog");
        var insertPictureDialogFileName = insertPictureDialog.FindControl(className: "Edit:Edit", title: "File name:", timeout: globalTimeoutInSeconds);
        Wait(seconds: globalWaitInSeconds, showOnScreen: true, onScreenText: "Focusing file name box");
        insertPictureDialog.Type("{alt+n}", hideInLogging: false);
        insertPictureDialogFileName.Click();
        Wait(seconds: globalWaitInSeconds, showOnScreen: true, onScreenText: "Typing BMP file path");
        ScriptHelpers.SetTextBoxText(this, insertPictureDialogFileName, bmpFile, cpm: typingTextCharacterPerMinute);
        Type("{ENTER}", hideInLogging: false);
        Wait(seconds: globalWaitInSeconds);

        // =====================================================
        // Expand Email Content via Copy & Paste
        // =====================================================
        Log("Copying and pasting email body content");
        newEmail.Type("{CTRL+A}{CTRL+C}", cpm: keyboardShortcutsCPM, hideInLogging: false);
        Wait(seconds: globalWaitInSeconds, showOnScreen: true, onScreenText: "Selecting email content; copying content");
        for (int i = 0; i < copyPasteRepetitions; i++)
        {
            newEmail.Type("{CTRL+V}", cpm: keyboardShortcutsCPM, hideInLogging: false);
            Wait(seconds: waitInBetweenKeyboardShortcuts);
        }

        // =====================================================
        // Scroll Within the New Email Window
        // =====================================================
        Log("Scrolling within the new email window");
        newEmail.Focus();
        newEmail.Maximize();
        Wait(seconds: globalWaitInSeconds, showOnScreen: true, onScreenText: "Preparing to scroll");
        newEmail.MoveMouseToCenter();
        newEmail.Click();
        Wait(seconds: globalWaitInSeconds, showOnScreen: true, onScreenText: "Scrolling email");
        InlineScroll("Up", newEmailScrollUpCount, 1, 0.1);
        InlineScroll("Down", newEmailScrollDownCount, 1, 0.1);
        Wait(seconds: globalWaitInSeconds, showOnScreen: true, onScreenText: "Scrolling complete");

        // =====================================================
        // Attach BMP File to the Email
        // =====================================================
        Log("Attaching BMP file to email");
        newEmail.Type("{ALT}naf", cpm: keyboardShortcutsCPM, hideInLogging: false); // Insert and Attach command.
        Wait(waitBeforePictureInsert);
        newEmail.Type("b", cpm: keyboardShortcutsCPM, hideInLogging: false);           // 'b' for browse.
        StartTimer("Add_Attachment_Dialog");
        var addAttachmentDialog = FindWindow(className: "Win32 Window:#32770", title: "Insert File", processName: "OUTLOOK", timeout: globalTimeoutInSeconds);
        StopTimer("Add_Attachment_Dialog");
        Wait(seconds: globalWaitInSeconds, showOnScreen: true, onScreenText: "Attachment dialog ready");
        addAttachmentDialog.Focus();
        Type("{alt+n}", hideInLogging: false);
        Wait(globalWaitInSeconds);
        ScriptHelpers.SetTextBoxText(this, addAttachmentDialog.FindControl(className: "Edit:Edit", title: "File name:", timeout: globalTimeoutInSeconds), bmpFile, cpm: typingTextCharacterPerMinute);
        Type("{ENTER}", hideInLogging: false);
        Wait(seconds: globalWaitInSeconds, showOnScreen: true, onScreenText: "Attachment added");

        // =====================================================
        // Finalize Email and Clean Up
        // =====================================================
        try
        {
            newEmail.Close();
            Wait(seconds: globalWaitInSeconds, showOnScreen: true, onScreenText: "Finalizing email");
        
            // Check if the confirmation message window exists with title "*Message*"
            var msgWindow = FindWindow(className: "Win32 Window:rctrl_renwnd32", title: "*Message*", processName: "OUTLOOK", timeout: 2, continueOnError: true);
            if (msgWindow != null)
            {
                msgWindow.Type("{ALT+N}", hideInLogging: false);
                Wait(seconds: globalWaitInSeconds, showOnScreen: true, onScreenText: "Finalizing after confirmation");
        
                // Check again if the message window still exists
                msgWindow = FindWindow(className: "Win32 Window:rctrl_renwnd32", title: "*Message*", processName: "OUTLOOK", timeout: 2, continueOnError: true);
                if (msgWindow != null)
                {
                    ABORT("The New email Outlook window is still existing");
                }
            }
        }
        catch (Exception ex)
        {
            ABORT("Error finalizing new email: " + ex.Message);
        }
        
        Log("Outlook prepared for next iteration. Main window persists; all other windows closed.");
    }
}

// =====================================================
// Helper Class for TextBox Operations
// =====================================================
public static class ScriptHelpers
{
    public static void SetTextBoxText(ScriptBase script, IWindow textBox, string text, int cpm = 600)
    {
        int globalWaitInSeconds = 3; // Wait time between actions
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
