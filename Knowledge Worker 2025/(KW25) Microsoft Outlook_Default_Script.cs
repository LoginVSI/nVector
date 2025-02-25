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
using System.Runtime.InteropServices;

public class M365Outlook_RunningScript : ScriptBase
{
    // Import the user32.dll function to simulate mouse events.
    [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
    public static extern void mouse_event(uint dwFlags, uint dx, uint dy, int dwData, UIntPtr dwExtraInfo);

    public const uint MOUSEEVENTF_WHEEL = 0x0800; // Constant for a mouse wheel event

    private void Execute()
    {
        
                // Simulate user interaction to open the Start Menu.
        Wait(seconds:2, showOnScreen:true, onScreenText:"Start Menu");
        Type("{LWIN}");
        Wait(3);
        Type("{LWIN}");
        Wait(1);
        Type("{ESC}");
        
        // Use a local variable for the main Outlook window.
        
        START();
        var mainWindow = FindWindow(processName:"outlook", timeout:3, continueOnError:true);
        mainWindow.Focus();
        Wait(1);

        // Dismiss the Activate Office popup dialog if it appears.
        try 
        {
            var signinWindow = mainWindow.FindControlWithXPath(xPath:"Win32 Window:NUIDialog", timeout:2);
            signinWindow.Type("{ESC}", cpm:50);
        }
        catch 
        {
            // If the dialog is not found, continue silently.
        }

        // Skip all first-run dialogs.
        SkipFirstRunDialogs();
        
        // Initiate composing a new email.
        var inboxWindow = mainWindow.FindControlWithXPath(xPath:"Table:SuperGrid");
        inboxWindow.Type("{ctrl+n}");
        var newEmailWindow = FindWindow(className:"Win32 Window:rctrl_renwnd32", title:"Untitled - Message (HTML) ", processName:"OUTLOOK");
        newEmailWindow.Focus();
        newEmailWindow.Maximize();
        Wait(1);
        newEmailWindow.FindControl(className:"Edit", title:"Page 1 content").Click();
        newEmailWindow.Type("{alt}naf", cpm:50);
        Wait(1);
        newEmailWindow.Type("b", cpm:50);
        Wait(2);
        Type(@"C:\Users\NVtestuser\Desktop\LoginVSI_BattlingRobots.bmp{enter}", cpm:1000);
        Wait(2);
        newEmailWindow.Type("{alt}n", cpm:50);
        Wait(1);
        newEmailWindow.Type("p", cpm:50);
        Wait(2);
        Type(@"C:\Users\NVtestuser\Desktop\LoginVSI_BattlingRobots.bmp{enter}", cpm:1000);
        Wait(2);
        
        // Scroll interactions on the active tab after switching:
        Scroll("Down", 10, 1, 0.2);
        Scroll("Up", 5, 2, 0.3);
        Wait(1);
                var typingSpeed = 2000;
        var newEmail = FindWindow(className:"Win32 Window:rctrl_renwnd32", title:"Untitled - Message (HTML) ", processName:"OUTLOOK");
        newEmail.Focus();
        newEmail.FindControl(className:"*RichEdit20WPT", title:"To").Type("marx@loginvsi.com; mank@loginvsi.com; blain@loginvsi.com", cpm:typingSpeed);
        newEmail.Type("{TAB}".Repeat(3), 50);
        newEmail.Type("Today's Topics - Words from Vonnegut's 2-B-R-0-2-B", cpm:typingSpeed);
        /*newEmail.Type("{TAB}", cpm:50);
        newEmail.Type("{ENTER}", cpm:50);
        newEmail.Type("{CTRL+B}", cpm:50);
        newEmail.Type("Young Wehling was hunched in his chair, his head in his hand. He was so rumpled, so still and colorless as to be virtually invisible.{ENTER}", cpm:typingSpeed);*/
        Wait(1);
        //newEmailWindow.Type("{esc}");

        // Select an item in the Inbox.
        Wait(seconds:1, showOnScreen:true, onScreenText:"Select An Item");
        inboxWindow.Click();
        Wait(1);

        // Scroll through the e-mail inbox.
        Wait(seconds:1, showOnScreen:true, onScreenText:"Scroll Inbox");
        inboxWindow.Type("{DOWN}".Repeat(3), cpm:350);
        Wait(1);
        inboxWindow.Type("{DOWN}".Repeat(4), cpm:350);
        inboxWindow.Type("{UP}".Repeat(8), cpm:350);
        Wait(1);
        DismissReminders();

        

        /*// Open an email, read it, and close it.
        Wait(seconds:1, showOnScreen:true, onScreenText:"Open and Read an Email");
        inboxWindow.Focus();
        inboxWindow.Click();
        inboxWindow.Type("{DOWN}");
        inboxWindow.Type("{ENTER}");
        Wait(1);
        var openEmail = FindWindow(className:"Win32 Window:rctrl_renwnd32", title:"Login Enterprise Continuity & Application Load Testing - Message (HTML) ", processName:"OUTLOOK");
        openEmail.Focus();
        openEmail.Type("{DOWN}".Repeat(5), cpm:500);
        Wait(1);
        openEmail.Type("{UP}".Repeat(3), cpm:500);
        Wait(1);
        openEmail.Type("{ESC}", cpm:50);
        Wait(1);
        */

        // Minimize and then maximize the main window.
        mainWindow.Minimize();
        Wait(1);
        mainWindow.Maximize();

        // Compose a new email with words from Vonnegut's 2-B-R-0-2-B.
        Wait(seconds:1, showOnScreen:true, onScreenText:"Compose a new email with words from Vonnegut's 2-B-R-0-2-B");
        mainWindow.Type("{CTRL+N}");
        Wait(1);
        /*var typingSpeed = 2000;
        var newEmail = FindWindow(className:"Win32 Window:rctrl_renwnd32", title:"Untitled - Message (HTML) ", processName:"OUTLOOK");
        newEmail.Focus();
        newEmail.FindControl(className:"*RichEdit20WPT", title:"To").Type("marx@loginvsi.com; mank@loginvsi.com; blain@loginvsi.com", cpm:typingSpeed);
        newEmail.Type("{TAB}".Repeat(3), 50);
        newEmail.Type("Today's Topics - Words from Vonnegut's 2-B-R-0-2-B", cpm:typingSpeed);
        newEmail.Type("{TAB}", cpm:50);
        newEmail.Type("{ENTER}", cpm:50);
        newEmail.Type("{CTRL+B}", cpm:50);
        newEmail.Type("Young Wehling was hunched in his chair, his head in his hand. He was so rumpled, so still and colorless as to be virtually invisible.{ENTER}", cpm:typingSpeed);
        Wait(1);
        */
        // End the script.
        // STOP();
    }

    private void DismissReminders()
    {
        var reminderWindow = FindWindow(className:"Win32 Window:#32770", title:"*Reminder(s)", processName:"OUTLOOK", timeout:2, continueOnError:true);
        if (reminderWindow != null)
        {
            Wait(1);
            reminderWindow.Focus();
            reminderWindow.FindControl(className:"Button:Button", title:"Dismiss &All").Click();
            Wait(1);
            reminderWindow.FindControl(className:"Button:Button", title:"&Yes").Click();
        }
    }

    private void SkipFirstRunDialogs()
    {
        var dialog = FindWindow(className:"Win32 Window:NUIDialog", processName:"OUTLOOK", continueOnError:true, timeout:1);
        while (dialog != null)
        {
            dialog.Close();
            dialog = FindWindow(className:"Win32 Window:NUIDialog", processName:"OUTLOOK", continueOnError:true, timeout:10);
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
