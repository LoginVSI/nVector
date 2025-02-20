// TARGET:winword.exe
// START_IN:

/////////////
// Outlook Start
// Workload: KnowledgeWorker 2025
// Version: 1.0
/////////////

using LoginPI.Engine.ScriptBase;

public class Start_Outlook_DefaultScript : ScriptBase
{
    void Execute()
    {
        START(mainWindowTitle: "*Word*", mainWindowClass: "Win32 Window:OpusApp", processName: "WINWORD", timeout: 5);

        STOP();
    }
}