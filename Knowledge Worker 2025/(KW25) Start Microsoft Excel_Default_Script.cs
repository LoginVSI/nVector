// TARGET:excel.exe
// START_IN:

/////////////
// Excel Start
// Workload: KnowledgeWorker 2025
// Version: 1.0
/////////////

using LoginPI.Engine.ScriptBase;

public class Start_Excel_DefaultScript : ScriptBase
{
    void Execute()
    {
        START(mainWindowTitle: "*Excel*", mainWindowClass: "*XLMAIN*", timeout: 5);

        STOP();
    }
}