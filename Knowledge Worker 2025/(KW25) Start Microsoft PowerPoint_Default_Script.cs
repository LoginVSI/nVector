// TARGET:powerpnt.exe
// START_IN:

/////////////
// PowerPoint Start
// Workload: KnowledgeWorker 2025
// Version: 1.0
/////////////

using LoginPI.Engine.ScriptBase;

public class Start_PowerPoint_DefaultScript : ScriptBase
{
    void Execute()
    {
        START(mainWindowTitle:"*PowerPoint*", mainWindowClass:"*PPTFrameClass*", timeout:5);

        STOP();
    }
}