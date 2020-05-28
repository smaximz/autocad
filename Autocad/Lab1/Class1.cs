using System;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

namespace Lab1
{
    public class Class1
    {
        //define command that can be called inside AutoCAD
        [CommandMethod("HelloWorld")]
        public void HelloWorld()
        {
            // object to print to the AutoCAD' command line
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            ed.WriteMessage("Hello World");
        }
    }
}
