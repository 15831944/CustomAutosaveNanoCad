using System;
using System.Windows.Forms;

namespace Autosave
{
    public static class GetAcadApplication
    {
        public static void GetNC(ref nanoCAD.Application NC)
        {
            var p = Application.ExecutablePath;

            System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
            startInfo.CreateNoWindow = false;
            startInfo.UseShellExecute = false;
            startInfo.FileName = p;
            startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
            startInfo.Arguments = "/register";

            System.Diagnostics.Process exeProcess = System.Diagnostics.Process.Start(startInfo);
            //If you need synchronous execution you can do this
            exeProcess.WaitForExit();

            try
            { 
               Type t1 = Type.GetTypeFromProgID("nanoCAD.Application");
               NC = (nanoCAD.Application)Activator.CreateInstance(t1);
            }
            catch {}        
        }
    }
}
