using System;

namespace Autosave
{
    public static class GetAcadApplication
    {
        public static void GetNC(ref nanoCAD.Application NC)
        {
            try
            {
                Type t1 = Type.GetTypeFromProgID("nanoCAD.Application");
                NC = (nanoCAD.Application)Activator.CreateInstance(t1);
            }
            catch { }            
        }
    }
}
