using HostMgd.ApplicationServices;
using System;

namespace Autosave
{
    public static class GetAcadApplication
    {
        public static void GetNC(ref nanoCAD.Application NC)
        {
            try
            {
                NC = (nanoCAD.Application)Application.AcadApplication;
            }
            catch { }
        }
    }
}
