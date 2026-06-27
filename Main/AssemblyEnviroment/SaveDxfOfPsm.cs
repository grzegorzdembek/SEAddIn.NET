using SolidEdgeAdd_In.Helpers.AssemblyEnviroment;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace SolidEdgeAdd_In.Main.AssemblyEnviroment
{
    public class SaveDxfOfPsm
    {
        public static void AddIn(SolidEdgeAssembly.AssemblyDocument assembly)
        {
            try
            {
                string location = SaveDxfOfPsmHelper.GetLocation(assembly);
                var metalSheets = SaveDxfOfPsmHelper.GetMetalSheets(assembly);
                List<string> dxfs = SaveDxfOfPsmHelper.SaveAndGetDxfs(assembly, metalSheets, location);
                SaveDxfOfPsmHelper.CopyDxfs(assembly, dxfs);
            }
            catch (Exception ex) { MessageBox.Show($"Exception: {ex.Message}"); }
        }
    }

}