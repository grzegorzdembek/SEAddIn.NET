using SolidEdgeAdd_In.Helpers.AssemblyEnviroment; 
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace SolidEdgeAdd_In.Main.AssemblyEnviroment
{
    public class SaveDxfOfPartsAndPsm
    {
        public static void AddIn(SolidEdgeAssembly.AssemblyDocument assembly)
        {
            try
            {
                string location = SaveDxfOfPartAndPsmHelper.GetLocation(assembly);
                var partsAndMetalSheets = SaveDxfOfPartAndPsmHelper.GetPartsAndMetalSheets(assembly);
                List<string> dxfs = SaveDxfOfPartAndPsmHelper.SaveAndGetDxfs(assembly, partsAndMetalSheets, location);
                SaveDxfOfPartAndPsmHelper.CopyDxfs(assembly, dxfs);
            }
            catch (Exception ex) { MessageBox.Show($"Exception: {ex.Message}"); }
        }
    }
}