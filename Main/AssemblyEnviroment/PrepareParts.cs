using SolidEdgeAdd_In.Helpers.AssemblyEnviroment;
using System;
using System.Windows.Forms;

namespace SolidEdgeAdd_In.Main.AssemblyEnviroment
{
    public class PrepareParts
    {
        public static void AddIn(SolidEdgeAssembly.AssemblyDocument assembly)
        {
            try
            {
                if (!PreparePartsHelper.IsConfirmedByUser()) return;
                var partsAndMetalSheets = PreparePartsHelper.GetParts(assembly);
                var (filesCount, proccessedFilesCount) = PreparePartsHelper.PrepareAndGetStats(assembly, partsAndMetalSheets);
                PreparePartsHelper.Report(filesCount, proccessedFilesCount);
            }
            catch (Exception ex) { MessageBox.Show($"Exception: {ex.Message}"); }
        }
    }
}
