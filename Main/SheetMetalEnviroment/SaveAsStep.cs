using SolidEdgeAdd_In.Helpers.SheetMetalEnviroment;
using System;
using System.Windows.Forms;

namespace SolidEdgeAdd_In.Main.SheetMetalEnviroment
{
    public class SaveAsStep
    {
        public static void AddIn(SolidEdgePart.SheetMetalDocument metalSheet)
        {
            try
            {
                string path = SaveAsStepHelper.GetPath(metalSheet);
                (bool isConfirmed, string editedPath) = SaveAsStepHelper.GetDecisionAndEditedPath(path);
                SaveAsStepHelper.Save(metalSheet, isConfirmed, editedPath);
            }
            catch (Exception ex) { MessageBox.Show($"Exception: {ex.Message}"); }
        }
    }
}
