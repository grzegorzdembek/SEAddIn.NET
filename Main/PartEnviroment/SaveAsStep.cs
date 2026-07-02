using SolidEdgeAdd_In.Helpers.PartEnviroment;
using System;
using System.Windows.Forms;

namespace SolidEdgeAdd_In.Main.PartEnviroment
{
    public class SaveAsStep
    {
        public static void AddIn(SolidEdgePart.PartDocument part)
        {
            try
            {
                string path = SaveAsStepHelper.GetPath(part);
                (bool isConfirmed, string editedPath) = SaveAsStepHelper.GetDecisionAndEditedPath(path);
                SaveAsStepHelper.Save(part, isConfirmed, editedPath);
            }
            catch (Exception ex) { MessageBox.Show($"Exception: {ex.Message}"); }
        }
    }
}
