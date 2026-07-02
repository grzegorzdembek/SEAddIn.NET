using SolidEdgeAdd_In.Helpers.PartEnviroment; 
using System;
using System.IO;
using System.Windows.Forms;

namespace SolidEdgeAdd_In.Main.PartEnviroment
{
    public class SaveFlatPatternAsDxf
    {
        public static void AddIn(SolidEdgePart.PartDocument part)
        {
            try
            {
                string path = SaveFlatPatternAsDxfHelper.GetPath(part);
                (bool isConfirmed, string editedPath) = SaveFlatPatternAsDxfHelper.GetDecisionAndEditedPath(path);
                SaveFlatPatternAsDxfHelper.Save(part, isConfirmed, editedPath);
            }
            catch (Exception ex) { MessageBox.Show($"Exception: {ex.Message}"); }
        }
    }
}
