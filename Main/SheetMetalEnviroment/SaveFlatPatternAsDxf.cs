using SolidEdgeAdd_In.Helpers.SheetMetalEnviroment; 
using System;
using System.Windows.Forms;

namespace SolidEdgeAdd_In.Main.SheetMetalEnviroment
{
    public class SaveFlatPatternAsDxf
    {
        public static void AddIn(SolidEdgePart.SheetMetalDocument metalSheet)
        {
            try
            {
                string path = SaveFlatPatternAsDxfHelper.GetPath(metalSheet);
                (bool isConfirmed, string editedPath) = SaveFlatPatternAsDxfHelper.GetDecisionAndEditedPath(path);
                SaveFlatPatternAsDxfHelper.Save(metalSheet, isConfirmed, editedPath);
            }
            catch (Exception ex) { MessageBox.Show($"Exception: {ex.Message}"); }
        }
    }
}