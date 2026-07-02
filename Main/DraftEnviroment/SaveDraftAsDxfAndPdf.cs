using SolidEdgeAdd_In.Helpers.DraftEnviroment; 
using System;
using System.Windows.Forms;

namespace SolidEdgeAdd_In.Main.DraftEnviroment
{
    public class SaveDraftAsDxfAndPdf
    {
        public static void AddIn(SolidEdgeDraft.DraftDocument draft)
        {
            try
            {
                (string dxfPath, string pdfPath) = SaveDraftAsDxfAndPdfHelper.Paths(draft);
                SaveDraftAsDxfAndPdfHelper.Save(draft, dxfPath, pdfPath);
            }
            catch (Exception ex) { MessageBox.Show($"Exception: {ex.Message}"); }
        }
    }
}
