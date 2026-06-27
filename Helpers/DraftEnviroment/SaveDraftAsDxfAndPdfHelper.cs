using System.IO;

namespace SolidEdgeAdd_In.Helpers.DraftEnviroment
{
    public class SaveDraftAsDxfAndPdfHelper
    {
        public static (string, string)
            Paths
            (SolidEdgeDraft.DraftDocument draft)
        {
            string dxfPath = Path.Combine(Path.GetDirectoryName(draft.FullName), Path.GetFileNameWithoutExtension(draft.FullName) + ".dxf");
            string pdfPath = Path.Combine(Path.GetDirectoryName(draft.FullName), Path.GetFileNameWithoutExtension(draft.FullName) + ".pdf");
            return (dxfPath, pdfPath);
        }

        public static void
            Save
            (SolidEdgeDraft.DraftDocument draft, string dxfPath, string pdfPath)
        {
            draft.SaveAs(dxfPath, 14);
            draft.SaveAs(pdfPath, 5);
        }
    }
}