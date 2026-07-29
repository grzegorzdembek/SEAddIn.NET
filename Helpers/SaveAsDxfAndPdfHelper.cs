namespace SolidEdgeAdd_In.Helpers
{
    public class SaveAsDxfAndPdfHelper
    {
        public static (bool isSaved, string dxfPath, string pdfPath) GetData(SeDraft draft)
        {
            if (string.IsNullOrEmpty(draft.FullName))
            { 
                MessageBox.Show("Zapisz najpierw rysunek, aby móc wyeksportować formaty DXF i PDF.", "Wymagany Zapis", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return (false, null, null); 
            }

            string dir = Path.GetDirectoryName(draft.FullName); 
            string file = Path.GetFileNameWithoutExtension(draft.FullName);

            string dxfPath = Path.Combine(dir, file + ".dxf"); 
            string pdfPath = Path.Combine(dir, file + ".pdf");

            return (true, dxfPath, pdfPath);
        }

        public static void Save(SeDraft draft, string dxfPath, string pdfPath) 
        { 
            draft.SaveAs(dxfPath, 14); 
            draft.SaveAs(pdfPath, 5); 
        }
    }
}