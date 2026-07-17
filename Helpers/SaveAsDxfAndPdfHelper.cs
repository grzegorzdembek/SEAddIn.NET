namespace SolidEdgeAdd_In.Helpers
{
    public class SaveAsDxfAndPdfHelper
    {
        public static (bool isValid, string dxfPath, string pdfPath) Paths(SeDraft draft)
        {
            if (string.IsNullOrEmpty(draft.FullName))
            {
                MessageBox.Show("Zapisz najpierw rysunek, aby móc wyeksportować formaty DXF i PDF.", "Wymagany Zapis", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return (false, null, null);
            }

            string directory = Path.GetDirectoryName(draft.FullName);
            string fileName = Path.GetFileNameWithoutExtension(draft.FullName);

            string dxfPath = Path.Combine(directory, fileName + ".dxf");
            string pdfPath = Path.Combine(directory, fileName + ".pdf");

            return (true, dxfPath, pdfPath);
        }

        public static void Save(SeDraft draft, string dxfPath, string pdfPath)
        {
            draft.SaveAs(dxfPath, 14);
            draft.SaveAs(pdfPath, 5);
        }
    }
}