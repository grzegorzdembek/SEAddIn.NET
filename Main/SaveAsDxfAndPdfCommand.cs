using SolidEdgeAdd_In.Helpers;

namespace SolidEdgeAdd_In.Main
{
    public class SaveAsDxfAndPdfCommand
    {
        public static void Execute(SeDraft draft)
        {
            try
            {
                var (isValid, dxfPath, pdfPath) = SaveAsDxfAndPdfHelper.Paths(draft);

                if (isValid)
                {
                    SaveAsDxfAndPdfHelper.Save(draft, dxfPath, pdfPath);
                }
            }
            catch (Exception ex) { MessageBox.Show($"Exception: {ex.Message}"); }
        }
    }
}