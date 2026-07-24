using Helper = SolidEdgeAdd_In.Helpers.SaveAsDxfAndPdfHelper;

namespace SolidEdgeAdd_In.Commands
{
    public class SaveAsDxfAndPdfCommand
    {
        public static void Execute(SeDraft draft)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();
            try
            {
                var (isValid, dxfPath, pdfPath) = Helper.Paths(draft); if (isValid) { Helper.Save(draft, dxfPath, pdfPath); }

                stopwatch.Stop(); string elapsedTime = stopwatch.Elapsed.ToString(@"mm\:ss\.fff"); MessageBox.Show($"Czas wykonywania: {elapsedTime}", "Zakończono", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex) { MessageBox.Show($"Exception: {ex.Message}"); }
        }
    }
}