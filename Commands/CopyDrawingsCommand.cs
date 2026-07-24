using Helper = SolidEdgeAdd_In.Helpers.CopyDrawingsHelper;

namespace SolidEdgeAdd_In.Commands
{
    public class CopyDrawingsCommand
    {
        public static void Execute(SeDocument document)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();

            try
            {
                var defaultDir = Helper.GetDefaultDirectory(document); if (string.IsNullOrEmpty(defaultDir)) { MessageBox.Show("Nie znaleziono folderu Paczki", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

                var selectedDir = Helper.GetSelectedDirectory(defaultDir); if (string.IsNullOrEmpty(selectedDir)) return;

                var excelSummary = Helper.GetExcelSummary(selectedDir); if (string.IsNullOrEmpty(excelSummary)) { MessageBox.Show("Brakuje zestawienia blach.", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

                Helper.ProcessCopyingDrawings(defaultDir, selectedDir, excelSummary);

                stopwatch.Stop(); string elapsedTime = stopwatch.Elapsed.ToString(@"mm\:ss\.fff"); MessageBox.Show($"Czas działania: {elapsedTime}", "Zakończono", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex) { MessageBox.Show($"Exception: {ex.Message}"); }
        }
    }
}
