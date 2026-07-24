using SolidEdgeAdd_In.Utils;

namespace SolidEdgeAdd_In.Commands
{
    public class ExportDxfsCommand
    {
        public static void Execute(SeAssembly assembly)
        {
            Stopwatch stopwatch = Stopwatch.StartNew(); SeApp application = assembly.Application; Logger logger = new(); string subDirectory = string.Empty;

            try
            {
                application.DelayCompute = true; application.ScreenUpdating = false;

                var (isConfirmed, multiplier) = ExportDxfsHelper.GetMultiplier(assembly); if (!isConfirmed || multiplier <= 0) return;

                if (isConfirmed || multiplier > 0) MessageBox.Show($"Wybrano mnożnik:{multiplier}", "Informacja", MessageBoxButtons.OK, MessageBoxIcon.Information);

                subDirectory = ExportDxfsHelper.GetSubDirectory(assembly); var occurrences = ExportDxfsHelper.GetData(assembly, logger);

                if (occurrences.Count == 0) { MessageBox.Show("Brak blach do przetworzenia", "Informacja", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }

                ExportDxfsHelper.ProcessData(assembly, occurrences, subDirectory, multiplier, logger);

                stopwatch.Stop(); string elapsedTime = stopwatch.Elapsed.ToString(@"mm\:ss\.fff"); MessageBox.Show($"Czas wykonywania: {elapsedTime}", "Zakończono", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex) { MessageBox.Show($"Exception: {ex.Message}"); }
            finally { application.DelayCompute = false; application.ScreenUpdating = true; if (!string.IsNullOrEmpty(subDirectory) && Directory.Exists(subDirectory)) { logger.SaveReport(subDirectory); } }
        }
    }
}