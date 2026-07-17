using SolidEdgeAdd_In.Helpers;
using SolidEdgeAdd_In.Utils;

namespace SolidEdgeAdd_In.Main
{
    public class ExportDxfsCommand
    {
        public static void Execute(SeAssembly assembly)
        {
            SeApp application = assembly.Application;
            DxfExportLogger logger = new ();
            string subDirectory = string.Empty;

            try
            {
                application.DelayCompute = true;
                application.Interactive = false;
                application.ScreenUpdating = false;

                var (isConfirmed, multiplier) = ExportDxfsHelper.GetMultiplier(assembly);
                if (!isConfirmed || multiplier<=0) return;
                if (isConfirmed || multiplier>0) MessageBox.Show($"Wybrano mnożnik:{multiplier}", "Informacja", MessageBoxButtons.OK, MessageBoxIcon.Information);

                subDirectory = ExportDxfsHelper.GetSubDirectory(assembly);

                var occurrences = ExportDxfsHelper.GetOccurrences(assembly, logger);

                if (occurrences.Count == 0)
                {
                    MessageBox.Show("Brak blach do przetworzenia", "Informacja", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                ExportDxfsHelper.SaveDxfs(assembly, occurrences, subDirectory, multiplier, logger);

                MessageBox.Show($"Koniec","Sukces", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Błąd w trakcie eksportu DXF: {ex.Message}");
            }
            finally
            {
                application.DelayCompute = false;
                application.Interactive = true;
                application.ScreenUpdating = true;
                if (!string.IsNullOrEmpty(subDirectory) && Directory.Exists(subDirectory))
                {
                    logger.SaveReport(subDirectory);
                }
            }
        }
    }
}