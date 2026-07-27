using Helper = SolidEdgeAdd_In.Helpers.ExportPartsListHelper;

namespace SolidEdgeAdd_In.Commands
{
    public class ExportPartsListCommand
    {
        public static void Execute(SeAssembly assembly)
        {
            Stopwatch stopwatch = Stopwatch.StartNew(); SeApp application = assembly.Application; ExcelApp excelApp = null; ExcelWorkbooks workbooks = null; ExcelWorkbook workbook = null; ExcelSheets sheets = null; ExcelWorksheet worksheet = null;

            try
            {
                application.DelayCompute = true;

                var (isConfirmed, multiplier) = Helper.GetMultiplier(assembly); if (!isConfirmed) return;

                Helper.CopyPartsList(assembly); bool hasShots = Helper.HasShots(); var shots = Helper.GetShots(assembly, hasShots);

                Helper.ExcelObjects(out excelApp, out workbooks, out workbook, out sheets, out worksheet); Helper.EditWorksheet(assembly, shots, hasShots, workbook, worksheet, multiplier); Helper.Export(assembly, excelApp, workbooks, workbook, worksheet);

                stopwatch.Stop(); string elapsedTime = stopwatch.Elapsed.ToString(@"mm\:ss\.fff"); MessageBox.Show($"Czas wykonywania: {elapsedTime}", "Zakończono", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex) { MessageBox.Show($"Exception: {ex.Message}"); }
            finally { application.DelayCompute = false; Helper.Release(ref excelApp, ref workbooks, ref workbook, ref sheets, ref worksheet); }
        }
    }
}