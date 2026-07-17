using SolidEdgeAdd_In.Helpers;

namespace SolidEdgeAdd_In.Main
{
    public class ExportPartsListCommand
    {
        public static void Execute(SeAssembly assembly)
        {
            SeApp application = assembly.Application;
            ExcelApp excelApp = null;
            ExcelWorkbooks workbooks = null;
            ExcelWorkbook workbook = null;
            ExcelSheets sheets = null;
            ExcelWorksheet worksheet = null;

            try
            {
                application.DelayCompute = true;
                application.Interactive = false;

                var (isConfirmed, multiplier) = ExportPartsListHelper.GetMultiplier(assembly);

                if (!isConfirmed) return;

                ExportPartsListHelper.CopyPartsList(assembly);
                bool hasShots = ExportPartsListHelper.HasShots();
                var shots = ExportPartsListHelper.GetShots(assembly, hasShots);
                ExportPartsListHelper.ExcelObjects(out excelApp, out workbooks, out workbook, out sheets, out worksheet);
                ExportPartsListHelper.EditWorksheet(assembly, shots, hasShots, workbook, worksheet, multiplier);
                ExportPartsListHelper.Export(assembly, excelApp, workbooks, workbook, worksheet);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Exception: {ex.Message}");
            }
            finally
            {
                application.DelayCompute = false;
                application.Interactive = true;
                ExportPartsListHelper.Release(ref excelApp, ref workbooks, ref workbook, ref sheets, ref worksheet);
            }
        }
    }
}