using SolidEdgeAdd_In.Helpers.AssemblyEnviroment; 
using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace SolidEdgeAdd_In.Main.AssemblyEnviroment
{
    public class ExportPartsList
    {
        public static void AddIn(SolidEdgeAssembly.AssemblyDocument assembly)
        {
            Excel.Application excelApp = null;
            Excel.Workbooks workbooks = null;
            Excel.Workbook workbook = null;
            Excel.Sheets sheets = null;
            Excel.Worksheet worksheet = null;
            try
            {
                int mulitplier = ExportPartsListHelper.GetMultiplier(assembly);
                ExportPartsListHelper.CopyPartsList(assembly);
                bool hasShots = ExportPartsListHelper.HasShots();
                var shots = ExportPartsListHelper.GetShots(assembly, hasShots);
                ExportPartsListHelper.ExcelObjects(out excelApp, out workbooks, out workbook, out sheets, out worksheet);
                ExportPartsListHelper.EditWorksheet(assembly, shots, hasShots, workbook, worksheet, mulitplier);
                ExportPartsListHelper.Export(assembly, excelApp, workbooks, workbook, worksheet);
            }
            catch (Exception ex) { MessageBox.Show($"Exception: {ex.Message}"); }
            finally { ExportPartsListHelper.Release(ref excelApp, ref workbooks, ref workbook, ref sheets, ref worksheet); }
        }
    }
}