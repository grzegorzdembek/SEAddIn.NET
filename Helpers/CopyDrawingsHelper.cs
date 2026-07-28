using SolidEdgeAdd_In.Utils;

namespace SolidEdgeAdd_In.Helpers
{
    public class CopyDrawingsHelper
    {
        public static string GetDefaultDirectory(SeDocument document)
        {
            if (document == null || string.IsNullOrEmpty(document.FullName)) return null;

            string path = document.FullName; string projectDirectory = Path.GetDirectoryName(path);

            string packagesDirectory = Path.Combine(projectDirectory, Constants.Folders.Packages); if (!Directory.Exists(packagesDirectory)) return null;

            return packagesDirectory;
        }

        public static string GetSelectedDirectory(string defaultDirectory)
        {
            using FolderBrowserDialog fbd = new() { Description = "Wybierz folder, w którym chcesz dodać folder z Rysunkami (zawierający zestawienie blach):", SelectedPath = defaultDirectory, ShowNewFolderButton = false };

            if (fbd.ShowDialog() == DialogResult.OK) return fbd.SelectedPath;

            return null;
        }

        public static string GetExcelSummary(string selectedDirectory)
        {
            string[] excelFiles = Directory.GetFiles(selectedDirectory, "*.xlsx"); if (excelFiles.Length == 0 || excelFiles.Length > 1) return null;

            return excelFiles[0];
        }

        public static void ProcessCopyingDrawings(SeDocument document, string selectedDir, string excelPath)
        {
            if (document == null || string.IsNullOrEmpty(document.FullName)) return ;

            string path = document.FullName; string projectDirectory = Path.GetDirectoryName(path);

            string drawingsDirectory = Path.Combine(selectedDir, Constants.Folders.Drawings); if (!Directory.Exists(drawingsDirectory)) Directory.CreateDirectory(drawingsDirectory);

            var PdfFiles = Directory.GetFiles(projectDirectory, "*.pdf", SearchOption.TopDirectoryOnly); var DxfFiles = Directory.GetFiles(projectDirectory, "*.dxf", SearchOption.TopDirectoryOnly);

            var pdfNames = PdfFiles.Select(f => new { Path = f, Name = Path.GetFileNameWithoutExtension(f) }).ToList(); var dxfNames = DxfFiles.Select(f => new { Path = f, Name = Path.GetFileNameWithoutExtension(f) }).ToList();

            ExcelApp excelApp = null; ExcelWorkbooks workbooks = null; ExcelWorkbook workbook = null; ExcelSheets sheets = null; ExcelWorksheet worksheet = null; ExcelRange usedRange = null; ExcelRange expandedRange = null;
            ExcelRange cells = null; ExcelRange columns = null; ExcelRange rows = null; ExcelRange headerCell = null; dynamic headerFont = null; dynamic headerInterior = null; dynamic expandedBorders = null; ExcelRange startCell = null; ExcelRange endCell = null;

            try
            {
                excelApp = new ExcelApp { Visible = false, DisplayAlerts = false, ScreenUpdating = false, EnableEvents = false };

                workbooks = excelApp.Workbooks; workbook = workbooks.Open(excelPath, ReadOnly: false); sheets = workbook.Sheets; worksheet = (ExcelWorksheet)sheets[1];

                usedRange = worksheet.UsedRange; cells = worksheet.Cells; rows = usedRange.Rows; columns = usedRange.Columns;

                int nameColIndex = ExcelWrapper.GetColumnNumber(usedRange, Constants.ExcelHeaders.PartNumber); if (nameColIndex == 0) { MessageBox.Show("Nie znaleziono kolumny w pliku Excel.", "Błąd"); return; }

                int rowCount = rows.Count; int colCount = columns.Count;

                int drawingsColIndex = ExcelWrapper.GetColumnNumber(usedRange, Constants.ExcelHeaders.Drawings);

                bool isNewColumn = false;

                if (drawingsColIndex == 0) { drawingsColIndex = colCount + 1; isNewColumn = true; }

                try { startCell = (ExcelRange)cells[1, 1]; endCell = (ExcelRange)cells[rowCount, drawingsColIndex]; expandedRange = worksheet.Range[startCell, endCell]; }
                finally { CoreUtils.ReleaseCom(ref startCell); CoreUtils.ReleaseCom(ref endCell); }

                object[,] data = (object[,])expandedRange.Value2; data[1, drawingsColIndex] = Constants.ExcelHeaders.Drawings;

                for (int i = 2; i <= rowCount; i++)
                {
                    bool isPdfReady = false; bool isDxfReady = false;

                    if (data[i, nameColIndex] != null)
                    {
                        string partName = data[i, nameColIndex].ToString().Trim();

                        if (!string.IsNullOrEmpty(partName))
                        {
                            string expectedPdfPath = Path.Combine(drawingsDirectory, partName + ".pdf"); string expectedDxfPath = Path.Combine(drawingsDirectory, partName + ".dxf");

                            if (File.Exists(expectedPdfPath)) { isPdfReady = true; }
                            else
                            {
                                var matchingPdfs = pdfNames.Where(f => f.Name.Equals(partName, StringComparison.OrdinalIgnoreCase)).ToList();
                                foreach (var pdf in matchingPdfs) { File.Copy(pdf.Path, expectedPdfPath, true); isPdfReady = true; }
                            }

                            if (File.Exists(expectedDxfPath)) { isDxfReady = true; }
                            else
                            {
                                var matchingDxfs = dxfNames.Where(f => f.Name.Equals(partName, StringComparison.OrdinalIgnoreCase)).ToList();
                                foreach (var dxf in matchingDxfs) { File.Copy(dxf.Path, expectedDxfPath, true); isDxfReady = true; }
                            }
                        }
                    }

                    if (isDxfReady && isPdfReady) data[i, drawingsColIndex] = "Skopiowany"; else data[i, drawingsColIndex] = "-";
                }

                expandedRange.Value2 = data;

                if (isNewColumn)
                {
                    headerCell = (ExcelRange)cells[1, drawingsColIndex]; headerFont = headerCell.Font; headerFont.Bold = true;

                    headerInterior = headerCell.Interior; headerInterior.Color = ColorTranslator.ToOle(Color.LightGray);

                    expandedRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; expandedBorders = expandedRange.Borders; expandedBorders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    columns.AutoFit();
                }

                workbook.Save();
            }
            finally
            {
                CoreUtils.ReleaseCom(ref expandedBorders); CoreUtils.ReleaseCom(ref headerInterior); CoreUtils.ReleaseCom(ref headerFont); CoreUtils.ReleaseCom(ref headerCell);
                CoreUtils.ReleaseCom(ref rows); CoreUtils.ReleaseCom(ref columns); CoreUtils.ReleaseCom(ref cells); CoreUtils.ReleaseCom(ref expandedRange); CoreUtils.ReleaseCom(ref usedRange);
                CoreUtils.ReleaseCom(ref worksheet); CoreUtils.ReleaseCom(ref sheets); workbook?.Close(false); CoreUtils.ReleaseCom(ref workbook); CoreUtils.ReleaseCom(ref workbooks); excelApp?.Quit(); CoreUtils.ReleaseCom(ref excelApp);
            }
        }
    }
}