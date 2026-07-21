using SolidEdgeAdd_In.Utils;

namespace SolidEdgeAdd_In.Helpers
{
    public class ExportDrawingsHelper
    {
        public static string GetDefaultPath(SeDocument document)
        {

            if (document == null || string.IsNullOrEmpty(document.FullName)) return null;

            string path = document.FullName;

            string projectDirectory = Path.GetDirectoryName(path);

            string packagesPath = Path.Combine(projectDirectory, "Paczki");

            if (!Directory.Exists(packagesPath)) return null;

            return packagesPath;
        }

        public static string GetSelectedFolder(string defaultPath)
        {
            using FolderBrowserDialog fbd = new ()
            {
                Description = "Wybierz folder paczki (zawierający Zestawienie_DXF.xlsx):",
                SelectedPath = defaultPath,
                ShowNewFolderButton = false
            };

            if (fbd.ShowDialog() == DialogResult.OK)
            {
                return fbd.SelectedPath;
            }
            return null;
        }

        public static string GetSummaryExcelPath(string selectedPath)
        {
           string[] excelFiles = Directory.GetFiles(selectedPath, "*.xlsx");
           if (excelFiles.Length == 0 || excelFiles.Length > 1) return null;
           return excelFiles[0];
        }

        public static void ProcessDrawings(string defaultPath, string targetPackageDir, string excelPath)
        {
            string drawingsDir = Path.Combine(targetPackageDir, "Rysunki");
            if (!Directory.Exists(drawingsDir)) Directory.CreateDirectory(drawingsDir);

            var allPdfFiles = Directory.GetFiles(defaultPath, "*.pdf", SearchOption.TopDirectoryOnly);
            var allDxfFiles = Directory.GetFiles(defaultPath, "*.dxf", SearchOption.TopDirectoryOnly);

            var pdfNames = allPdfFiles.Select(f => new { Path = f, Name = Path.GetFileNameWithoutExtension(f) }).ToList();
            var dxfNames = allDxfFiles.Select(f => new { Path = f, Name = Path.GetFileNameWithoutExtension(f) }).ToList();

            ExcelApp excelApp = null;
            ExcelWorkbooks workbooks = null;
            ExcelWorkbook workbook = null;
            ExcelSheets sheets = null;
            ExcelWorksheet worksheet = null;
            ExcelRange usedRange = null;
            ExcelRange expandedRange = null;

            try
            {
                excelApp = new ExcelApp
                {
                    Visible = false,
                    DisplayAlerts = false,
                    ScreenUpdating = false,
                    EnableEvents = false
                };

                workbooks = excelApp.Workbooks;
                workbook = workbooks.Open(excelPath, ReadOnly: false);
                sheets = workbook.Sheets;
                worksheet = (ExcelWorksheet)sheets[1];
                usedRange = worksheet.UsedRange;

                int nameColIndex = ExcelWrapper.GetColumnNumber(usedRange, "Nr części");
                if (nameColIndex == 0)
                {
                    MessageBox.Show("Nie znaleziono kolumny 'Nr części' w pliku Excel.", "Błąd");
                    return;
                }

                int rowCount = usedRange.Rows.Count;
                int colCount = usedRange.Columns.Count;

                // Sprawdzamy, czy kolumna "Rysunki" już czasem nie istnieje (gdybyśmy odpalali makro 2 raz)
                int drawingsColIndex = ExcelWrapper.GetColumnNumber(usedRange, "Rysunki");
                bool isNewColumn = false;

                if (drawingsColIndex == 0)
                {
                    drawingsColIndex = colCount + 1; // Tworzymy nową na końcu tabeli
                    isNewColumn = true;
                }

                // Powiększamy zakres czytania Excela o 1 nową kolumnę
                expandedRange = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[rowCount, drawingsColIndex]];
                object[,] data = (object[,])expandedRange.Value2; // Wrzucamy do RAM

                data[1, drawingsColIndex] = "Rysunki"; // Ustawiamy nagłówek

                // Pętla od wiersza nr 2 (omijamy nagłówki)
                for (int i = 2; i <= rowCount; i++)
                {
                    bool wasDrawingCopied = false;

                    if (data[i, nameColIndex] != null)
                    {
                        string partName = data[i, nameColIndex].ToString().Trim();
                        if (!string.IsNullOrEmpty(partName))
                        {
                            // Szukamy i kopiujemy PDF (niezależnie od wielkości liter)
                            var matchingPdfs = pdfNames.Where(f => f.Name.Equals(partName, StringComparison.OrdinalIgnoreCase)).ToList();
                            foreach (var pdf in matchingPdfs)
                            {
                                string destFile = Path.Combine(drawingsDir, Path.GetFileName(pdf.Path));
                                File.Copy(pdf.Path, destFile, true);
                                wasDrawingCopied = true;
                            }

                            // Szukamy i kopiujemy DXF
                            var matchingDxfs = dxfNames.Where(f => f.Name.Equals(partName, StringComparison.OrdinalIgnoreCase)).ToList();
                            foreach (var dxf in matchingDxfs)
                            {
                                string destFile = Path.Combine(drawingsDir, Path.GetFileName(dxf.Path));
                                File.Copy(dxf.Path, destFile, true);
                                wasDrawingCopied = true;
                            }
                        }
                    }

                    // Uzupełniamy Excela po skopiowaniu
                    if (wasDrawingCopied)
                        data[i, drawingsColIndex] = "OK";
                    else
                        data[i, drawingsColIndex] = "-";
                }

                // Wpisujemy zaktualizowaną tablicę z RAM jednym ruchem z powrotem do Excela
                expandedRange.Value2 = data;

                // Formatowanie wizualne (tylko jeśli dodaliśmy kolumnę jako nową)
                if (isNewColumn)
                {
                    ExcelRange headerCell = null;
                    try
                    {
                        headerCell = (ExcelRange)worksheet.Cells[1, drawingsColIndex];
                        headerCell.Font.Bold = true;
                        headerCell.Interior.Color = ColorTranslator.ToOle(Color.LightGray);

                        expandedRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        expandedRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        worksheet.UsedRange.Columns.AutoFit();
                    }
                    finally { CoreUtils.ReleaseCom(ref headerCell); }
                }

                workbook.Save(); // Zapisujemy dokument
            }
            finally
            {
                workbook?.Close(false);
                excelApp?.Quit();

                CoreUtils.ReleaseCom(ref expandedRange);
                CoreUtils.ReleaseCom(ref usedRange);
                CoreUtils.ReleaseCom(ref worksheet);
                CoreUtils.ReleaseCom(ref sheets);
                CoreUtils.ReleaseCom(ref workbook);
                CoreUtils.ReleaseCom(ref workbooks);
                CoreUtils.ReleaseCom(ref excelApp);

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}