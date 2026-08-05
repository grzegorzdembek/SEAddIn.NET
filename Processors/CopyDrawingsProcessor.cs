using SolidEdgeAdd_In.Utils;

namespace SolidEdgeAdd_In.Processors
{
    public class CopyDrawingsProcessor
    {
        private readonly SeDocument _document;
        private List<(string FileName, string Path)> _pdfFiles;
        private List<(string FileName, string Path)> _dxfFiles;

        private string _projectDirectory;
        private string _packagesDirectory;
        private string _excelFilePath;
        private string _drawingsDirectory;

        public CopyDrawingsProcessor(SeDocument document)
        {
            _document = document;
            _pdfFiles = new List<(string FileName, string Path)>();
            _dxfFiles = new List<(string FileName, string Path)>();
        }

        public bool Initialize()
        {
            if (_document == null || string.IsNullOrEmpty(_document.FullName)) { return false; }

            _projectDirectory = Path.GetDirectoryName(_document.FullName);
            string expectedPackagesDirectory = Path.Combine(_projectDirectory, Constants.Folders.Packages);

            if (!Directory.Exists(expectedPackagesDirectory))
            {
                MessageBox.Show("Packages folder not found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            // DIALOG
            FolderBrowserDialog fbd = new()
            {
                Description = "Select the folder where you want to add the Drawings folder (containing the sheet metal summary):",
                SelectedPath = expectedPackagesDirectory,
                ShowNewFolderButton = false
            };

            if (fbd.ShowDialog() == DialogResult.OK) { _packagesDirectory = fbd.SelectedPath; }
            else { return false; }

            // VALIDATION
            string[] excelFiles = Directory.GetFiles(_packagesDirectory, "*.xlsx");
            if (excelFiles.Length == 0 || excelFiles.Length > 1)
            {
                MessageBox.Show("Missing or multiple sheet metal summaries found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            _excelFilePath = excelFiles[0];
            _drawingsDirectory = Path.Combine(_packagesDirectory, Constants.Folders.Drawings);

            _pdfFiles = Directory.GetFiles(_projectDirectory, "*.pdf", SearchOption.TopDirectoryOnly)
                                 .Select(f => (Path.GetFileNameWithoutExtension(f), f)).ToList();

            _dxfFiles = Directory.GetFiles(_projectDirectory, "*.dxf", SearchOption.TopDirectoryOnly)
                                 .Select(f => (Path.GetFileNameWithoutExtension(f), f)).ToList();

            return true;
        }

        public void Process()
        {
            if (!Directory.Exists(_drawingsDirectory)) { Directory.CreateDirectory(_drawingsDirectory); }

            ExcelApp excelApp = null;
            ExcelWorkbooks workbooks = null;
            ExcelWorkbook workbook = null;
            ExcelSheets sheets = null;
            ExcelWorksheet worksheet = null;

            ExcelRange usedRange = null;
            ExcelRange expandedRange = null;
            ExcelRange cells = null;
            ExcelRange columns = null;
            ExcelRange rows = null;

            ExcelRange headerCell = null;
            dynamic headerFont = null;
            dynamic headerInterior = null;
            dynamic expandedBorders = null;

            ExcelRange startCell = null;
            ExcelRange endCell = null;

            try
            {
                excelApp = new ExcelApp { Visible = false, DisplayAlerts = false, ScreenUpdating = false, EnableEvents = false };
                workbooks = excelApp.Workbooks;
                workbook = workbooks.Open(_excelFilePath, ReadOnly: false);
                sheets = workbook.Sheets;
                worksheet = (ExcelWorksheet)sheets[1];

                usedRange = worksheet.UsedRange;
                cells = worksheet.Cells;
                rows = usedRange.Rows;
                columns = usedRange.Columns;

                // COLUMNS SETUP
                int fileNameColumnIndex = ExcelUtils.GetColumnNumber(usedRange, Constants.ExcelHeaders.FileName);
                if (fileNameColumnIndex == 0)
                {
                    MessageBox.Show("Column not found in the Excel file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                int rowCount = rows.Count;
                int colCount = columns.Count;
                int drawingsColumnIndex = ExcelUtils.GetColumnNumber(usedRange, Constants.ExcelHeaders.Drawings);
                bool isNewColumn = false;

                if (drawingsColumnIndex == 0)
                {
                    drawingsColumnIndex = colCount + 1;
                    isNewColumn = true;
                }

                try
                {
                    startCell = (ExcelRange)cells[1, 1];
                    endCell = (ExcelRange)cells[rowCount, drawingsColumnIndex];
                    expandedRange = worksheet.Range[startCell, endCell];
                }
                finally { Helpers.ReleaseCom(ref startCell); Helpers.ReleaseCom(ref endCell); }

                // DATA PROCESSING
                object[,] data = (object[,])expandedRange.Value2;
                data[1, drawingsColumnIndex] = Constants.ExcelHeaders.Drawings;

                for (int i = 2; i <= rowCount; i++)
                {
                    bool isPdfReady = false;
                    bool isDxfReady = false;

                    if (data[i, fileNameColumnIndex] != null)
                    {
                        string fileName = data[i, fileNameColumnIndex].ToString().Trim();

                        if (!string.IsNullOrEmpty(fileName))
                        {
                            string expectedPdfPath = Path.Combine(_drawingsDirectory, fileName + ".pdf");
                            string expectedDxfPath = Path.Combine(_drawingsDirectory, fileName + ".dxf");

                            // PDF
                            if (File.Exists(expectedPdfPath)) { isPdfReady = true; }
                            else
                            {
                                var matchingPdfs = _pdfFiles.Where(f => f.FileName.Equals(fileName, StringComparison.OrdinalIgnoreCase)).ToList();
                                foreach (var pdf in matchingPdfs) { File.Copy(pdf.Path, expectedPdfPath, true); isPdfReady = true; }
                            }

                            // DXF
                            if (File.Exists(expectedDxfPath)) { isDxfReady = true; }
                            else
                            {
                                var matchingDxfs = _dxfFiles.Where(f => f.FileName.Equals(fileName, StringComparison.OrdinalIgnoreCase)).ToList();
                                foreach (var dxf in matchingDxfs) { File.Copy(dxf.Path, expectedDxfPath, true); isDxfReady = true; }
                            }
                        }
                    }

                    if (isDxfReady && isPdfReady) { data[i, drawingsColumnIndex] = "Skopiowano"; }
                    else { data[i, drawingsColumnIndex] = "-"; }
                }

                expandedRange.Value2 = data;

                // FORMATTING
                if (isNewColumn)
                {
                    headerCell = (ExcelRange)cells[1, drawingsColumnIndex];
                    headerFont = headerCell.Font;
                    headerFont.Bold = true;
                    headerInterior = headerCell.Interior;
                    headerInterior.Color = ColorTranslator.ToOle(Color.LightGray);

                    expandedRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    expandedBorders = expandedRange.Borders;
                    expandedBorders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    columns.AutoFit();
                }

                workbook.Save();
            }
            finally
            {
                Helpers.ReleaseCom(ref expandedBorders); Helpers.ReleaseCom(ref headerInterior);
                Helpers.ReleaseCom(ref headerFont); Helpers.ReleaseCom(ref headerCell);
                Helpers.ReleaseCom(ref rows); Helpers.ReleaseCom(ref columns);
                Helpers.ReleaseCom(ref cells); Helpers.ReleaseCom(ref expandedRange);
                Helpers.ReleaseCom(ref usedRange); Helpers.ReleaseCom(ref worksheet);
                Helpers.ReleaseCom(ref sheets);

                if (workbook != null) { try { workbook.Close(false); } catch { } }
                Helpers.ReleaseCom(ref workbook); Helpers.ReleaseCom(ref workbooks);

                if (excelApp != null) { try { excelApp.Quit(); } catch { } }
                Helpers.ReleaseCom(ref excelApp);
            }
        }
    }
}