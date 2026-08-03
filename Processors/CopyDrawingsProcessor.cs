using SolidEdgeAdd_In.Utils;

namespace SolidEdgeAdd_In.Processors
{
    public class CopyDrawingsProcessor
    {
        private readonly SeDocument _document;
        private string _projectDirectory;
        private string _selectedDirectory;
        private string _excelPath;
        private string _drawingsDirectory;

        private List<string> _pdfFiles;
        private List<string> _dxfFiles;

        public CopyDrawingsProcessor(SeDocument document)
        {
            _document = document;
            _pdfFiles = new List<string>();
            _dxfFiles = new List<string>();
        }

        public bool Initialize()
        {
            if (_document == null || string.IsNullOrEmpty(_document.FullName))
            {
                return false;
            }

            _projectDirectory = Path.GetDirectoryName(_document.FullName);

            string packagesDirectory = Path.Combine(_projectDirectory, Constants.Folders.Packages);

            if (!Directory.Exists(packagesDirectory))
            {
                MessageBox.Show("Packages folder not found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            using (FolderBrowserDialog fbd = new FolderBrowserDialog
            {
                Description = "Select the folder where you want to add the Drawings folder (containing the sheet metal summary):",
                SelectedPath = packagesDirectory,
                ShowNewFolderButton = false
            })
            {
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    _selectedDirectory = fbd.SelectedPath;
                }
                else
                {
                    return false;
                }
            }

            string[] excelFiles = Directory.GetFiles(_selectedDirectory, "*.xlsx");

            if (excelFiles.Length == 0 || excelFiles.Length > 1)
            {
                MessageBox.Show("Missing or multiple sheet metal summaries found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            _excelPath = excelFiles[0];
            _drawingsDirectory = Path.Combine(_selectedDirectory, Constants.Folders.Drawings);

            _pdfFiles = Directory.GetFiles(_projectDirectory, "*.pdf", SearchOption.TopDirectoryOnly).ToList();
            _dxfFiles = Directory.GetFiles(_projectDirectory, "*.dxf", SearchOption.TopDirectoryOnly).ToList();

            return true;
        }

        public void Process()
        {
            if (!Directory.Exists(_drawingsDirectory))
            {
                Directory.CreateDirectory(_drawingsDirectory);
            }

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
                excelApp = new ExcelApp
                {
                    Visible = false,
                    DisplayAlerts = false,
                    ScreenUpdating = false,
                    EnableEvents = false
                };

                workbooks = excelApp.Workbooks;
                workbook = workbooks.Open(_excelPath, ReadOnly: false);
                sheets = workbook.Sheets;
                worksheet = (ExcelWorksheet)sheets[1];

                usedRange = worksheet.UsedRange;
                cells = worksheet.Cells;
                rows = usedRange.Rows;
                columns = usedRange.Columns;

                int nameColIndex = ExcelWrapper.GetColumnNumber(usedRange, Constants.ExcelHeaders.PartNumber);

                if (nameColIndex == 0)
                {
                    MessageBox.Show("Column not found in the Excel file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                int rowCount = rows.Count;
                int colCount = columns.Count;
                int drawingsColIndex = ExcelWrapper.GetColumnNumber(usedRange, Constants.ExcelHeaders.Drawings);
                bool isNewColumn = false;

                if (drawingsColIndex == 0)
                {
                    drawingsColIndex = colCount + 1;
                    isNewColumn = true;
                }

                try
                {
                    startCell = (ExcelRange)cells[1, 1];
                    endCell = (ExcelRange)cells[rowCount, drawingsColIndex];
                    expandedRange = worksheet.Range[startCell, endCell];
                }
                finally
                {
                    CoreUtils.ReleaseCom(ref startCell);
                    CoreUtils.ReleaseCom(ref endCell);
                }

                object[,] data = (object[,])expandedRange.Value2;
                data[1, drawingsColIndex] = Constants.ExcelHeaders.Drawings;

                for (int i = 2; i <= rowCount; i++)
                {
                    bool isPdfReady = false;
                    bool isDxfReady = false;

                    if (data[i, nameColIndex] != null)
                    {
                        string partName = data[i, nameColIndex].ToString().Trim();

                        if (!string.IsNullOrEmpty(partName))
                        {
                            string expectedPdfPath = Path.Combine(_drawingsDirectory, partName + ".pdf");
                            string expectedDxfPath = Path.Combine(_drawingsDirectory, partName + ".dxf");

                            if (File.Exists(expectedPdfPath))
                            {
                                isPdfReady = true;
                            }
                            else
                            {
                                var matchingPdfs = _pdfFiles.Where(f => Path.GetFileNameWithoutExtension(f).Equals(partName, StringComparison.OrdinalIgnoreCase)).ToList();
                                foreach (var pdf in matchingPdfs)
                                {
                                    File.Copy(pdf, expectedPdfPath, true);
                                    isPdfReady = true;
                                }
                            }

                            if (File.Exists(expectedDxfPath))
                            {
                                isDxfReady = true;
                            }
                            else
                            {
                                var matchingDxfs = _dxfFiles.Where(f => Path.GetFileNameWithoutExtension(f).Equals(partName, StringComparison.OrdinalIgnoreCase)).ToList();
                                foreach (var dxf in matchingDxfs)
                                {
                                    File.Copy(dxf, expectedDxfPath, true);
                                    isDxfReady = true;
                                }
                            }
                        }
                    }

                    if (isDxfReady && isPdfReady)
                    {
                        data[i, drawingsColIndex] = "Copied";
                    }
                    else
                    {
                        data[i, drawingsColIndex] = "-";
                    }
                }

                expandedRange.Value2 = data;

                if (isNewColumn)
                {
                    headerCell = (ExcelRange)cells[1, drawingsColIndex];
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
                CoreUtils.ReleaseCom(ref expandedBorders);
                CoreUtils.ReleaseCom(ref headerInterior);
                CoreUtils.ReleaseCom(ref headerFont);
                CoreUtils.ReleaseCom(ref headerCell);
                CoreUtils.ReleaseCom(ref rows);
                CoreUtils.ReleaseCom(ref columns);
                CoreUtils.ReleaseCom(ref cells);
                CoreUtils.ReleaseCom(ref expandedRange);
                CoreUtils.ReleaseCom(ref usedRange);
                CoreUtils.ReleaseCom(ref worksheet);
                CoreUtils.ReleaseCom(ref sheets);

                if (workbook != null)
                {
                    try
                    {
                        workbook.Close(false);
                    }
                    catch
                    {
                    }
                }

                CoreUtils.ReleaseCom(ref workbook);
                CoreUtils.ReleaseCom(ref workbooks);

                if (excelApp != null)
                {
                    try
                    {
                        excelApp.Quit();
                    }
                    catch
                    {
                    }
                }
                CoreUtils.ReleaseCom(ref excelApp);
            }
        }
    }
}