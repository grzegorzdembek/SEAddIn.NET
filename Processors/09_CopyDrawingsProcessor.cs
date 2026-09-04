using SolidEdgeAdd_In.Utils;

namespace SolidEdgeAdd_In.Processors
{
    public class CopyDrawingsProcessor
    {
        private readonly SeDocument _document;

        private string _documentPath;
        private string _projectDirectory;

        private string _packagesDirectory;
        private string _selectedDirectory;
        private string _drawingsDirectory;

        private string _excelFilePath;

        private Dictionary<string, string> _pdfFiles;
        private Dictionary<string, string> _dxfFiles;

        public CopyDrawingsProcessor(SeDocument document)
        {
            _document = document;
        }

        public bool Initialize()
        {
            _documentPath = _document.FullName;
            _projectDirectory = Path.GetDirectoryName(_documentPath);

            if (!IsPackagesDirectory_Loaded())
            {
                return false;
            }

            if (!IsSelectedDirectory_Loaded())
            {
                return false;
            }

            if (!IsExcelFile_Loaded())
            {
                return false;
            }

            if (!IsDrawingsDirectory_Loaded())
            {
                return false;
            }

            if (!IsFiles_Loaded())
            {
                return false;
            }

            return true;
        }

        public void Process()
        {
            Directory.CreateDirectory(_drawingsDirectory);

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
                workbook = workbooks.Open(_excelFilePath, ReadOnly: false);
                sheets = workbook.Sheets;
                worksheet = (ExcelWorksheet)sheets[1];

                usedRange = worksheet.UsedRange;
                cells = worksheet.Cells;
                rows = usedRange.Rows;
                columns = usedRange.Columns;

                int fileNameColumnIndex = ExcelUtils.GetColumnIndex(usedRange, Constants.ExcelHeaders.FileName);

                if (fileNameColumnIndex == 0)
                {
                    MessageBox.Show("Nie znaleziono kolumny 'Numer części'.");
                    return;
                }

                int rowCount = rows.Count;
                int colCount = columns.Count;

                int drawingsColumnIndex = ExcelUtils.GetColumnIndex(usedRange, Constants.ExcelHeaders.Drawings);
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
                finally
                {
                    Helpers.ReleaseCom(ref startCell);
                    Helpers.ReleaseCom(ref endCell);
                }

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

                            if (File.Exists(expectedPdfPath))
                            {
                                isPdfReady = true;
                            }
                            else if (_pdfFiles.TryGetValue(fileName, out string pdfPath))
                            {
                                File.Copy(pdfPath, expectedPdfPath, true);
                                isPdfReady = true;
                            }

                            if (File.Exists(expectedDxfPath))
                            {
                                isDxfReady = true;
                            }
                            else if (_dxfFiles.TryGetValue(fileName, out string dxfPath))
                            {
                                File.Copy(dxfPath, expectedDxfPath, true);
                                isDxfReady = true;
                            }
                        }
                    }

                    if (isDxfReady && isPdfReady)
                    {
                        data[i, drawingsColumnIndex] = "Skopiowano";
                    }
                    else
                    {
                        data[i, drawingsColumnIndex] = "-";
                    }
                }

                expandedRange.Value2 = data;

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
                Helpers.ReleaseCom(ref expandedBorders);
                Helpers.ReleaseCom(ref headerInterior);
                Helpers.ReleaseCom(ref headerFont);
                Helpers.ReleaseCom(ref headerCell);
                Helpers.ReleaseCom(ref rows);
                Helpers.ReleaseCom(ref columns);
                Helpers.ReleaseCom(ref cells);
                Helpers.ReleaseCom(ref expandedRange);
                Helpers.ReleaseCom(ref usedRange);
                Helpers.ReleaseCom(ref worksheet);
                Helpers.ReleaseCom(ref sheets);

                try
                {
                    workbook?.Close(false);
                }
                catch
                {
                }

                Helpers.ReleaseCom(ref workbook);
                Helpers.ReleaseCom(ref workbooks);

                try
                {
                    excelApp?.Quit();
                }
                catch
                {
                }

                Helpers.ReleaseCom(ref excelApp);
            }
        }

        private bool IsPackagesDirectory_Loaded()
        {
            _packagesDirectory = Path.Combine(_projectDirectory, Constants.Folders.Packages);

            if (!Directory.Exists(_packagesDirectory))
            {
                MessageBox.Show("Folder 'Paczki' nie odnaleziony.");
                return false;
            }

            return true;
        }

        private bool IsSelectedDirectory_Loaded()
        {
            FolderBrowserDialog fbd = new()
            {
                Description = "Wybierz folder, do którego chcesz dodać folder Rysunki (zawierający podsumowający arkusza .xlsx):",
                SelectedPath = _packagesDirectory,
                ShowNewFolderButton = false
            };

            if (fbd.ShowDialog() == DialogResult.OK)
            {
                _selectedDirectory = fbd.SelectedPath;
                return true;
            }

            return false;
        }

        private bool IsExcelFile_Loaded()
        {
            string[] excelFiles = Directory.GetFiles(_selectedDirectory, "*.xlsx");

            if (excelFiles.Length == 0 || excelFiles.Length > 1)
            {
                MessageBox.Show("Brakuje pliku podsumowującego lub znaleziono więcej niż jeden.");
                return false;
            }

            _excelFilePath = excelFiles[0];

            return true;
        }

        private bool IsDrawingsDirectory_Loaded()
        {
            _drawingsDirectory = Path.Combine(_selectedDirectory, Constants.Folders.Drawings);

            return true;
        }

        private bool IsFiles_Loaded()
        {
            _pdfFiles = Directory.EnumerateFiles(_projectDirectory, "*.pdf", SearchOption.TopDirectoryOnly)
                .ToDictionary(f => Path.GetFileNameWithoutExtension(f), f => f, StringComparer.OrdinalIgnoreCase);

            _dxfFiles = Directory.EnumerateFiles(_projectDirectory, "*.dxf", SearchOption.TopDirectoryOnly)
                .ToDictionary(f => Path.GetFileNameWithoutExtension(f), f => f, StringComparer.OrdinalIgnoreCase);

            if (_pdfFiles.Count == 0 && _dxfFiles.Count == 0)
            {
                MessageBox.Show("Nie znaleziono plików PDF ani DXF w katalogu projektu.");
                return false;
            }

            return true;
        }
    }
}