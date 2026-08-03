using SolidEdgeAdd_In.Utils;

namespace SolidEdgeAdd_In.Processors
{
    public class ExportDxfsProcessor
    {
        private readonly SeAssembly _assembly;
        private readonly SeApp _application;
        private readonly Logger _logger;

        private string _assemblyFullName;
        private string _assemblyDirectory;

        private int _multiplier;
        private string _subDirectory;
        private Dictionary<string, FileData> _occurrencesData;

        public ExportDxfsProcessor(SeAssembly assembly)
        {
            _assembly = assembly;
            _application = assembly.Application;
            _logger = new Logger();
            _occurrencesData = new Dictionary<string, FileData>(StringComparer.OrdinalIgnoreCase);
        }

        public bool Initialize()
        {
            _assemblyFullName = _assembly.FullName;
            _assemblyDirectory = Path.GetDirectoryName(_assemblyFullName);

            if (!TryGetMultiplier())
            {
                return false;
            }

            SetSubDirectory();
            LoadOccurrencesData();

            if (_occurrencesData.Count == 0)
            {
                MessageBox.Show("No sheet metal parts found to process.", "Stop", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }

            return true;
        }

        public void Process()
        {
            SeDocument document = null;
            SeModels models = null;
            SeModel model = null;
            SeFlanges flanges = null;
            SeFlatPatternModels flatPatterns = null;

            ExcelApp excelApp = null;
            ExcelWorkbooks workbooks = null;
            ExcelWorkbook workbook = null;
            ExcelSheets xlSheets = null;
            ExcelWorksheet worksheet = null;
            ExcelRange headerRange = null;
            ExcelRange usedRange = null;
            ExcelRange columns = null;
            ExcelRange cells = null;
            ExcelRange startHeaderCell = null;
            ExcelRange endHeaderCell = null;

            dynamic headerFont = null;
            dynamic headerInterior = null;
            dynamic headerBorders = null;
            dynamic usedBorders = null;

            string mainDirectory = _assemblyDirectory;

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
                workbook = workbooks.Add();
                excelApp.Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationManual;
                xlSheets = workbook.Sheets;
                worksheet = (ExcelWorksheet)xlSheets[1];
                cells = worksheet.Cells;

                object[,] headerData = new object[1, 12]
                {
                    {
                        Constants.ExcelHeaders.Lp,
                        Constants.ExcelHeaders.PartNumber,
                        Constants.ExcelHeaders.Quantity,
                        Constants.ExcelHeaders.Name,
                        Constants.ExcelHeaders.Thickness,
                        Constants.ExcelHeaders.Width,
                        Constants.ExcelHeaders.Length,
                        Constants.ExcelHeaders.Color,
                        Constants.ExcelHeaders.Finish,
                        Constants.ExcelHeaders.Flanges,
                        Constants.ExcelHeaders.Material,
                        Constants.ExcelHeaders.Date
                    }
                };

                try
                {
                    startHeaderCell = (ExcelRange)cells[1, 1];
                    endHeaderCell = (ExcelRange)cells[1, 12];
                    headerRange = worksheet.Range[startHeaderCell, endHeaderCell];
                    headerRange.Value = headerData;
                }
                finally
                {
                    CoreUtils.ReleaseCom(ref startHeaderCell);
                    CoreUtils.ReleaseCom(ref endHeaderCell);
                }

                headerFont = headerRange.Font;
                headerFont.Bold = true;
                headerInterior = headerRange.Interior;
                headerInterior.Color = ColorTranslator.ToOle(Color.LightGray);
                headerBorders = headerRange.Borders;
                headerBorders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                object[,] excelData = new object[_occurrencesData.Count, 12];
                int dataRowIndex = 0;

                var sortedData = _occurrencesData.OrderBy(item => item.Value.Material)
                     .ThenBy(item => double.TryParse(item.Value.Thickness?.Replace('_', ','), out double t) ? t : 0)
                     .ThenBy(item => item.Value.Name)
                     .ToList();

                int i = 1;
                foreach (var item in sortedData)
                {
                    string name = item.Value.Name;
                    int count = item.Value.OccurrenceCount * _multiplier;
                    string material = item.Value.Material;
                    string thickness = item.Value.Thickness;

                    string mainDxFilefName = $"{thickness}mm_{material}_{name}.dxf";
                    string subDxfFileName = $"{thickness}mm_{count}pcs_{material}_{name}.dxf";

                    string mainDxfPath = Path.Combine(mainDirectory, mainDxFilefName);
                    string subDxfPath = Path.Combine(_subDirectory, subDxfFileName);

                    object finalSizeX = "-";
                    string sizeX = item.Value.SizeX;

                    if (!string.IsNullOrEmpty(sizeX))
                    {
                        sizeX = sizeX.ToLower().Replace("mm", "").Replace(" ", "").Replace(".", ",").Trim();

                        if (double.TryParse(sizeX, out double parsedX))
                        {
                            finalSizeX = parsedX;
                        }
                        else
                        {
                            finalSizeX = sizeX;
                        }
                    }

                    object finalSizeY = "-";
                    string sizeY = item.Value.SizeY;

                    if (!string.IsNullOrEmpty(sizeY))
                    {
                        sizeY = sizeY.ToLower().Replace("mm", "").Replace(" ", "").Replace(".", ",").Trim();

                        if (double.TryParse(sizeY, out double parsedY))
                        {
                            finalSizeY = parsedY;
                        }
                        else
                        {
                            finalSizeY = sizeY;
                        }
                    }

                    string color = string.IsNullOrEmpty(item.Value.Color) ? "-" : item.Value.Color;
                    string finish = string.IsNullOrEmpty(item.Value.Finish) ? "-" : item.Value.Finish;
                    string title = string.IsNullOrEmpty(item.Value.Title) ? "-" : item.Value.Title;

                    string dxfDate = item.Value.DxfDate;
                    bool needGenerationDxf = string.IsNullOrEmpty(dxfDate);

                    bool isOpen = false;
                    string flangesValue = "-";

                    try
                    {
                        document = CoreUtils.GetOpenDocument(_application, item.Key);
                        isOpen = true;

                        if (document is SePart part)
                        {
                            models = part.Models;
                            flatPatterns = part.FlatPatternModels;
                        }
                        else if (document is SeSheetMetal sheetMetal)
                        {
                            models = sheetMetal.Models;
                            flatPatterns = sheetMetal.FlatPatternModels;
                        }

                        if (flatPatterns == null || models == null || flatPatterns.Count == 0 || models.Count == 0)
                        {
                            _logger.LogSkip(name, "Missing flat pattern.");
                            continue;
                        }

                        model = models.Item(1);
                        flanges = model.Flanges;

                        if (flanges.Count != 0)
                        {
                            flangesValue = flanges.Count.ToString();
                        }

                        if (needGenerationDxf)
                        {
                            if (File.Exists(subDxfPath))
                            {
                                try { File.Delete(subDxfPath); } catch { }
                            }

                            using var properties = new PropertyProvider(document);
                            properties.UpdateDxfDate();
                            dxfDate = properties.DxfDate;

                            models.SaveAsFlatDXFEx(subDxfPath, null, null, null, true);
                            _logger.LogSuccess($"{name} - Created new flat pattern Dxf.");
                        }
                        else
                        {
                            _logger.LogSuccess($"{name} - File has Dxf property.");
                        }

                        if (File.Exists(subDxfPath))
                        {
                            File.Copy(subDxfPath, mainDxfPath, true);

                            excelData[dataRowIndex, 0] = i++;
                            excelData[dataRowIndex, 1] = name;
                            excelData[dataRowIndex, 2] = count;
                            excelData[dataRowIndex, 3] = title;
                            excelData[dataRowIndex, 4] = $"{thickness} mm";
                            excelData[dataRowIndex, 5] = finalSizeX;
                            excelData[dataRowIndex, 6] = finalSizeY;
                            excelData[dataRowIndex, 7] = color;
                            excelData[dataRowIndex, 8] = finish;
                            excelData[dataRowIndex, 9] = flangesValue;
                            excelData[dataRowIndex, 10] = material;
                            excelData[dataRowIndex, 11] = dxfDate;

                            dataRowIndex++;
                        }
                        else
                        {
                            _logger.LogError(name, "File has Dxf property, but physical file not found.");
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(name, $"Error: {ex.Message}");
                        continue;
                    }
                    finally
                    {
                        CoreUtils.ReleaseCom(ref flanges);
                        CoreUtils.ReleaseCom(ref model);
                        CoreUtils.ReleaseCom(ref flatPatterns);
                        CoreUtils.ReleaseCom(ref models);

                        if (isOpen && document != null)
                        {
                            try
                            {
                                document.Save();
                                document.Close(true);
                            }
                            catch { }
                        }

                        CoreUtils.ReleaseCom(ref document);
                    }
                }

                if (dataRowIndex > 0)
                {
                    ExcelRange startCell = null;
                    ExcelRange endCell = null;
                    ExcelRange writeRange = null;

                    try
                    {
                        startCell = (ExcelRange)cells[2, 1];
                        endCell = (ExcelRange)cells[dataRowIndex + 1, 12];
                        writeRange = worksheet.Range[startCell, endCell];
                        writeRange.Value = excelData;
                    }
                    finally
                    {
                        CoreUtils.ReleaseCom(ref writeRange);
                        CoreUtils.ReleaseCom(ref endCell);
                        CoreUtils.ReleaseCom(ref startCell);
                    }
                }

                usedRange = worksheet.UsedRange;
                columns = usedRange.Columns;
                columns.AutoFit();
                usedRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                usedBorders = usedRange.Borders;
                usedBorders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                string excelPath = Path.Combine(_subDirectory, "DXF_Summary.xlsx");

                if (File.Exists(excelPath))
                {
                    File.Delete(excelPath);
                }

                workbook.SaveAs(excelPath);
            }
            finally
            {
                CoreUtils.ReleaseCom(ref usedBorders);
                CoreUtils.ReleaseCom(ref headerBorders);
                CoreUtils.ReleaseCom(ref headerInterior);
                CoreUtils.ReleaseCom(ref headerFont);
                CoreUtils.ReleaseCom(ref cells);
                CoreUtils.ReleaseCom(ref columns);
                CoreUtils.ReleaseCom(ref usedRange);
                CoreUtils.ReleaseCom(ref headerRange);
                CoreUtils.ReleaseCom(ref worksheet);
                CoreUtils.ReleaseCom(ref xlSheets);

                if (workbook != null)
                {
                    try { workbook.Close(false); } catch { }
                }

                CoreUtils.ReleaseCom(ref workbook);
                CoreUtils.ReleaseCom(ref workbooks);

                if (excelApp != null)
                {
                    try { excelApp.Quit(); } catch { }
                }

                CoreUtils.ReleaseCom(ref excelApp);
            }
        }

        private bool TryGetMultiplier()
        {
            SeDocument document = (SeDocument)_assembly;

            using var properties = new PropertyProvider(document);
            int count = properties.Count;

            if (count == 0)
            {
                var (isConfirmed, multiplier) = DialogService.GetMultiplier();

                if (isConfirmed)
                {
                    properties.Count = multiplier;
                    _multiplier = multiplier;
                    return true;
                }

                return false;
            }

            _multiplier = count;
            return true;
        }

        private void SetSubDirectory()
        {
            string fileName = Path.GetFileNameWithoutExtension(_assemblyFullName);
            string number = fileName.Length >= 4 ? fileName.Substring(0, 4) : fileName;
            string date = DateTime.Now.ToString("yyyy-MM-dd");

            string packagesDirectory = Path.Combine(_assemblyDirectory, Constants.Folders.Packages);

            if (!Directory.Exists(packagesDirectory))
            {
                Directory.CreateDirectory(packagesDirectory);
            }

            _subDirectory = Path.Combine(packagesDirectory, $"{number}_{date}");

            if (!Directory.Exists(_subDirectory))
            {
                Directory.CreateDirectory(_subDirectory);
            }
        }

        private void LoadOccurrencesData()
        {
            SeOccurrences occurrences = null;

            try
            {
                occurrences = _assembly.Occurrences;
                AssemblyTreeWalker.BuildDataForExportDxfs(occurrences, _occurrencesData, _logger);
            }
            finally
            {
                CoreUtils.ReleaseCom(ref occurrences);
            }
        }
    }
}