using SolidEdgeAdd_In.Utils;

namespace SolidEdgeAdd_In.Processors
{
    public class ExportDxfsProcessor
    {
        private readonly SeAssembly _assembly;
        private readonly SeApp _application;

        private string _assemblyPath;
        private string _assemblyName;
        private string _projectDirectory;

        private readonly Dictionary<string, FileData> _data;
        private int _dataCount;

        private int _multiplier;
        private string _targetDirectory;

        private readonly Logger _logger;

        public ExportDxfsProcessor(SeAssembly assembly)
        {
            _assembly = assembly;
            _application = _assembly.Application;

            _data = new Dictionary<string, FileData>(StringComparer.OrdinalIgnoreCase);
            _logger = new Logger();
        }

        public bool Initialize()
        {
            _assemblyPath = _assembly.FullName;
            _projectDirectory = Path.GetDirectoryName(_assemblyPath);
            _assemblyName = Path.GetFileNameWithoutExtension(_assemblyPath);

            if (!IsLoaded_Data())
            {
                return false;
            }

            if (!IsLoaded_Multiplier())
            {
                return false;
            }

            if (!IsLoaded_TargetDirectory())
            {
                return false;
            }

            return true;
        }

        public void Process()
        {
            SeDocument document = null;
            SeModels models = null;
            SeModel model = null;
            SeFlatPatternModels flatPatterns = null;
            SeFlanges flanges = null;
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
                        Constants.ExcelHeaders.FileName,
                        Constants.ExcelHeaders.Count,
                        Constants.ExcelHeaders.Title,
                        Constants.ExcelHeaders.Thickness,
                        Constants.ExcelHeaders.SizeX,
                        Constants.ExcelHeaders.SizeY,
                        Constants.ExcelHeaders.Color,
                        Constants.ExcelHeaders.Finish,
                        Constants.ExcelHeaders.Flanges,
                        Constants.ExcelHeaders.Material,
                        Constants.ExcelHeaders.DxfDate
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
                    Helpers.ReleaseCom(ref startHeaderCell);
                    Helpers.ReleaseCom(ref endHeaderCell);
                }

                headerFont = headerRange.Font;
                headerFont.Bold = true;
                headerInterior = headerRange.Interior;
                headerInterior.Color = ColorTranslator.ToOle(Color.LightGray);
                headerBorders = headerRange.Borders;
                headerBorders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                object[,] excelData = new object[_data.Count, 12];
                int dataRowIndex = 0;

                var sortedData = _data.OrderBy(item => item.Value.Material)
                     .ThenBy(item => double.TryParse(item.Value.Thickness?.Replace('_', ','), out double t) ? t : 0)
                     .ThenBy(item => item.Value.Name).ToList();

                int lp = 1;

                foreach (var item in sortedData)
                {
                    string fileName = item.Value.Name;
                    int count = item.Value.OccurrenceCount * _multiplier;
                    string material = item.Value.Material;
                    string thickness = item.Value.Thickness;

                    string mainDxfFileName = $"{thickness}mm_{material}_{fileName}.dxf";
                    string subDxfFileName = $"{thickness}mm_{count}szt_{material}_{fileName}.dxf";

                    string mainDxfPath = Path.Combine(_projectDirectory, mainDxfFileName);
                    string subDxfPath = Path.Combine(_targetDirectory, subDxfFileName);

                    object finalSizeX = "-";
                    string sizeX = item.Value.SizeX;

                    if (!string.IsNullOrEmpty(sizeX))
                    {
                        sizeX = sizeX.ToLower().Replace("mm", "").Replace(" ", "").Replace(",", ".").Trim();

                        if (double.TryParse(sizeX, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double parsedX))
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
                        sizeY = sizeY.ToLower().Replace("mm", "").Replace(" ", "").Replace(",", ".").Trim();

                        if (double.TryParse(sizeY, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double parsedY))
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
                    string flangesValue = "-";

                    string dxfDate = item.Value.DxfDate;
                    bool needGenerationDxf = string.IsNullOrEmpty(dxfDate);
                    bool isOpen = false;

                    try
                    {
                        document = Helpers.GetOpenDocument(_application, item.Key);
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
                            _logger.LogSkip(fileName, "Missing flat pattern.");
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
                                File.Delete(subDxfPath);
                            }

                            using PropertyUtils properties = new (document);
                            properties.UpdateDxfDate();
                            dxfDate = properties.DxfDate;
                            models.SaveAsFlatDXFEx(subDxfPath, null, null, null, true);
                            document.Save();
                            _logger.LogSuccess($"{fileName} - Created new flat pattern Dxf.");
                        }
                        else
                        {
                            _logger.LogSuccess($"{fileName} - File has Dxf property.");
                        }

                        if (File.Exists(subDxfPath))
                        {
                            File.Copy(subDxfPath, mainDxfPath, true);

                            excelData[dataRowIndex, 0] = lp++;
                            excelData[dataRowIndex, 1] = fileName;
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
                            _logger.LogError(fileName, "File has Dxf property, but file not found.");
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(fileName, $"Error: {ex.Message}");
                        continue;
                    }
                    finally
                    {
                        Helpers.ReleaseCom(ref flanges);
                        Helpers.ReleaseCom(ref model);
                        Helpers.ReleaseCom(ref flatPatterns);
                        Helpers.ReleaseCom(ref models);

                        if (isOpen)
                        {
                            try
                            {
                                document?.Close(false);
                            }
                            catch
                            {
                            }
                        }

                        Helpers.ReleaseCom(ref document);
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
                        Helpers.ReleaseCom(ref writeRange);
                        Helpers.ReleaseCom(ref endCell);
                        Helpers.ReleaseCom(ref startCell);
                    }
                }

                usedRange = worksheet.UsedRange;
                columns = usedRange.Columns;
                columns.AutoFit();
                usedRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                usedBorders = usedRange.Borders;
                usedBorders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                string excelFilePath = Path.Combine(_targetDirectory, "DXF_Summary.xlsx");

                if (File.Exists(excelFilePath))
                {
                    File.Delete(excelFilePath);
                }

                workbook.SaveAs(excelFilePath);
            }
            finally
            {
                _logger.SaveReport(_targetDirectory);
                Helpers.ReleaseCom(ref usedBorders);
                Helpers.ReleaseCom(ref headerBorders);
                Helpers.ReleaseCom(ref headerInterior);
                Helpers.ReleaseCom(ref headerFont);
                Helpers.ReleaseCom(ref cells);
                Helpers.ReleaseCom(ref columns);
                Helpers.ReleaseCom(ref usedRange);
                Helpers.ReleaseCom(ref headerRange);
                Helpers.ReleaseCom(ref worksheet);
                Helpers.ReleaseCom(ref xlSheets);

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

        private bool IsLoaded_Data()
        {
            SeOccurrences occurrences = null;

            try
            {
                occurrences = _assembly.Occurrences;
                DataUtils.BuildDataForExportDxfs(occurrences, _data, _logger);
            }
            finally
            {
                Helpers.ReleaseCom(ref occurrences);
            }

            _dataCount = _data.Count;

            if (!Helpers.IsMessageAccepted($"Liczba blach w złożeniu: {_dataCount}."))
            {
                return false;
            }

            return true;
        }

        private bool IsLoaded_Multiplier()
        {
            SeDocument document = (SeDocument)_assembly;
            using PropertyUtils properties = new PropertyUtils(document);
            int count = properties.Count;

            if (count == 0)
            {
                (bool isConfirmed, int multiplier) = DialogUtils.GetMultiplier();

                if (isConfirmed)
                {
                    properties.Count = multiplier;
                    _multiplier = multiplier;
                    return true;
                }

                return false;
            }

            _multiplier = count;

            if (!Helpers.IsMessageAccepted($"Przyjęto mnożnik: {_multiplier}."))
            {
                return false;
            }

            return true;
        }

        private bool IsLoaded_TargetDirectory()
        {
            string number = _assemblyName.Length >= 4 ? _assemblyName.Substring(0, 4) : _assemblyName;
            string date = DateTime.Now.ToString("yyyy-MM-dd");

            string packagesDirectory = Path.Combine(_projectDirectory, Constants.Folders.Packages);
            Directory.CreateDirectory(packagesDirectory);

            _targetDirectory = Path.Combine(packagesDirectory, $"{number}_{date}");
            Directory.CreateDirectory(_targetDirectory);

            return true;
        }
    }
}