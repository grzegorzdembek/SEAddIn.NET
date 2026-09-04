using SolidEdgeAdd_In.Utils;

namespace SolidEdgeAdd_In.Processors
{
    public class ExportOccurrencesListProcessor
    {
        private readonly SeAssembly _assembly;
        private readonly SeApp _application;

        private string _assemblyPath;
        private string _assemblyName;
        private string _projectDirectory;

        private readonly Dictionary<string, FileData> _data;
        private int _dataCount;

        private Dictionary<string, string> _thumbnails;
        private string _thumbnailsDirectory;
        private int _thumbnailsCount;
        private bool _isGenerateThumbnails;

        private int _multiplier;
        private string _targetDirectory;

        private List<string> _types;

        public ExportOccurrencesListProcessor(SeAssembly assembly)
        {
            _assembly = assembly;
            _application = _assembly.Application;

            _data = new Dictionary<string, FileData>();
            _thumbnails = new(StringComparer.OrdinalIgnoreCase);
        }

        public bool Initialize()
        {
            _assemblyPath = _assembly.FullName;
            _assemblyName = Path.GetFileNameWithoutExtension(_assemblyPath);
            _projectDirectory = Path.GetDirectoryName(_assemblyPath);

            if (!IsLoaded_Types())
            {
                return false;
            }

            if (!IsLoaded_Data())
            {
                return false;
            }

            if (!IsLoaded_Thumbnails())
            {
                return false;
            }

            if (!IsLoaded_GeneratingThumbnails())
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
            ExcelApp excelApp = null;
            ExcelWorkbooks workbooks = null;

            try
            {
                excelApp = new ExcelApp
                {
                    Visible = false,
                    DisplayAlerts = false,
                    AskToUpdateLinks = false,
                    EnableEvents = false
                };

                workbooks = excelApp.Workbooks;

                foreach (string type in _types)
                {
                    var filteredData = _data.Where(kvp => kvp.Value.Type == type).ToList();

                    if (filteredData.Count == 0)
                    {
                        continue;
                    }

                    ExcelWorkbook workbook = null;
                    ExcelSheets xlSheets = null;
                    ExcelWorksheet worksheet = null;
                    ExcelRange cells = null;
                    ExcelRange startHeaderCell = null;
                    ExcelRange endHeaderCell = null;
                    ExcelRange headerRange = null;
                    ExcelRange startDataCell = null;
                    ExcelRange endDataCell = null;
                    ExcelRange writeRange = null;
                    ExcelRange usedRange = null;
                    ExcelRange columns = null;

                    int lp = 1;

                    try
                    {
                        workbook = workbooks.Add();
                        excelApp.Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationManual;
                        xlSheets = workbook.Sheets;
                        worksheet = (ExcelWorksheet)xlSheets[1];
                        cells = worksheet.Cells;

                        object[,] headerData = new object[1, 6]
                        {
                            {
                                Constants.ExcelHeaders.Lp,
                                Constants.ExcelHeaders.FileName,
                                Constants.ExcelHeaders.Title,
                                Constants.ExcelHeaders.Type,
                                Constants.ExcelHeaders.Count,
                                Constants.ExcelHeaders.Thumbnail
                            }
                        };

                        try
                        {
                            startHeaderCell = (ExcelRange)cells[1, 1];
                            endHeaderCell = (ExcelRange)cells[1, 6];
                            headerRange = worksheet.Range[startHeaderCell, endHeaderCell];
                            headerRange.Value = headerData;
                        }
                        finally
                        {
                            Helpers.ReleaseCom(ref startHeaderCell);
                            Helpers.ReleaseCom(ref endHeaderCell);
                            Helpers.ReleaseCom(ref headerRange);
                        }

                        object[,] excelData = new object[filteredData.Count, 6];
                        int dataRowIndex = 0;

                        var sortedData = filteredData.OrderBy(item => item.Value.Name).ToList();

                        foreach (var item in sortedData)
                        {
                            excelData[dataRowIndex, 0] = lp++;
                            excelData[dataRowIndex, 1] = item.Value.Name;
                            excelData[dataRowIndex, 2] = item.Value.Title;
                            excelData[dataRowIndex, 3] = item.Value.Type;
                            excelData[dataRowIndex, 4] = item.Value.OccurrenceCount * _multiplier;
                            excelData[dataRowIndex, 5] = string.Empty;
                            dataRowIndex++;
                        }

                        try
                        {
                            startDataCell = (ExcelRange)cells[2, 1];
                            endDataCell = (ExcelRange)cells[filteredData.Count + 1, 6];
                            writeRange = worksheet.Range[startDataCell, endDataCell];
                            writeRange.Value = excelData;
                        }
                        finally
                        {
                            Helpers.ReleaseCom(ref startDataCell);
                            Helpers.ReleaseCom(ref endDataCell);
                            Helpers.ReleaseCom(ref writeRange);
                        }

                        ExcelUtils.Edit(worksheet);

                        usedRange = worksheet.UsedRange;
                        int fileNameColumnIndex = ExcelUtils.GetColumnIndex(usedRange, Constants.ExcelHeaders.FileName);
                        int thumbnailColumnIndex = ExcelUtils.GetColumnIndex(usedRange, Constants.ExcelHeaders.Thumbnail);

                        if (fileNameColumnIndex > 0 && thumbnailColumnIndex > 0)
                        {                     
                            ReportUtils.InsertThumbnailsOnly(worksheet, _thumbnails, fileNameColumnIndex, thumbnailColumnIndex);
                        }

                        string excelFilePath = Path.Combine(_targetDirectory, $"{_assemblyName}_{type}.xlsx");

                        if (File.Exists(excelFilePath))
                        {
                            File.Delete(excelFilePath);
                        }

                        workbook.SaveAs(excelFilePath);
                    }
                    finally
                    {
                        Helpers.ReleaseCom(ref columns);
                        Helpers.ReleaseCom(ref usedRange);
                        Helpers.ReleaseCom(ref cells);
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
                    }
                }
            }
            finally
            {
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

        private bool IsLoaded_Types()
        {
            List<string> selectedStrings = DialogUtils.GetSelectedTypes();

            if (selectedStrings == null || selectedStrings.Count == 0)
            {
                return false;
            }

            _types = new List<string>();

            foreach (var str in selectedStrings)
            {
                int start = str.IndexOf('(');
                int end = str.IndexOf(')');

                if (start >= 0 && end > start)
                {
                    _types.Add(str.Substring(start + 1, end - start - 1));
                }
            }

            return _types.Count > 0;
        }

        private bool IsLoaded_Data()
        {
            SeOccurrences occurrences = null;

            try
            {
                occurrences = _assembly.Occurrences;
                DataUtils.BuildDataForExportOccurrencesList(occurrences, _data, _types);
            }
            finally
            {
                Helpers.ReleaseCom(ref occurrences);
            }

            _dataCount = _data.Count;

            if (!Helpers.IsMessageAccepted($"Liczba plików w złożeniu: {_dataCount}."))
            {
                return false;
            }

            return true;
        }

        private bool IsLoaded_Thumbnails()
        {
            _thumbnailsDirectory = Path.Combine(_projectDirectory, Constants.Folders.Thumbnails);
            Directory.CreateDirectory(_thumbnailsDirectory);

            _thumbnails = Directory.EnumerateFiles(_thumbnailsDirectory, "*.jpg", SearchOption.TopDirectoryOnly).ToDictionary(
                                 file => Path.GetFileNameWithoutExtension(file),
                                 file => file,
                                 StringComparer.OrdinalIgnoreCase);

            _thumbnailsCount = _thumbnails.Count;

            if (!Helpers.IsMessageAccepted($"Liczba miniatur w folderze Miniatury: {_thumbnailsCount}."))
            {
                return false;
            }

            _isGenerateThumbnails = DialogUtils.IsGenerateThumbnails();

            return true;
        }

        private bool IsLoaded_GeneratingThumbnails()
        {
            if (!_isGenerateThumbnails)
            {
                return true;
            }

            foreach (var item in _data)
            {
                try
                {
                    string sourceFilePath = item.Key;
                    string thumbnailPath = Path.Combine(_thumbnailsDirectory, item.Value.Name + ".jpg");

                    if (File.Exists(thumbnailPath))
                    {
                        continue;
                    }

                    ReportUtils.ExtractAndSaveThumbnail(sourceFilePath, thumbnailPath, 256);

                    if (File.Exists(thumbnailPath))
                    {
                        if (!_thumbnails.ContainsKey(item.Value.Name))
                        {
                            _thumbnails.Add(item.Value.Name, thumbnailPath);
                        }
                    }
                }
                catch
                {
                    continue;
                }
            }

            _thumbnailsCount = _thumbnails.Count;

            if (!Helpers.IsMessageAccepted($"Liczba miniatur w folderze Miniatury: {_thumbnailsCount}."))
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
            _targetDirectory = Path.Combine(_projectDirectory, Constants.Folders.Lists);
            Directory.CreateDirectory(_targetDirectory);

            return true;
        }
    }
}