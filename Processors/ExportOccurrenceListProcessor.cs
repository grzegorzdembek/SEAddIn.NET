using SolidEdgeAdd_In.Utils;

namespace SolidEdgeAdd_In.Processors
{
    public class ExportOccurrenceListProcessor
    {
        private readonly SeAssembly _assembly;
        private readonly SeApp _application;
        private readonly Dictionary<string, FileData> _occurrencesData;
        private readonly List<string> _shots;

        private List<string> _types;

        private string _assemblyFilePath;
        private string _assemblyDirectory;

        private int _multiplier;
        private bool _hasShots;

        public ExportOccurrenceListProcessor(SeAssembly assembly)
        {
            _assembly = assembly;
            _application = assembly.Application;
            _occurrencesData = new Dictionary<string, FileData>(StringComparer.OrdinalIgnoreCase);
            _shots = new List<string>();
        }

        public bool Initialize()
        {
            _assemblyFilePath = _assembly.FullName;
            _assemblyDirectory = Path.GetDirectoryName(_assemblyFilePath);

            if (!TryGetMultiplier()) { return false; }
            if (!GetTypes()) { return false; }

            LoadOccurrencesData();

            if (_occurrencesData.Count == 0)
            {
                MessageBox.Show("No occurrences found to process.", "Stop", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }

            _hasShots = !DialogUtils.IsShotsNeeded();
            LoadShots();

            return true;
        }

        public void Process()
        {
            ExcelApp excelApp = null;
            ExcelWorkbooks workbooks = null;
            string assemblyFileName = Path.GetFileNameWithoutExtension(_assemblyFilePath);

            try
            {
                excelApp = new ExcelApp { Visible = false, DisplayAlerts = false, ScreenUpdating = false, EnableEvents = false };
                workbooks = excelApp.Workbooks;
                excelApp.Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationManual;

                foreach (string currentType in _types)
                {
                    var filteredData = _occurrencesData.Where(kvp => kvp.Value.Type == currentType).ToList();
                    if (filteredData.Count == 0) { continue; }

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
                        finally { Helpers.ReleaseCom(ref startHeaderCell); Helpers.ReleaseCom(ref endHeaderCell); Helpers.ReleaseCom(ref headerRange); }

                        object[,] excelData = new object[filteredData.Count, 6];
                        int dataRowIndex = 0;

                        var sortedData = filteredData.OrderBy(item => item.Value.FileName).ToList();

                        foreach (var item in sortedData)
                        {
                            excelData[dataRowIndex, 0] = lp++;
                            excelData[dataRowIndex, 1] = item.Value.FileName;
                            excelData[dataRowIndex, 2] = item.Value.Title;
                            excelData[dataRowIndex, 3] = item.Value.Type;
                            excelData[dataRowIndex, 4] = item.Value.OccurrenceCount * _multiplier;
                            excelData[dataRowIndex, 5] = "";
                            dataRowIndex++;
                        }

                        try
                        {
                            startDataCell = (ExcelRange)cells[2, 1];
                            endDataCell = (ExcelRange)cells[filteredData.Count + 1, 6];
                            writeRange = worksheet.Range[startDataCell, endDataCell];
                            writeRange.Value = excelData;
                        }
                        finally { Helpers.ReleaseCom(ref startDataCell); Helpers.ReleaseCom(ref endDataCell); Helpers.ReleaseCom(ref writeRange); }

                        // EDIT & FORMAT
                        ExcelUtils.Edit(worksheet);

                        if (!_hasShots)
                        {
                            usedRange = worksheet.UsedRange;
                            int typeColumnIndex = ExcelUtils.GetColumnNumber(usedRange, Constants.ExcelHeaders.Type);
                            int fileNameColumnIndex = ExcelUtils.GetColumnNumber(usedRange, Constants.ExcelHeaders.FileName);
                            int thumbnailColumnIndex = ExcelUtils.GetColumnNumber(usedRange, Constants.ExcelHeaders.Thumbnail);

                            string shotsDirectoryPath = Path.Combine(_assemblyDirectory, Constants.Folders.Thumbnails);

                            if (typeColumnIndex > 0 && fileNameColumnIndex > 0 && thumbnailColumnIndex > 0)
                            {
                                RaportUtils.Shots(worksheet, _shots, _hasShots, shotsDirectoryPath, typeColumnIndex, fileNameColumnIndex, thumbnailColumnIndex);
                            }
                        }

                        string excelFilePath = Path.Combine(_assemblyDirectory, $"{assemblyFileName}_{currentType}.xlsx");
                        if (File.Exists(excelFilePath)) { File.Delete(excelFilePath); }
                        workbook.SaveAs(excelFilePath);
                    }
                    finally
                    {
                        Helpers.ReleaseCom(ref columns); Helpers.ReleaseCom(ref usedRange);
                        Helpers.ReleaseCom(ref cells); Helpers.ReleaseCom(ref worksheet);
                        Helpers.ReleaseCom(ref xlSheets);

                        if (workbook != null) { try { workbook.Close(false); } catch { } }
                        Helpers.ReleaseCom(ref workbook);
                    }
                }
            }
            finally
            {
                if (workbooks != null) { Helpers.ReleaseCom(ref workbooks); }
                if (excelApp != null) { try { excelApp.Quit(); } catch { } Helpers.ReleaseCom(ref excelApp); }
            }
        }

        private bool GetTypes()
        {
            var selectedStrings = DialogUtils.GetSelectedTypes();
            if (selectedStrings == null || selectedStrings.Count == 0) { return false; }

            _types = new List<string>();
            foreach (var str in selectedStrings)
            {
                int start = str.IndexOf('(');
                int end = str.IndexOf(')');
                if (start >= 0 && end > start) { _types.Add(str.Substring(start + 1, end - start - 1)); }
            }
            return _types.Count > 0;
        }

        private bool TryGetMultiplier()
        {
            SeDocument document = (SeDocument)_assembly;
            using var properties = new PropertyUtils(document);
            int count = properties.Count;

            if (count == 0)
            {
                var (isConfirmed, multiplier) = DialogUtils.GetMultiplier();
                if (isConfirmed) { properties.Count = multiplier; _multiplier = multiplier; return true; }
                return false;
            }

            _multiplier = count;
            return true;
        }

        private void LoadOccurrencesData()
        {
            SeOccurrences occurrences = null;
            try
            {
                occurrences = _assembly.Occurrences;
                DataUtils.BuildDataForExportOccurrenceList(occurrences, _occurrencesData, _types);
            }
            finally { Helpers.ReleaseCom(ref occurrences); }
        }

        private void LoadShots()
        {
            if (_hasShots) { return; }

            SeDocument document = null;
            SeWindow window = null;

            string shotsDirectoryPath = Path.Combine(_assemblyDirectory, Constants.Folders.Thumbnails);
            if (!Directory.Exists(shotsDirectoryPath)) { Directory.CreateDirectory(shotsDirectoryPath); }

            try
            {
                foreach (var item in _occurrencesData)
                {
                    try
                    {
                        string fileName = Path.GetFileNameWithoutExtension(item.Key);
                        document = Helpers.GetOpenDocument(_application, item.Key);
                        window = _application.ActiveWindow as SeWindow;

                        if (document is SePart pDoc)
                        {
                            Helpers.ManageCoordinateSystemsInPart(pDoc, false);
                            _shots.Add(RaportUtils.GetShotPath(Path.Combine(shotsDirectoryPath, fileName), window));
                            Helpers.ManageCoordinateSystemsInPart(pDoc, true);
                        }
                        else if (document is SeAssembly aDoc)
                        {
                            Helpers.ManageCoordinateSystemsInAssembly(aDoc, false);
                            _shots.Add(RaportUtils.GetShotPath(Path.Combine(shotsDirectoryPath, fileName), window));
                            Helpers.ManageCoordinateSystemsInAssembly(aDoc, true);
                        }
                        else if (document is SeSheetMetal smDoc)
                        {
                            Helpers.ManageCoordinateSystemsInSheetMetal(smDoc, false);
                            _shots.Add(RaportUtils.GetShotPath(Path.Combine(shotsDirectoryPath, fileName), window));
                            Helpers.ManageCoordinateSystemsInSheetMetal(smDoc, true);
                        }
                    }
                    catch { continue; }
                    finally
                    {
                        Helpers.ReleaseCom(ref window);
                        if (document != null) { try { document.Close(false); } catch { } }
                        Helpers.ReleaseCom(ref document);
                    }
                }
            }
            finally { }
        }
    }
}