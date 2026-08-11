using SolidEdgeAdd_In.Utils;

namespace SolidEdgeAdd_In.Processors
{
    public class ExportPartsListProcessor
    {
        private readonly SeAssembly _assembly; 
        private readonly SeApp _application; 

        private string _assemblyPath; 
        private string _assemblyName;
        private string _projectDirectory;

        private readonly Dictionary<string, FileData> _data;
        private int _dataCount;

        private List<(string FileName, string Path)> _thumbnails;
        private string _thumbnailsDirectory;
        private int _thumbnailsCount;
        private bool _isGenerateThumbnails;

        private int _multiplier;
        private string _targetDirectory;

        public ExportPartsListProcessor(SeAssembly assembly)
        {
            _assembly = assembly; 
            _application = _assembly.Application;

            _data = new Dictionary<string, FileData>();
            _thumbnails = new List<(string FileName, string Path)>();     
        }

        public bool Initialize()
        {
            _assemblyPath = _assembly.FullName;
            _assemblyName = Path.GetFileNameWithoutExtension(_assemblyPath);
            _projectDirectory = Path.GetDirectoryName(_assemblyPath);

            if (!IsLoaded_Data()) return false;
            if (!IsLoaded_Thumbnails()) return false;
            if (!IsLoaded_GeneratingThumbnails()) return false;
            if (!IsLoaded_Multiplier()) return false;
            if (!IsLoaded_TargetDirectory()) return false;
            if (!IsLoaded_PartsList()) return false;

            return true;
        }

        public void Process()
        {
            ExcelApp excelApp = null;
            ExcelWorkbooks workbooks = null; ExcelWorkbook workbook = null;
            ExcelSheets sheets = null;
            ExcelWorksheet worksheet = null;

            try
            {
                excelApp = new ExcelApp
                { 
                    Visible = false, 
                    DisplayAlerts = false,
                    AskToUpdateLinks = false, 
                    EnableEvents = false 
                };

                workbooks = excelApp.Workbooks; workbook = workbooks.Add();
                excelApp.Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationManual;
                sheets = workbook.Sheets;
                worksheet = (ExcelWorksheet)sheets[1]; worksheet.Paste();

                EditWorksheet(workbook, worksheet);

                Export(workbook);
            }
            finally
            {
                Helpers.ReleaseCom(ref worksheet); Helpers.ReleaseCom(ref sheets);
                Helpers.ReleaseCom(ref workbook); Helpers.ReleaseCom(ref workbooks);
                excelApp?.Quit();
                Helpers.ReleaseCom(ref excelApp);
            }
            
        }

        private void EditWorksheet(ExcelWorkbook workbook, ExcelWorksheet worksheet)
        {
            ExcelRange usedRange = null;
            ExcelRange startCell = null; ExcelRange endCell = null;
            ExcelRange expandedRange = null;

            try
            {
                usedRange = worksheet.UsedRange;

                int typeCol = ExcelUtils.GetColumnIndex(usedRange, Constants.ExcelHeaders.Type);
                int fileNameCol = ExcelUtils.GetColumnIndex(usedRange, Constants.ExcelHeaders.FileName);
                int thumbnailCol = ExcelUtils.GetColumnIndex(usedRange, Constants.ExcelHeaders.Thumbnail);
                int countCol = ExcelUtils.GetColumnIndex(usedRange, Constants.ExcelHeaders.Count);

                if (typeCol == 0 || fileNameCol == 0 || thumbnailCol == 0 || countCol == 0)
                {
                    MessageBox.Show("Selected parts list does not have correct columns.");
                    return;
                }

                int rowCount = usedRange.Rows.Count;
                int colCount = usedRange.Columns.Count;
                int dxfCol = colCount + 1;

                startCell = (ExcelRange)worksheet.Cells[1, 1];
                endCell = (ExcelRange)worksheet.Cells[rowCount, dxfCol];
                expandedRange = worksheet.Range[startCell, endCell];

                object[,] data = (object[,])expandedRange.Value2;

                ExcelUtils.ProcessDataInMemory(data, typeCol, fileNameCol, countCol, dxfCol, rowCount, _multiplier, _data);
                expandedRange.Value2 = data;

                ExcelUtils.Styles(workbook);
                ExcelUtils.Colors(worksheet, typeCol);
                ExcelUtils.Edit(worksheet);

                var thumbnailsDict = _thumbnails.ToDictionary(t => t.FileName, t => t.Path, StringComparer.OrdinalIgnoreCase);
                ReportUtils.SetThumbnails(worksheet, thumbnailsDict, typeCol, fileNameCol, thumbnailCol);
            }
            finally
            {
                Helpers.ReleaseCom(ref expandedRange); Helpers.ReleaseCom(ref endCell);
                Helpers.ReleaseCom(ref startCell); Helpers.ReleaseCom(ref usedRange);
            }
        }

        private void Export(ExcelWorkbook workbook)
        {
            string partsListPath = Path.Combine(_targetDirectory, $"{_assemblyName}_PartsList.xlsx");

            if (File.Exists(partsListPath)) File.Delete(partsListPath); 

            workbook?.SaveAs(partsListPath);
        }

        private bool IsLoaded_Data()
        {
            SeOccurrences occurrences = null;
            try
            {
                occurrences = _assembly.Occurrences;
                DataUtils.BuildDataForExportPartsList(occurrences, _data);
            }
            finally { Helpers.ReleaseCom(ref occurrences); }

            _dataCount = _data.Count;
            if (!Helpers.IsMessageAccepted($"Liczba plików w złożeniu: {_dataCount}.")) return false;

            return true;
        }

        private bool IsLoaded_Thumbnails()
        {
            _thumbnailsDirectory = Path.Combine(_projectDirectory, Constants.Folders.Thumbnails); Directory.CreateDirectory(_thumbnailsDirectory);

            _thumbnails = Directory.GetFiles(_thumbnailsDirectory, "*.jpg", SearchOption.TopDirectoryOnly)
                               .Select(f => (Path.GetFileNameWithoutExtension(f), f)).ToList();

            _thumbnailsCount = _thumbnails.Count;
            if (!Helpers.IsMessageAccepted($"Liczba miniatur w folderze Miniatury: {_thumbnailsCount}.")) return false;

            if (_thumbnailsCount < _dataCount) 
            { 
                if (!Helpers.IsMessageAccepted($"W folderze Miniatury jest mniej miniatur niż plików w złożeniu.")) return false; 
            }

            _isGenerateThumbnails = DialogUtils.IsGenerateThumbnails();

            return true;
        }

        private bool IsLoaded_GeneratingThumbnails()
        {
            if (!_isGenerateThumbnails) return true;

            SeDocument document = null;
            SeWindow window = null;
            foreach (var item in _data)
            {
                try
                {
                    document = Helpers.GetOpenDocument(_application, item.Key);
                    window = _application.ActiveWindow as SeWindow;

                    string thumbnailPath = Path.Combine(_thumbnailsDirectory, item.Value.Name + ".jpg");
                    if (File.Exists(thumbnailPath)) continue; 

                    if (document is SePart part)
                    {
                        Helpers.ManageNonModelElementsInPart(part, false);
                        ReportUtils.SaveThumbnail(thumbnailPath, window);
                        Helpers.ManageNonModelElementsInPart(part, true);
                    }
                    else if (document is SeAssembly assembly)
                    {
                        Helpers.ManageNonModelElementsInAssembly(assembly, false);
                        ReportUtils.SaveThumbnail(thumbnailPath, window);
                        Helpers.ManageNonModelElementsInAssembly(assembly, true);
                    }
                    else if (document is SeSheetMetal sheetMetal)
                    {
                        Helpers.ManageNonModelElementsInSheetMetal(sheetMetal, false);
                        ReportUtils.SaveThumbnail(thumbnailPath, window);
                        Helpers.ManageNonModelElementsInSheetMetal(sheetMetal, true);
                    }

                    if (File.Exists(thumbnailPath)) _thumbnails.Add((item.Value.Name, thumbnailPath));

                }
                catch { continue; }
                finally
                {
                    Helpers.ReleaseCom(ref window);
                    try { document?.Close(false); } catch { }
                    Helpers.ReleaseCom(ref document);

                    System.Windows.Forms.Application.DoEvents();
                    _application.DoIdle();
                }         
            }

            _thumbnailsCount = _thumbnails.Count;
            if (!Helpers.IsMessageAccepted($"Liczba miniatur w folderze Miniatury: {_thumbnailsCount}.")) return false;

            if (_thumbnailsCount < _dataCount)
            {
                if (!Helpers.IsMessageAccepted($"W folderze Miniatury jest mniej miniatur niż plików w złożeniu.")) return false;
            }

            return true;
        }

        private bool IsLoaded_Multiplier()
        {
            SeDocument document = (SeDocument)_assembly;
            using PropertyUtils properties = new(document);
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
            if (!Helpers.IsMessageAccepted($"Przyjęto mnożnik: {_multiplier}.")) return false;

            return true;
        }

        private bool IsLoaded_TargetDirectory()
        {
             _targetDirectory = Path.Combine(_projectDirectory, Constants.Folders.Lists); Directory.CreateDirectory(_targetDirectory);
    
             return true;
        }

        private bool IsLoaded_PartsList()
        {
            try { ReportUtils.CopyPartsList(_application, _assemblyPath); }
            catch { return false; }
            return true;
        }
    }
}