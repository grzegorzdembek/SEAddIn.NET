using SolidEdgeAdd_In.Utils;

namespace SolidEdgeAdd_In.Processors
{
    public class ExportPartsListProcessor
    {
        private readonly SeAssembly _assembly;

        private string _assemblyFilePath;
        private string _assemblyDirectory;

        private int _multiplier;
        private bool _hasShots;
        private List<string> _shots;

        public ExportPartsListProcessor(SeAssembly assembly)
        {
            _assembly = assembly;
            _shots = new List<string>();
        }

        public bool Initialize()
        {
            _assemblyFilePath = _assembly.FullName;
            _assemblyDirectory = Path.GetDirectoryName(_assemblyFilePath);

            SeDocument document = (SeDocument)_assembly;
            using var properties = new PropertyUtils(document);

            int count = properties.Count;
            if (count == 0)
            {
                var result = DialogUtils.GetMultiplier();
                if (!result.isConfirmed) { return false; }
                _multiplier = result.multiplier;
            }
            else { _multiplier = count; }

            _hasShots = !DialogUtils.IsShotsNeeded();
            LoadShots();

            return true;
        }

        public void Process()
        {
            CopyPartsList();

            ExcelApp excelApp = null;
            ExcelWorkbooks workbooks = null;
            ExcelWorkbook workbook = null;
            ExcelSheets sheets = null;
            ExcelWorksheet worksheet = null;

            try
            {
                excelApp = new ExcelApp { Visible = false, DisplayAlerts = false, AskToUpdateLinks = false, EnableEvents = false };
                workbooks = excelApp.Workbooks;
                workbook = workbooks.Add();
                excelApp.Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationManual;
                sheets = workbook.Sheets;
                worksheet = (ExcelWorksheet)sheets[1];

                worksheet.Paste();
                EditWorksheet(workbook, worksheet);
                Export(workbook);
            }
            finally
            {
                if (excelApp != null) { excelApp.DisplayAlerts = false; excelApp.Quit(); }
                Helpers.ReleaseCom(ref worksheet); Helpers.ReleaseCom(ref sheets);
                Helpers.ReleaseCom(ref workbook); Helpers.ReleaseCom(ref workbooks);
                Helpers.ReleaseCom(ref excelApp);
            }
        }

        private void CopyPartsList()
        {
            SeApp application = null;
            SeDocuments documents = null;
            SeDraft draft = null;
            SeDraftSheet sheet = null;
            SeModelLinks modelLinks = null;
            SeModelLink modelLink = null;
            SeDrawingViews drawingViews = null;
            SeDrawingView drawingView = null;
            SePartsLists partsLists = null;
            SePartsList partsList = null;

            try
            {
                application = _assembly.Application;
                documents = application.Documents;
                draft = (SeDraft)documents.Add("SolidEdge.DraftDocument", Missing.Value);
                sheet = draft.ActiveSheet;
                modelLinks = draft.ModelLinks;

                modelLink = modelLinks.Add(_assemblyFilePath);
                drawingViews = sheet.DrawingViews;
                drawingView = drawingViews.AddAssemblyView(modelLink, SeViewOrientation.igFrontView, 0.1, 0.2, 0.2, SeAssemblyDrawingViewType.seAssemblyDesignedView);

                partsLists = draft.PartsLists;
                partsList = partsLists.AddEx(drawingView, 0, DialogUtils.GetPartsListType(application, _assembly), 0, 1);

                for (int i = 0; i < 5; i++)
                {
                    try { partsList.CopyToClipboard(); System.Threading.Thread.Sleep(300); break; }
                    catch { System.Threading.Thread.Sleep(300); }
                }
            }
            finally
            {
                Helpers.ReleaseCom(ref partsList); Helpers.ReleaseCom(ref partsLists);
                Helpers.ReleaseCom(ref drawingView); Helpers.ReleaseCom(ref drawingViews);
                Helpers.ReleaseCom(ref modelLink); Helpers.ReleaseCom(ref modelLinks);
                Helpers.ReleaseCom(ref sheet);

                if (draft != null) { try { draft.Close(false); } catch { } }

                Helpers.ReleaseCom(ref draft); Helpers.ReleaseCom(ref documents);
            }
        }

        private void LoadShots()
        {
            if (_hasShots) { return; }

            Dictionary<string, int> data = new(StringComparer.OrdinalIgnoreCase);
            DataUtils.BuildDataForExportPartsList(_assembly.Occurrences, data);

            SeApp application;
            SeDocument document = null;
            SeWindow window = null;

            string shotsDirectoryPath = Path.Combine(_assemblyDirectory, Constants.Folders.Thumbnails);
            if (!Directory.Exists(shotsDirectoryPath)) { Directory.CreateDirectory(shotsDirectoryPath); }

            try
            {
                application = _assembly.Application;

                foreach (var item in data)
                {
                    try
                    {
                        string fileName = Path.GetFileNameWithoutExtension(item.Key);
                        document = Helpers.GetOpenDocument(application, item.Key);
                        window = application.ActiveWindow as SeWindow;

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

        private void EditWorksheet(ExcelWorkbook workbook, ExcelWorksheet worksheet)
        {
            ExcelRange usedRange = null;
            ExcelRange rows = null;
            ExcelRange columns = null;
            ExcelRange cells = null;
            ExcelRange startCell = null;
            ExcelRange endCell = null;
            ExcelRange expandedRange = null;

            try
            {
                usedRange = worksheet.UsedRange;

                int typeColumnIndex = ExcelUtils.GetColumnNumber(usedRange, Constants.ExcelHeaders.Type);
                int fileNameColumnIndex = ExcelUtils.GetColumnNumber(usedRange, Constants.ExcelHeaders.FileName);
                int thumbnailColumnIndex = ExcelUtils.GetColumnNumber(usedRange, Constants.ExcelHeaders.Thumbnail);
                int countColumnIndex = ExcelUtils.GetColumnNumber(usedRange, Constants.ExcelHeaders.Count);

                if (typeColumnIndex == 0 || fileNameColumnIndex == 0 || thumbnailColumnIndex == 0 || countColumnIndex == 0)
                {
                    MessageBox.Show("Selected parts list does not have correct columns.");
                    return;
                }

                rows = usedRange.Rows;
                int rowCount = rows.Count;
                columns = usedRange.Columns;
                int colCount = columns.Count;
                int dxfColumnIndex = colCount + 1;

                cells = worksheet.Cells;
                startCell = (ExcelRange)cells[1, 1];
                endCell = (ExcelRange)cells[rowCount, dxfColumnIndex];
                expandedRange = worksheet.Range[startCell, endCell];

                object[,] data = (object[,])expandedRange.Value2;

                ExcelUtils.TypeMemory(data, typeColumnIndex, rowCount);
                ExcelUtils.CountMemory(data, _multiplier, countColumnIndex, rowCount);

                RaportUtils.DxfsMemory(_assemblyDirectory, data, typeColumnIndex, fileNameColumnIndex, dxfColumnIndex, rowCount);

                expandedRange.Value2 = data;

                ExcelUtils.Styles(workbook);
                ExcelUtils.Colors(worksheet, typeColumnIndex);
                ExcelUtils.Edit(worksheet);

                string shotsDirectoryPath = Path.Combine(_assemblyDirectory, Constants.Folders.Thumbnails);
                if (!Directory.Exists(shotsDirectoryPath)) { Directory.CreateDirectory(shotsDirectoryPath); }

                RaportUtils.Shots(worksheet, _shots, _hasShots, shotsDirectoryPath, typeColumnIndex, fileNameColumnIndex, thumbnailColumnIndex);
            }
            finally
            {
                Helpers.ReleaseCom(ref expandedRange); Helpers.ReleaseCom(ref endCell);
                Helpers.ReleaseCom(ref startCell); Helpers.ReleaseCom(ref cells);
                Helpers.ReleaseCom(ref columns); Helpers.ReleaseCom(ref rows);
                Helpers.ReleaseCom(ref usedRange);
            }
        }

        private void Export(ExcelWorkbook workbook)
        {
            string assemblyFileName = Path.GetFileNameWithoutExtension(_assemblyFilePath);
            string partsListPath = Path.Combine(_assemblyDirectory, $"{assemblyFileName}_PartsList.xlsx");

            if (File.Exists(partsListPath)) { File.Delete(partsListPath); }
            workbook?.SaveAs(partsListPath);
        }
    }
}