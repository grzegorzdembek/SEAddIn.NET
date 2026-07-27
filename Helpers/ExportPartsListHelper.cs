using SolidEdgeAdd_In.Utils;

namespace SolidEdgeAdd_In.Helpers
{
    public class ExportPartsListHelper
    {
        public static (bool isConfirmed, int multiplier) GetMultiplier(SeAssembly assembly)
        {
            SeDocument document = (SeDocument)assembly; using var properties = new PropertyProvider(document);

            int count = properties.Count; if (count == 0) { return DialogService.GetMultiplier(); }

            return (true, count);
        }

        public static void CopyPartsList(SeAssembly assembly)
        {
            SeApp application = null; SeDocuments documents = null; SeDraft draft = null; SeDraftSheet sheet = null; SeModelLinks modelLinks = null; SeModelLink modelLink = null; SeDrawingViews drawingViews = null; SeDrawingView drawingView = null; SePartsLists partsLists = null; SePartsList partsList = null;

            try
            {
                application = assembly.Application; documents = application.Documents; draft = (SeDraft)documents.Add("SolidEdge.DraftDocument", Missing.Value); sheet = draft.ActiveSheet; modelLinks = draft.ModelLinks; modelLink = modelLinks.Add(assembly.FullName); drawingViews = sheet.DrawingViews;

                drawingView = drawingViews.AddAssemblyView(modelLink, SeViewOrientation.igFrontView, 0.1, 0.2, 0.2, SeAssemblyDrawingViewType.seAssemblyDesignedView);

                partsLists = draft.PartsLists; partsList = partsLists.AddEx(drawingView, 0, DialogService.GetPartsListType(application, assembly), 0, 1);

                for (int i = 0; i < 5; i++) { try { partsList.CopyToClipboard(); System.Threading.Thread.Sleep(300); break; } catch { System.Threading.Thread.Sleep(300); } }
            }
            finally
            {
                CoreUtils.ReleaseCom(ref partsList); CoreUtils.ReleaseCom(ref partsLists); CoreUtils.ReleaseCom(ref drawingView); CoreUtils.ReleaseCom(ref drawingViews); CoreUtils.ReleaseCom(ref modelLink); CoreUtils.ReleaseCom(ref modelLinks); CoreUtils.ReleaseCom(ref sheet);
                if (draft != null) { try { draft.Close(false); } catch { } }
                CoreUtils.ReleaseCom(ref draft); CoreUtils.ReleaseCom(ref documents);
            }
        }

        public static bool HasShots() { return !DialogService.IsShotsNeeded(); }

        public static List<string> GetShots(SeAssembly assembly, bool hasShots)
        {
            List<string> shots = new(); if (hasShots) return shots;

            Dictionary<string, int> data = new(StringComparer.OrdinalIgnoreCase); AssemblyTreeWalker.BuildDataForExportPartsList(assembly.Occurrences, data);

            SeApp application; SeDocument document = null; SeWindow window = null;

            var shotPaths = new List<string>(); string shotsLocation = Path.Combine(Path.GetDirectoryName(assembly.FullName), Constants.Folders.Thumbnails); if (!Directory.Exists(shotsLocation)) Directory.CreateDirectory(shotsLocation);

            try
            {
                application = assembly.Application;

                foreach (var item in data)
                {
                    try
                    {
                        string fileName = Path.GetFileNameWithoutExtension(item.Key); document = CoreUtils.GetOpenDocument(application, item.Key); window = application.ActiveWindow as SeWindow;

                        if (document is SePart pDoc) { CoreUtils.ManageCoordinateSystemsInPart(pDoc, false); shotPaths.Add(RaportGenerationUtils.GetShotPath(Path.Combine(shotsLocation, fileName), window)); CoreUtils.ManageCoordinateSystemsInPart(pDoc, true); continue; }
                        else if (document is SeAssembly aDoc) { CoreUtils.ManageCoordinateSystemsInAssembly(aDoc, false); shotPaths.Add(RaportGenerationUtils.GetShotPath(Path.Combine(shotsLocation, fileName), window)); CoreUtils.ManageCoordinateSystemsInAssembly(aDoc, true); continue; }
                        else if (document is SeSheetMetal smDoc) { CoreUtils.ManageCoordinateSystemsInSheetMetal(smDoc, false); shotPaths.Add(RaportGenerationUtils.GetShotPath(Path.Combine(shotsLocation, fileName), window)); CoreUtils.ManageCoordinateSystemsInSheetMetal(smDoc, true); continue; }
                    }
                    catch { continue; }
                    finally { CoreUtils.ReleaseCom(ref window); if (document != null) { try { document.Close(false); } catch { } } CoreUtils.ReleaseCom(ref document); }
                }
            }
            finally { }

            return shotPaths;
        }

        public static void ExcelObjects(out ExcelApp excelApp, out ExcelWorkbooks workbooks, out ExcelWorkbook workbook, out ExcelSheets sheets, out ExcelWorksheet worksheet)
        {
            excelApp = new ExcelApp { Visible = false, DisplayAlerts = false, AskToUpdateLinks = false, EnableEvents = false };

            workbooks = excelApp.Workbooks; workbook = workbooks.Add(); excelApp.Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationManual; sheets = workbook.Sheets; worksheet = (ExcelWorksheet)sheets[1];

            worksheet.Paste();
        }

        public static void EditWorksheet(SeAssembly assembly, List<string> shots, bool hasShots, ExcelWorkbook workbook, ExcelWorksheet worksheet, int multiplier)
        {
            ExcelRange usedRange = null; ExcelRange rows = null; ExcelRange columns = null; ExcelRange cells = null; ExcelRange startCell = null; ExcelRange endCell = null; ExcelRange expandedRange = null;

            try
            {
                usedRange = worksheet.UsedRange;

                int typeNumber = ExcelWrapper.GetColumnNumber(usedRange, Constants.ExcelHeaders.Type); int nameNumber = ExcelWrapper.GetColumnNumber(usedRange, Constants.ExcelHeaders.Number); int imageNumber = ExcelWrapper.GetColumnNumber(usedRange, Constants.ExcelHeaders.Thumbnail); int countNumber = ExcelWrapper.GetColumnNumber(usedRange, Constants.ExcelHeaders.Quantity);

                if (typeNumber == 0 || nameNumber == 0 || imageNumber == 0 || countNumber == 0) { MessageBox.Show("Wybrana lista czêœci nie posiada odpowiednich kolumn."); return; }

                rows = usedRange.Rows; int rowCount = rows.Count; columns = usedRange.Columns; int colCount = columns.Count; int dxfColNumber = colCount + 1;

                cells = worksheet.Cells; startCell = (ExcelRange)cells[1, 1]; endCell = (ExcelRange)cells[rowCount, dxfColNumber]; expandedRange = worksheet.Range[startCell, endCell]; object[,] data = (object[,])expandedRange.Value2;

                ExcelWrapper.TypeMemory(data, typeNumber, rowCount); ExcelWrapper.CountMemory(data, multiplier, countNumber, rowCount); RaportGenerationUtils.DxfsMemory(Path.GetDirectoryName(assembly.FullName), data, typeNumber, nameNumber, dxfColNumber, rowCount);

                expandedRange.Value2 = data; ExcelWrapper.Styles(workbook); ExcelWrapper.Colors(worksheet, typeNumber); ExcelWrapper.Edit(worksheet);

                string shotDir = Path.Combine(Path.GetDirectoryName(assembly.FullName), Constants.Folders.Thumbnails); if (!Directory.Exists(shotDir)) Directory.CreateDirectory(shotDir);

                RaportGenerationUtils.Shots(worksheet, shots, hasShots, shotDir, typeNumber, nameNumber, imageNumber);
            }
            finally { CoreUtils.ReleaseCom(ref expandedRange); CoreUtils.ReleaseCom(ref endCell); CoreUtils.ReleaseCom(ref startCell); CoreUtils.ReleaseCom(ref cells); CoreUtils.ReleaseCom(ref columns); CoreUtils.ReleaseCom(ref rows); CoreUtils.ReleaseCom(ref usedRange); }
        }

        public static void Export(SeAssembly assembly, ExcelApp excelApp, ExcelWorkbooks workbooks, ExcelWorkbook workbook, ExcelWorksheet worksheet)
        {
            string partsListPath = Path.Combine(Path.GetDirectoryName(assembly.FullName), Path.GetFileNameWithoutExtension(assembly.FullName) + "_PartsList.xlsx");

            if (File.Exists(partsListPath)) { File.Delete(partsListPath); }

            workbook?.SaveAs(partsListPath);
        }

        public static void Release(ref ExcelApp excelApp, ref ExcelWorkbooks workbooks, ref ExcelWorkbook workbook, ref ExcelSheets sheets, ref ExcelWorksheet worksheet)
        {
            if (excelApp != null) { excelApp.DisplayAlerts = false; excelApp.Quit(); }

            CoreUtils.ReleaseCom(ref worksheet); CoreUtils.ReleaseCom(ref sheets); CoreUtils.ReleaseCom(ref workbook); CoreUtils.ReleaseCom(ref workbooks); CoreUtils.ReleaseCom(ref excelApp);
        }
    }
}