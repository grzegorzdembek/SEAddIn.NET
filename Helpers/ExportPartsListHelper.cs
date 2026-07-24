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
            SeDocuments documents = null; SeDraft draft = null; SeDraftSheet sheet = null; SeModelLinks modelLinks = null; SeModelLink modelLink = null; SeDrawingViews drawingViews = null; SeDrawingView drawingView = null; SePartsLists partsLists = null; SePartsList partsList = null;

            try
            {
                documents = assembly.Application.Documents; draft = (SeDraft)documents.Add("SolidEdge.DraftDocument", Missing.Value); sheet = draft.ActiveSheet; modelLinks = draft.ModelLinks; modelLink = modelLinks.Add(assembly.FullName); drawingViews = sheet.DrawingViews;

                drawingView = drawingViews.AddAssemblyView(modelLink, SeViewOrientation.igFrontView, 0.1, 0.2, 0.2, SeAssemblyDrawingViewType.seAssemblyDesignedView);

                partsLists = draft.PartsLists; partsList = partsLists.AddEx(drawingView, 0, DialogService.GetPartsListType(assembly.Application, assembly), 0, 1); partsList.CopyToClipboard(); System.Threading.Thread.Sleep(200);
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
            var shots = new List<string>(); if (hasShots) return shots;

            Dictionary<string, int> data = new(StringComparer.OrdinalIgnoreCase); AssemblyTreeWalker.BuildDataForExportPartsList(assembly.Occurrences, data);

            SeDocument document = null; SeAssembly subAssembly = null; SePart part = null; SeSheetMetal sheetMetal = null; SeWindow window = null;

            var shotPaths = new List<string>(); string shotsLocation = Path.Combine(Path.GetDirectoryName(assembly.FullName), "Miniatury"); if (!Directory.Exists(shotsLocation)) Directory.CreateDirectory(shotsLocation);

            foreach (var item in data)
            {
                try
                {
                    string fileName = Path.GetFileNameWithoutExtension(item.Key); document = CoreUtils.GetOpenDocument(assembly.Application, item.Key); window = assembly.Application.ActiveWindow as SeWindow;

                    if (document is SePart pDoc) { part = pDoc; CoreUtils.ManageCoordinateSystemsInPart(part, false); shotPaths.Add(RaportGenerationUtils.GetShotPath(Path.Combine(shotsLocation, fileName), window)); CoreUtils.ManageCoordinateSystemsInPart(part, true); continue; }
                    else if (document is SeAssembly aDoc) { subAssembly = aDoc; CoreUtils.ManageCoordinateSystemsInAssembly(subAssembly, false); shotPaths.Add(RaportGenerationUtils.GetShotPath(Path.Combine(shotsLocation, fileName), window)); CoreUtils.ManageCoordinateSystemsInAssembly(subAssembly, true); continue; }
                    else if (document is SeSheetMetal smDoc) { sheetMetal = smDoc; CoreUtils.ManageCoordinateSystemsInSheetMetal(sheetMetal, false); shotPaths.Add(RaportGenerationUtils.GetShotPath(Path.Combine(shotsLocation, fileName), window)); CoreUtils.ManageCoordinateSystemsInSheetMetal(sheetMetal, true); continue; }
                }
                catch { continue; }
                finally { CoreUtils.ReleaseCom(ref window); CoreUtils.ReleaseCom(ref sheetMetal); CoreUtils.ReleaseCom(ref part); CoreUtils.ReleaseCom(ref subAssembly); document?.Close(false); CoreUtils.ReleaseCom(ref document); }
            }

            return shotPaths;
        }

        public static void ExcelObjects(out ExcelApp excelApp, out ExcelWorkbooks workbooks, out ExcelWorkbook workbook, out ExcelSheets sheets, out ExcelWorksheet worksheet)
        {
            excelApp = new ExcelApp { Visible = false, DisplayAlerts = false, AskToUpdateLinks = false, ScreenUpdating = false, EnableEvents = false };

            workbooks = excelApp.Workbooks; workbook = workbooks.Add(); excelApp.Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationManual; sheets = workbook.Sheets; worksheet = sheets[1];

            worksheet.Paste();
        }

        public static void EditWorksheet(SeAssembly assembly, List<string> shots, bool hasShots, ExcelWorkbook workbook, ExcelWorksheet worksheet, int multiplier)
        {
            int typeNumber = ExcelWrapper.GetColumnNumber(worksheet.UsedRange, "Typ"); int nameNumber = ExcelWrapper.GetColumnNumber(worksheet.UsedRange, "Numer czêœci"); int imageNumber = ExcelWrapper.GetColumnNumber(worksheet.UsedRange, "Miniatura"); int countNumber = ExcelWrapper.GetColumnNumber(worksheet.UsedRange, "Iloœæ");

            if (typeNumber == 0 || nameNumber == 0 || imageNumber == 0 || countNumber == 0) { MessageBox.Show("Wybrana lista czêœci nie zawiera kolumny Typ lub Numer czêœci lub Miniatura"); return; }

            ExcelRange usedRange = null; ExcelRange expandedRange = null;

            try
            {
                usedRange = worksheet.UsedRange; int rowCount = usedRange.Rows.Count; int colCount = usedRange.Columns.Count; int dxfColNumber = colCount + 1;

                expandedRange = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[rowCount, dxfColNumber]]; object[,] data = (object[,])expandedRange.Value2;

                ExcelWrapper.TypeMemory(data, typeNumber, rowCount); ExcelWrapper.CountMemory(data, multiplier, countNumber, rowCount); RaportGenerationUtils.DxfsMemory(Path.GetDirectoryName(assembly.FullName), data, typeNumber, nameNumber, dxfColNumber, rowCount);

                expandedRange.Value2 = data;
            }
            finally { CoreUtils.ReleaseCom(ref expandedRange); CoreUtils.ReleaseCom(ref usedRange); }

            ExcelWrapper.Styles(workbook); ExcelWrapper.Colors(worksheet, typeNumber); ExcelWrapper.Edit(worksheet);

            string shotDir = Path.Combine(Path.GetDirectoryName(assembly.FullName), "Miniatury"); if (!Directory.Exists(shotDir)) Directory.CreateDirectory(shotDir);

            RaportGenerationUtils.Shots(worksheet, shots, hasShots, shotDir, typeNumber, nameNumber, imageNumber);
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