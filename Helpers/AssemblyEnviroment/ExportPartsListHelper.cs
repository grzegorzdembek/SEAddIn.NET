using SolidEdgeAdd_In.Utils; 
using System;
using System.Collections.Generic;
using System.IO; 
using System.Reflection;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel; 

namespace SolidEdgeAdd_In.Helpers.AssemblyEnviroment
{
    public class ExportPartsListHelper
    {
        public static int GetMultiplier(SolidEdgeAssembly.AssemblyDocument assembly)
        {
            SolidEdgeFramework.SolidEdgeDocument document = (SolidEdgeFramework.SolidEdgeDocument)assembly;
            int count = PropertyProvider.GetCount(document);
            if (count == 0) PropertyProvider.SetCount(document, count = DialogService.GetMultiplier());
            return count;
        }

        public static void CopyPartsList(SolidEdgeAssembly.AssemblyDocument assembly)
        {
            SolidEdgeFramework.Documents documents = null;
            SolidEdgeDraft.DraftDocument draft = null;
            SolidEdgeDraft.Sheet sheet = null;
            SolidEdgeDraft.ModelLinks modelLinks = null;
            SolidEdgeDraft.ModelLink modelLink = null;
            SolidEdgeDraft.DrawingViews drawingViews = null;
            SolidEdgeDraft.DrawingView drawingView = null;
            SolidEdgeDraft.PartsLists partsLists = null;
            SolidEdgeDraft.PartsList partsList = null;
            try
            {
                documents = assembly.Application.Documents;
                draft = (SolidEdgeDraft.DraftDocument)documents.Add("SolidEdge.DraftDocument", Missing.Value);
                sheet = draft.ActiveSheet;
                modelLinks = draft.ModelLinks;
                modelLink = modelLinks.Add(assembly.FullName);
                drawingViews = sheet.DrawingViews;

                drawingView = drawingViews.AddAssemblyView(
                        modelLink,
                        SolidEdgeDraft.ViewOrientationConstants.igFrontView, 0.1, 0.2, 0.2,
                        SolidEdgeDraft.AssemblyDrawingViewTypeConstants.seAssemblyDesignedView);

                partsLists = draft.PartsLists;
                partsList = partsLists.AddEx(drawingView, 0, DialogService.GetPartsListType(assembly.Application, assembly), 0, 1);
                partsList.CopyToClipboard();
                System.Threading.Thread.Sleep(200);
            }
            finally
            {
                draft?.Close(false);
                CoreUtils.ReleaseCom(ref partsList);
                CoreUtils.ReleaseCom(ref partsLists);
                CoreUtils.ReleaseCom(ref drawingView);
                CoreUtils.ReleaseCom(ref drawingViews);
                CoreUtils.ReleaseCom(ref modelLink);
                CoreUtils.ReleaseCom(ref modelLinks);
                CoreUtils.ReleaseCom(ref sheet);
                CoreUtils.ReleaseCom(ref draft);
                CoreUtils.ReleaseCom(ref documents);
            }
        }

        public static bool HasShots()
        {
            return !DialogService.IsShotsNeeded();
        }

        public static List<string> GetShots(SolidEdgeAssembly.AssemblyDocument assembly, bool hasShots)
        {
            var shots = new List<string>();
            if (hasShots) return shots;

            var occurrences = new Dictionary<string, int>();
            AssemblyTreeWalker.AllOccurrences(assembly.Occurrences, occurrences);

            SolidEdgeFramework.SolidEdgeDocument document = null;
            SolidEdgeAssembly.AssemblyDocument subAssembly = null;
            SolidEdgePart.PartDocument part = null;
            SolidEdgePart.SheetMetalDocument sheetMetal = null;
            SolidEdgeFramework.Window window = null;

            var shotPaths = new List<string>();
            string shotsLocation = Path.Combine(Path.GetDirectoryName(assembly.FullName), "Miniatury");
            if (!Directory.Exists(shotsLocation)) Directory.CreateDirectory(shotsLocation);

            foreach (var occurrence in occurrences)
            {
                try
                {
                    string fileName = Path.GetFileNameWithoutExtension(occurrence.Key);

                    document = CoreUtils.GetOpenDocument(assembly.Application, occurrence.Key);
                    window = assembly.Application.ActiveWindow as SolidEdgeFramework.Window;

                    if (document is SolidEdgePart.PartDocument pDoc)
                    {
                        part = (SolidEdgePart.PartDocument)pDoc;
                        CoreUtils.ManageCoordinateSystemsInPart(part, false);
                        var shotPath = RaportGenerationUtils.GetShotPath(Path.Combine(shotsLocation, fileName), window);
                        shotPaths.Add(shotPath);
                        CoreUtils.ManageCoordinateSystemsInPart(part, true);
                        continue;
                    }
                    else if (document is SolidEdgeAssembly.AssemblyDocument aDoc)
                    {
                        subAssembly = (SolidEdgeAssembly.AssemblyDocument)aDoc;
                        CoreUtils.ManageCoordinateSystemsInAssembly(subAssembly, false);
                        var shotPath = RaportGenerationUtils.GetShotPath(Path.Combine(shotsLocation, fileName), window);
                        shotPaths.Add(shotPath);
                        CoreUtils.ManageCoordinateSystemsInAssembly(subAssembly, true);
                        continue;
                    }
                    else if (document is SolidEdgePart.SheetMetalDocument smDoc)
                    {
                        sheetMetal = (SolidEdgePart.SheetMetalDocument)smDoc;
                        CoreUtils.ManageCoordinateSystemsInSheetMetal(sheetMetal, false);
                        var shotPath = RaportGenerationUtils.GetShotPath(Path.Combine(shotsLocation, fileName), window);
                        shotPaths.Add(shotPath);
                        CoreUtils.ManageCoordinateSystemsInSheetMetal(sheetMetal, true);
                        continue;
                    }
                }
                catch { continue; }
                finally
                {
                    document?.Close(false);
                    CoreUtils.ReleaseCom(ref window);
                    CoreUtils.ReleaseCom(ref sheetMetal);
                    CoreUtils.ReleaseCom(ref part);
                    CoreUtils.ReleaseCom(ref subAssembly);
                    CoreUtils.ReleaseCom(ref document);
                }
            }
            return shotPaths;
        }

        public static void ExcelObjects(out Excel.Application excelApp, out Excel.Workbooks workbooks, out Excel.Workbook workbook, out Excel.Sheets sheets, out Excel.Worksheet worksheet)
        {
            excelApp = new Excel.Application { Visible = false };

            excelApp.DisplayAlerts = false;
            excelApp.AskToUpdateLinks = false;

            workbooks = excelApp.Workbooks;
            workbook = workbooks.Add();

            sheets = workbook.Sheets; 
            worksheet = sheets[1];

            worksheet.Paste();
        }

        public static void EditWorksheet(SolidEdgeAssembly.AssemblyDocument assembly, List<string> shots, bool hasShots, Excel.Workbook workbook, Excel.Worksheet worksheet, int multiplier)
        {
            int typeNumber = ExcelWrapper.GetColumnNumber(worksheet.UsedRange,"Typ");
            int nameNumber = ExcelWrapper.GetColumnNumber(worksheet.UsedRange, "Numer czêœci");
            int imageNumber = ExcelWrapper.GetColumnNumber(worksheet.UsedRange, "Miniatura");
            int countNumber = ExcelWrapper.GetColumnNumber(worksheet.UsedRange, "Iloœæ");

            if (typeNumber == 0 || nameNumber == 0 || imageNumber == 0 || countNumber == 0)
            {
                MessageBox.Show("Wybrana lista czêœci nie zawiera kolumny Typ lub Numer czêœci lub Miniatura");
                return;
            }

            ExcelWrapper.Styles(workbook);
            ExcelWrapper.Type(worksheet, typeNumber);
            ExcelWrapper.Colors(worksheet, typeNumber);
            ExcelWrapper.Count(worksheet, multiplier, countNumber);
            RaportGenerationUtils.Dxfs(Path.GetDirectoryName(assembly.FullName), worksheet, typeNumber, nameNumber);
            ExcelWrapper.Edit(worksheet);

            string shotDir = Path.Combine(Path.GetDirectoryName(assembly.FullName), "Miniatury");
            if (!Directory.Exists(shotDir)) Directory.CreateDirectory(shotDir);
            RaportGenerationUtils.Shots(worksheet, shots, hasShots, shotDir, typeNumber, nameNumber, imageNumber);
        }

        public static void Export(SolidEdgeAssembly.AssemblyDocument assembly, Excel.Application excelApp, Excel.Workbooks workbooks, Excel.Workbook workbook, Excel.Worksheet worksheet)
        {
            string partsListPath = Path.Combine(Path.GetDirectoryName(assembly.FullName),
            Path.GetFileNameWithoutExtension(assembly.FullName) + "_PartsList.xlsx");

            if (File.Exists(partsListPath))
            {
                File.Delete(partsListPath);
            }

            workbook?.SaveAs(partsListPath);
        }
        public static void Release(ref Excel.Application excelApp, ref Excel.Workbooks workbooks, ref Excel.Workbook workbook, ref Excel.Sheets sheets, ref Excel.Worksheet worksheet)
        {
            if (excelApp != null)
            {
                excelApp.DisplayAlerts = false;
                excelApp.Quit();
            }

            CoreUtils.ReleaseCom(ref worksheet);
            CoreUtils.ReleaseCom(ref sheets);
            CoreUtils.ReleaseCom(ref workbook);
            CoreUtils.ReleaseCom(ref workbooks);
            CoreUtils.ReleaseCom(ref excelApp);

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}