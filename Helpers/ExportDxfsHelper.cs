using SolidEdgeAdd_In.Utils;

namespace SolidEdgeAdd_In.Helpers
{
    public class ExportDxfsHelper
    {
        public static (bool isConfirmed, int multiplier) GetMultiplier(SeAssembly assembly)
        {
            SeDocument document = (SeDocument)assembly;

            using var properties = new PropertyProvider(document);

            int count = properties.Count;
            if (count == 0)
            {
                var result = DialogService.GetMultiplier();
                if (result.isConfirmed) { properties.Count = result.multiplier; return result; }
                return (false, 1);
            }

            return (true, count);
        }

        public static Dictionary<string, FileData> GetData(SeAssembly assembly, DxfExportLogger logger)
        {
            Dictionary<string, FileData> data = new(StringComparer.OrdinalIgnoreCase);
            AssemblyTreeWalker.BuildDataForExportDxfs(assembly.Occurrences, data, logger);
            return data;
        }

        public static string GetSubDirectory(SeAssembly assembly)
        {
            string fileName = Path.GetFileNameWithoutExtension(assembly.FullName);
            string number = fileName.Length >= 4 ? fileName.Substring(0, 4) : fileName;
            string date = DateTime.Now.ToString("yyyy-MM-dd");

            string packagesDirectory = Path.Combine(Path.GetDirectoryName(assembly.FullName), "Paczki");
            if (!Directory.Exists(packagesDirectory)) Directory.CreateDirectory(packagesDirectory);

            string subDirectory = Path.Combine(packagesDirectory, $"{number}_{date}");
            if (!Directory.Exists(subDirectory)) Directory.CreateDirectory(subDirectory);

            return subDirectory;
        }

        public static void ProcessData(SeAssembly assembly, Dictionary<string, FileData> data, string subDirectory, int multiplier, DxfExportLogger logger)
        {
            SeDocument document = null;
            SeSheetMetal metalSheet = null; SePart part = null;
            SeModels models = null; SeFlatPatternModels flatPatterns = null;

            ExcelApp excelApp = null;
            ExcelWorkbooks workbooks = null; ExcelWorkbook workbook = null;
            ExcelSheets xlSheets = null; ExcelWorksheet worksheet = null;
            ExcelRange headerRange = null; ExcelRange usedRange = null;
            ExcelRange columns = null;

            string mainDirectory = Path.GetDirectoryName(assembly.FullName);

            try
            {
                excelApp = new ExcelApp
                {
                    Visible = false,
                    DisplayAlerts = false,
                    ScreenUpdating = false,
                    Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationManual,
                    EnableEvents = false
                };
                workbooks = excelApp.Workbooks;
                workbook = workbooks.Add();
                xlSheets = workbook.Sheets;
                worksheet = (ExcelWorksheet)xlSheets[1];

                worksheet.Cells[1, 1] = "Nr części";
                worksheet.Cells[1, 2] = "Grubość";
                worksheet.Cells[1, 3] = "Szerokość";
                worksheet.Cells[1, 4] = "Długość";
                worksheet.Cells[1, 5] = "Gatunek";
                worksheet.Cells[1, 6] = "Ilość";

                headerRange = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, 6]];
                headerRange.Font.Bold = true;
                headerRange.Interior.Color = ColorTranslator.ToOle(Color.LightGray);
                headerRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                object[,] excelData = new object[data.Count, 6];
                int dataRowIndex = 0;

                foreach (var item in data)
                {
                    bool isOpen = false;
                    /* Business Decision For Naming Files:
                     * w glownym folderze pomijamy ilosc
                     */
                    string name = item.Value.Name;
                    string thickness = item.Value.Thickness;
                    int count = item.Value.OccurrenceCount * multiplier;
                    string material = item.Value.Material;
                    string sizeX = item.Value.SizeX;
                    string sizeY = item.Value.SizeY;
                    string dxfDate = item.Value.DxfDate;

                    string mainDxfName = $"{thickness}mm_{material}_{name}.dxf";
                    string subDxfName = $"{thickness}mm_{count}szt_{material}_{name}.dxf";

                    string mainDxfPath = Path.Combine(mainDirectory, mainDxfName);
                    string subDxfPath = Path.Combine(subDirectory, subDxfName);

                    try
                    {
                        /*
                         * Business Decision For Genereting: 
                         * zapisz jesli brakuje daty wygenerowania dxf. - OK.
                         * zawsze przekopiuj plik z glownego folderu do SubDirectory
                         * 
                         */
                        bool isDxfDateEmpty = string.IsNullOrEmpty(dxfDate);
                        if (isDxfDateEmpty)
                        {
                            document = CoreUtils.GetOpenDocument(assembly.Application, item.Key); isOpen = true;

                            if (document is SePart pDoc) { part = pDoc; models = part.Models; flatPatterns = part.FlatPatternModels; }
                            else if (document is SeSheetMetal msDoc) { metalSheet = msDoc; models = metalSheet.Models; flatPatterns = metalSheet.FlatPatternModels; }

                            if (flatPatterns == null || models == null || flatPatterns.Count == 0 || models.Count == 0) { logger.LogSkip(name, "Brak rozwinięcia"); continue; }

                            if (File.Exists(mainDxfPath)) { try { File.Delete(mainDxfPath); } catch { } }

                            models.SaveAsFlatDXFEx(mainDxfPath, null, null, null, true);

                            using var properties = new PropertyProvider(document); properties.UpdateDxfDate();

                            document.Close(true); isOpen = false;

                            logger.LogSuccess($"{name} (Utworzono nowy plik DXF)");
                        }
                        else { logger.LogSuccess($"{name} (Pobrano gotowy plik)"); }

                        if (File.Exists(mainDxfPath))
                        {
                            File.Copy(mainDxfPath, subDxfPath, true);

                            excelData[dataRowIndex, 0] = name;
                            excelData[dataRowIndex, 1] = $"{thickness} mm";
                            excelData[dataRowIndex, 2] = sizeX;
                            excelData[dataRowIndex, 3] = sizeY;
                            excelData[dataRowIndex, 4] = material;
                            excelData[dataRowIndex, 5] = count;
                            dataRowIndex++;
                        }
                    }
                    catch (Exception ex) { logger.LogError(name, $"Błąd procesu: {ex.Message}"); continue; }
                    finally
                    {
                        if (isOpen) document?.Close(true);
                        CoreUtils.ReleaseCom(ref flatPatterns);
                        CoreUtils.ReleaseCom(ref models);
                        CoreUtils.ReleaseCom(ref metalSheet);
                        CoreUtils.ReleaseCom(ref part);
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
                        startCell = (ExcelRange)worksheet.Cells[2, 1];
                        endCell = (ExcelRange)worksheet.Cells[dataRowIndex + 1, 6];
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
                usedRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                string excelPath = Path.Combine(subDirectory, "Zestawienie_DXF.xlsx");
                if (File.Exists(excelPath)) File.Delete(excelPath);
                workbook.SaveAs(excelPath);
            }
            finally
            {
                workbook?.Close(false);
                excelApp?.Quit();

                CoreUtils.ReleaseCom(ref columns);
                CoreUtils.ReleaseCom(ref usedRange);
                CoreUtils.ReleaseCom(ref headerRange);
                CoreUtils.ReleaseCom(ref worksheet);
                CoreUtils.ReleaseCom(ref xlSheets);
                CoreUtils.ReleaseCom(ref workbook);
                CoreUtils.ReleaseCom(ref workbooks);
                CoreUtils.ReleaseCom(ref excelApp);

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}