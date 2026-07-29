using SolidEdgeAdd_In.Utils;

namespace SolidEdgeAdd_In.Commands
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
                if (result.isConfirmed)
                { 
                    properties.Count = result.multiplier; 
                    return result; 
                }
                return (false, 1);
            }
            return (true, count);
        }

        public static Dictionary<string, FileData> GetData(SeAssembly assembly, Logger logger)
        {
            Dictionary<string, FileData> data = new(StringComparer.OrdinalIgnoreCase);
            SeOccurrences occurrences = null;
            try
            {
                occurrences = assembly.Occurrences;
                AssemblyTreeWalker.BuildDataForExportDxfs(occurrences, data, logger);
            }
            finally
            {
                CoreUtils.ReleaseCom(ref occurrences);
            }
            return data;
        }

        public static string GetSubDirectory(SeAssembly assembly)
        {
            string fileName = Path.GetFileNameWithoutExtension(assembly.FullName);
            string number = fileName.Length >= 4 ? fileName.Substring(0, 4) : fileName;
            string date = DateTime.Now.ToString("yyyy-MM-dd");
            string packagesDirectory = Path.Combine(Path.GetDirectoryName(assembly.FullName), Constants.Folders.Packages);
            if (!Directory.Exists(packagesDirectory)) Directory.CreateDirectory(packagesDirectory);
            string subDirectory = Path.Combine(packagesDirectory, $"{number}_{date}");
            if (!Directory.Exists(subDirectory)) Directory.CreateDirectory(subDirectory);
            return subDirectory;
        }

        public static void ProcessData(SeAssembly assembly, Dictionary<string, FileData> data, string subDirectory, int multiplier, Logger logger)
        {
            SeApp app = null;
            SeDocument document = null; 
            SeModels models = null; 
            SeModel model = null;
            SeBends bends = null;
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

            string mainDirectory = Path.GetDirectoryName(assembly.FullName);

            try
            {
                app = (SeApp)assembly.Application;

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

                object[,] headerData = new object[1, 8]
                { 
                    { 
                        Constants.ExcelHeaders.PartNumber, 
                        Constants.ExcelHeaders.Name, 
                        Constants.ExcelHeaders.Thickness, 
                        Constants.ExcelHeaders.Width, 
                        Constants.ExcelHeaders.Length, 
                        Constants.ExcelHeaders.Bending, 
                        Constants.ExcelHeaders.Material, 
                        Constants.ExcelHeaders.Quantity 
                    } 
                };

                try
                {
                    startHeaderCell = (ExcelRange)cells[1, 1]; 
                    endHeaderCell = (ExcelRange)cells[1, 8]; 
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

                object[,] excelData = new object[data.Count, 8]; 
                int dataRowIndex = 0;

                foreach (var item in data)
                {
                    /*
                     * values for path
                     */

                    string name = item.Value.Name; 
                    int count = item.Value.OccurrenceCount * multiplier;
                    string material = item.Value.Material; 
                    string thickness = item.Value.Thickness;

                    /*
                     * values for excel
                     */

                    string sizeX = item.Value.SizeX; 
                    string sizeY = item.Value.SizeY; 
                    string title = item.Value.Title;
                    string dxfDate = item.Value.DxfDate;
                    int bendsCount = 0;

                    /*
                     * different locations and filenames
                     */

                    string mainDxFilefName = $"{thickness}mm_{material}_{name}.dxf";
                    string subDxfFileName = $"{thickness}mm_{count}szt_{material}_{name}.dxf";

                    string mainDxfPath = Path.Combine(mainDirectory, mainDxFilefName);
                    string subDxfPath = Path.Combine(subDirectory, subDxfFileName);

                    /*
                     * conditions for generating DXF
                     */

                    bool needGenerationDxf = false;

                    if (string.IsNullOrEmpty(dxfDate) || !File.Exists(subDxfPath) || !File.Exists(mainDxfPath))
                    {
                        needGenerationDxf = true;
                    }

                    /*
                     * generation required open document
                     */

                    bool isOpen = false;
                    try
                    {
                        if (needGenerationDxf)
                        {
                            document = CoreUtils.GetOpenDocument(app, item.Key); 
                            isOpen = true;
                           
                            if (document is SePart part) 
                            { 
                                models = part.Models;
                                flatPatterns = part.FlatPatternModels;

                                model = models.Item(1);
                                bends = model.Bends;
                                bendsCount = bends.Count;
                            }
                            else if (document is SeSheetMetal sheetMetal) 
                            {
                                models = sheetMetal.Models;
                                flatPatterns = sheetMetal.FlatPatternModels;

                                model = models.Item(1);
                                bends = model.Bends;
                                bendsCount = bends.Count;
                            }

                            if (flatPatterns == null || models == null || flatPatterns.Count == 0 || models.Count == 0)
                            { 
                                logger.LogSkip(name, "Brak rozwinięcia"); 
                                continue; 
                            }

                            if (File.Exists(mainDxfPath))
                            { 
                                try 
                                { 
                                    File.Delete(mainDxfPath); 
                                } 
                                catch 
                                {
                                    /* */
                                } 
                            }

                            using var properties = new PropertyProvider(document); properties.UpdateDxfDate(); 

                            models.SaveAsFlatDXFEx(mainDxfPath, null, null, null, true); 

                            logger.LogSuccess($"{name} Utworzono nowy plik DXF");
                        }
                        else 
                        { 
                            logger.LogSuccess($"{name} Plik DXF już istnieje lub posiada właściwość DXF"); 
                        }

                        if (File.Exists(mainDxfPath))
                        {
                            File.Copy(mainDxfPath, subDxfPath, true);

                            excelData[dataRowIndex, 0] = name; 
                            excelData[dataRowIndex, 1] = title;
                            excelData[dataRowIndex, 2] = $"{thickness} mm";
                            excelData[dataRowIndex, 3] = sizeX; 
                            excelData[dataRowIndex, 4] = sizeY; 
                            excelData[dataRowIndex, 5] = bendsCount;
                            excelData[dataRowIndex, 6] = material;
                            excelData[dataRowIndex, 7] = count;
                            dataRowIndex++;
                        }
                    }
                    catch (Exception ex)
                    { 
                        logger.LogError(name, $"Błąd procesu: {ex.Message}"); 
                        continue; 
                    }
                    finally
                    {
                        CoreUtils.ReleaseCom(ref bends);
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
                            catch 
                            {
                                /* */
                            }
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
                        endCell = (ExcelRange)cells[dataRowIndex + 1, 6]; 
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

                string excelPath = Path.Combine(subDirectory, "Zestawienie_DXF.xlsx");

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
                    try 
                    { 
                        workbook.Close(false); 
                    } 
                    catch 
                    {
                        /* */
                    } 
                }

                CoreUtils.ReleaseCom(ref workbook); 
                CoreUtils.ReleaseCom(ref workbooks);

                if (excelApp != null) 
                { 
                    try 
                    { 
                        excelApp.Quit(); 
                    } 
                    catch 
                    {
                        /* */
                    }
                }

                CoreUtils.ReleaseCom(ref excelApp);
            }
        }
    }
}