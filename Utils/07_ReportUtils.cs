namespace SolidEdgeAdd_In.Utils
{
    public class ReportUtils
    {
        public static void SaveThumbnail(string thumbnailPath, SeWindow window)
        {
            SeView view = null;
            try
            {
                view = window.View;
                view.Update();
                view.Fit();
                view.SaveAsImage(Filename: thumbnailPath,
                    Width: window.UsableWidth,
                    Height: window.UsableHeight,
                    AltViewStyle: null,
                    Resolution: 1,
                    ColorDepth: 24,
                    ImageQuality: SeImageQualityType.seImageQualityHigh,
                    Invert: false);
            }
            finally { Helpers.ReleaseCom(ref view); }
        }

        public static void CopyPartsList(SeApp application, string assemblyFilePath)
        {
            SeDocuments documents = null;
            SeDraft draft = null;
            SeDraftSheet sheet = null;
            SeDrawingViews drawingViews = null;
            SeDrawingView drawingView = null;
            SeModelLinks modelLinks = null;
            SeModelLink modelLink = null;
            SePartsLists partsLists = null;
            SePartsList partsList = null;

            try
            {
                documents = application.Documents;
                draft = (SeDraft)documents.Add("SolidEdge.DraftDocument", Missing.Value);
                sheet = draft.ActiveSheet;

                modelLinks = draft.ModelLinks; modelLink = modelLinks.Add(assemblyFilePath);

                drawingViews = sheet.DrawingViews; drawingView = drawingViews.AddAssemblyView(modelLink, SeViewOrientation.igFrontView, 0.1, 0.2, 0.2, SeAssemblyDrawingViewType.seAssemblyDesignedView);

                partsLists = draft.PartsLists; partsList = partsLists.AddEx(drawingView, 0, "", 0, 1);

                Array listOfSavedSettings = Array.CreateInstance(typeof(object), 0);
                partsList.GetListOfSavedSettings(out int numSavedSettings, ref listOfSavedSettings);

                Helpers.ReleaseCom(ref partsList); Helpers.ReleaseCom(ref partsLists);

                List<string> settingsList = new ();

                if (listOfSavedSettings != null)
                {
                    foreach (var o in listOfSavedSettings)
                    {
                        if (o != null) { settingsList.Add(o.ToString()); }
                    }
                }

                if (settingsList.Count == 0) { settingsList.Add("<No saved parts list styles available>"); }

                string partListType = DialogUtils.GetPartsListType(settingsList);

                partsLists = draft.PartsLists;
                partsList = partsLists.AddEx(drawingView, 0, partListType, 0, 1);

                for (int i = 0; i < 5; i++)
                {
                    try { partsList.CopyToClipboard(); System.Threading.Thread.Sleep(300); break; }
                    catch { System.Threading.Thread.Sleep(300); }
                }
            }
            finally
            {
                Helpers.ReleaseCom(ref partsList); Helpers.ReleaseCom(ref partsLists);
                Helpers.ReleaseCom(ref modelLink); Helpers.ReleaseCom(ref modelLinks);
                Helpers.ReleaseCom(ref drawingView); Helpers.ReleaseCom(ref drawingViews);
                Helpers.ReleaseCom(ref sheet);

                try { draft?.Close(false); } catch { }

                Helpers.ReleaseCom(ref draft); Helpers.ReleaseCom(ref documents);
            }
        }

        public static void InsertThumbnailsOnly(ExcelWorksheet worksheet, Dictionary<string, string> thumbnailPaths, int fileNameColIdx, int thumbnailColIdx)
        {
            ExcelShapes shapes = null;
            ExcelRange range = null;

            try
            {
                worksheet.Activate();
                shapes = worksheet.Shapes;
                range = worksheet.UsedRange;

                // Pojedynczy strzał do pamięci
                object[,] data = (object[,])range.Value2;
                int rowCount = data.GetLength(0);

                // Zaczynamy od 2, aby pominąć nagłówek
                for (int i = 2; i <= rowCount; i++)
                {
                    ExcelRange cell = null;
                    ExcelApp excelApp = null;
                    Microsoft.Office.Core.CommandBars cmdBars = null;

                    try
                    {
                        object fileNameObj = data[i, fileNameColIdx];
                        if (fileNameObj == null) continue;

                        string fileName = fileNameObj.ToString().Trim();
                        if (string.IsNullOrEmpty(fileName)) continue;

                        if (thumbnailPaths.TryGetValue(fileName, out string matchedShotPath))
                        {
                            cell = (ExcelRange)worksheet.Cells[i, thumbnailColIdx];

                            cell.RowHeight = 120;
                            cell.ColumnWidth = 20;

                            var picture = shapes.AddPicture(
                                matchedShotPath,
                                Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoCTrue,
                                (float)cell.Left + 1f,
                                (float)cell.Top + 1f,
                                (float)cell.Width - 2f,
                                (float)cell.Height - 2f);

                            picture.Placement = Microsoft.Office.Interop.Excel.XlPlacement.xlMoveAndSize;

                            // Bezpieczne wstawienie w komórkę
                            try
                            {
                                picture.Select();
                                excelApp = worksheet.Application;
                                cmdBars = excelApp.CommandBars;
                                cmdBars.ExecuteMso("PicturePlaceInCell");
                            }
                            catch
                            {
                                // Ignorujemy wyjątek w starszych wersjach Excela
                            }
                            finally { Helpers.ReleaseCom(ref picture); }
                        }
                    }
                    finally
                    {
                        Helpers.ReleaseCom(ref cmdBars);
                        Helpers.ReleaseCom(ref excelApp);
                        Helpers.ReleaseCom(ref cell);
                    }
                }
            }
            finally
            {
                Helpers.ReleaseCom(ref range);
                Helpers.ReleaseCom(ref shapes);
            }
        }

        public static void SetThumbnails(ExcelWorksheet worksheet, Dictionary<string, string> thumbnailPaths, int typeColIdx, int fileNameColIdx, int thumbnailColIdx)
        {
            ExcelShapes shapes = null;
            ExcelRange range = null;
            ExcelRange rows = null;

            try
            {
                worksheet.Activate();
                shapes = worksheet.Shapes;

                range = worksheet.UsedRange;
                object[,] initialData = (object[,])range.Value2;
                int initialRowCount = initialData.GetLength(0);

                if (initialRowCount < 2) return;

                List<int> rowsToDelete = new ();

                for (int i = 2; i <= initialRowCount; i++)
                {
                    object typeObj = initialData[i, typeColIdx];

                    if (typeObj == null || string.IsNullOrWhiteSpace(typeObj.ToString()))
                    {
                        rowsToDelete.Add(i);
                    }
                }

                rowsToDelete.Reverse();
                rows = range.Rows;

                foreach (int rowIndex in rowsToDelete)
                {
                    ExcelRange rowToDelete = null;
                    try
                    {
                        rowToDelete = (ExcelRange)rows[rowIndex];
                        rowToDelete.Delete(Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftUp);
                    }
                    finally { Helpers.ReleaseCom(ref rowToDelete); }
                }

                Helpers.ReleaseCom(ref rows);
                Helpers.ReleaseCom(ref range);

                range = worksheet.UsedRange;
                object[,] finalData = (object[,])range.Value2;
                int finalRowCount = finalData.GetLength(0);

                for (int i = 2; i <= finalRowCount; i++)
                {
                    ExcelRange cell = null;
                    ExcelApp excelApp = null;
                    Microsoft.Office.Core.CommandBars cmdBars = null;

                    try
                    {
                        object fileNameObj = finalData[i, fileNameColIdx];
                        if (fileNameObj == null) continue;

                        string fileName = fileNameObj.ToString().Trim();
                        if (string.IsNullOrEmpty(fileName)) continue;

                        if (thumbnailPaths.TryGetValue(fileName, out string matchedShotPath))
                        {
                            cell = (ExcelRange)worksheet.Cells[i, thumbnailColIdx];

                            cell.RowHeight = 120;
                            cell.ColumnWidth = 20;

                            var picture = shapes.AddPicture(
                                matchedShotPath,
                                Microsoft.Office.Core.MsoTriState.msoFalse,
                                Microsoft.Office.Core.MsoTriState.msoCTrue,
                                (float)cell.Left + 1f,
                                (float)cell.Top + 1f,
                                (float)cell.Width - 2f,
                                (float)cell.Height - 2f);

                            picture.Placement = Microsoft.Office.Interop.Excel.XlPlacement.xlMoveAndSize;

                            try
                            {
                                picture.Select();
                                excelApp = worksheet.Application;
                                cmdBars = excelApp.CommandBars;
                                cmdBars.ExecuteMso("PicturePlaceInCell");
                            }
                            catch
                            {
                                continue;
                            }
                            finally { Helpers.ReleaseCom(ref picture); }
                        }
                    }
                    finally
                    {
                        Helpers.ReleaseCom(ref cmdBars);
                        Helpers.ReleaseCom(ref excelApp);
                        Helpers.ReleaseCom(ref cell);
                    }
                }
            }
            finally
            {
                Helpers.ReleaseCom(ref rows);
                Helpers.ReleaseCom(ref range);
                Helpers.ReleaseCom(ref shapes);
            }
        }
    }
}