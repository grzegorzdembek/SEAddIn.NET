namespace SolidEdgeAdd_In.Utils
{
    public class RaportUtils
    {
        public static void DxfsMemory(string directoryPath, object[,] data, int typeColumnIndex, int fileNameColumnIndex, int dxfColumnIndex, int rowCount)
        {
            data[1, dxfColumnIndex] = Constants.Properties.DxfDate;

            string[] allFiles = Directory.GetFiles(directoryPath, "*.*", SearchOption.TopDirectoryOnly);
            var allDxfFiles = allFiles.Where(f => f.EndsWith(".dxf", StringComparison.OrdinalIgnoreCase)).ToList();
            var allCadFiles = allFiles.Where(f => f.EndsWith(".par", StringComparison.OrdinalIgnoreCase) || f.EndsWith(".psm", StringComparison.OrdinalIgnoreCase)).ToList();

            var dxfFileNames = allDxfFiles.Select(f => new { Path = f, FileName = Path.GetFileName(f) }).ToList();
            var cadFileNames = allCadFiles.Select(f => new { Path = f, FileName = Path.GetFileName(f) }).ToList();

            for (int i = 2; i <= rowCount; i++)
            {
                object rawFileName = data[i, fileNameColumnIndex];
                object rawType = data[i, typeColumnIndex];

                if (rawFileName == null || rawType == null) { continue; }

                string fileName = rawFileName.ToString().Trim();
                string type = rawType.ToString().Trim();

                if (type.Equals(Constants.Styles.SheetMetal, StringComparison.OrdinalIgnoreCase))
                {
                    bool hasDxf = dxfFileNames.Any(f => f.FileName.IndexOf(fileName, StringComparison.OrdinalIgnoreCase) >= 0);

                    if (hasDxf)
                    {
                        string cadFilePath = allCadFiles.FirstOrDefault(f => Path.GetFileName(f).IndexOf(fileName, StringComparison.OrdinalIgnoreCase) >= 0);

                        if (!string.IsNullOrEmpty(cadFilePath))
                        {
                            string dxfDate = null;
                            using (var properties = new PropertyUtils(cadFilePath, true)) { dxfDate = properties.DxfDate; }

                            if (!string.IsNullOrEmpty(dxfDate)) { data[i, dxfColumnIndex] = dxfDate; }
                            else { data[i, dxfColumnIndex] = "Missing DXF Property"; }
                        }
                    }
                    else { data[i, dxfColumnIndex] = "Missing DXF"; }
                }
                else { data[i, dxfColumnIndex] = "-"; }
            }
        }

        public static void Shots(ExcelWorksheet worksheet, List<string> shotPaths, bool hasShots, string shotsDirectoryPath, int typeColumnIndex, int fileNameColumnIndex, int thumbnailColumnIndex)
        {
            ExcelShapes shapes = null;
            ExcelRange range = null;
            ExcelRange rows = null;
            ExcelRange columns = null;
            ExcelRange firstRow = null;

            try
            {
                worksheet.Activate();
                shapes = worksheet.Shapes;
                range = worksheet.UsedRange;

                rows = range.Rows;
                columns = range.Columns;
                firstRow = rows[1];

                List<int> numberRowToDelete = new();
                int rowCount = rows.Count;

                Dictionary<string, string> shotDict = new(StringComparer.OrdinalIgnoreCase);

                if (!hasShots && shotPaths.Count > 0)
                {
                    foreach (string sp in shotPaths) { shotDict[Path.GetFileNameWithoutExtension(sp)] = sp; }
                }
                else if (hasShots && shotPaths.Count == 0)
                {
                    string[] folderShots = Directory.GetFiles(shotsDirectoryPath, "*.jpg", SearchOption.TopDirectoryOnly);
                    foreach (string sp in folderShots) { shotDict[Path.GetFileNameWithoutExtension(sp)] = sp; }
                }

                for (int i = 1; i <= rowCount; i++)
                {
                    ExcelRange row = null;
                    ExcelRange cells = null;
                    ExcelRange shotCell = null;
                    ExcelApp excelApp = null;
                    Microsoft.Office.Core.CommandBars cmdBars = null;

                    try
                    {
                        row = (ExcelRange)rows[i];
                        if (firstRow.Row == row.Row) { continue; }

                        string fileName = ExcelUtils.GetValue(range, row.Row, fileNameColumnIndex);
                        if (fileName == null) { continue; }

                        cells = row.Cells;
                        shotCell = (ExcelRange)cells[1, thumbnailColumnIndex];
                        shotCell.RowHeight = 120;
                        shotCell.ColumnWidth = 20;

                        string type = ExcelUtils.GetValue(range, row.Row, typeColumnIndex);
                        if (type == null) { numberRowToDelete.Add(row.Row); continue; }

                        if (shotDict.TryGetValue(fileName, out string matchedShotPath))
                        {
                            var picture = shapes.AddPicture(matchedShotPath, MsoTriState.msoFalse, MsoTriState.msoCTrue, (float)shotCell.Left + 1f, (float)shotCell.Top + 1f, (float)shotCell.Width - 2f, (float)shotCell.Height - 2f);
                            picture.Placement = ExcelXlPlacement.xlMoveAndSize;
                            picture.Select();

                            excelApp = worksheet.Application;
                            cmdBars = excelApp.CommandBars;
                            cmdBars.ExecuteMso("PicturePlaceInCell");

                            Helpers.ReleaseCom(ref picture);
                        }
                    }
                    finally
                    {
                        Helpers.ReleaseCom(ref cmdBars); Helpers.ReleaseCom(ref excelApp);
                        Helpers.ReleaseCom(ref shotCell); Helpers.ReleaseCom(ref cells); Helpers.ReleaseCom(ref row);
                    }
                }

                var sortedIndices = numberRowToDelete.OrderByDescending(i => i).ToList();
                foreach (int rowIndex in sortedIndices)
                {
                    ExcelRange row = null;
                    try { row = (ExcelRange)rows[rowIndex]; row.Delete(ExcelXlDeleteShiftDirection.xlShiftUp); }
                    finally { Helpers.ReleaseCom(ref row); }
                }
            }
            finally
            {
                Helpers.ReleaseCom(ref firstRow); Helpers.ReleaseCom(ref columns);
                Helpers.ReleaseCom(ref rows); Helpers.ReleaseCom(ref range); Helpers.ReleaseCom(ref shapes);
            }
        }

        public static string GetShotPath(string filePath, SeWindow window)
        {
            string shotPath = filePath + ".jpg";
            SeView view = null;

            try
            {
                view = window.View;
                view.Update();
                view.Fit();
                view.SaveAsImage(Filename: shotPath, Width: window.UsableWidth, Height: window.UsableHeight, AltViewStyle: null, Resolution: 1, ColorDepth: 24, ImageQuality: SeImageQualityType.seImageQualityHigh, Invert: false);
                return shotPath;
            }
            finally { Helpers.ReleaseCom(ref view); }
        }
    }
}