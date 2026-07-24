namespace SolidEdgeAdd_In.Utils
{
    public class RaportGenerationUtils
    {
        public static void DxfsMemory(string location, object[,] data, int typeNumber, int nameNumber, int checkDxfColumn, int rowCount)
        {
            data[1, checkDxfColumn] = "DXF:";

            string[] allFiles = Directory.GetFiles(location, "*.*", SearchOption.TopDirectoryOnly); var allDxfFiles = allFiles.Where(f => f.EndsWith(".dxf", StringComparison.OrdinalIgnoreCase)).ToList(); var allCadFiles = allFiles.Where(f => f.EndsWith(".par", StringComparison.OrdinalIgnoreCase) || f.EndsWith(".psm", StringComparison.OrdinalIgnoreCase)).ToList();

            var dxfFileNames = allDxfFiles.Select(f => new { Path = f, Name = Path.GetFileName(f) }).ToList(); var cadFileNames = allCadFiles.Select(f => new { Path = f, Name = Path.GetFileName(f) }).ToList();

            for (int i = 2; i <= rowCount; i++)
            {
                object rawName = data[i, nameNumber]; object rawType = data[i, typeNumber];

                if (rawName == null || rawType == null) continue;

                string name = rawName.ToString().Trim(); string type = rawType.ToString().Trim();

                if (type.Equals("Blacha", StringComparison.OrdinalIgnoreCase))
                {
                    bool hasDxf = dxfFileNames.Any(f => f.Name.IndexOf(name, StringComparison.OrdinalIgnoreCase) >= 0);

                    if (hasDxf)
                    {
                        string cadFilePath = allCadFiles.FirstOrDefault(f => Path.GetFileName(f).IndexOf(name, StringComparison.OrdinalIgnoreCase) >= 0);

                        if (!string.IsNullOrEmpty(cadFilePath))
                        {
                            string dxfDate = null; using (var properties = new PropertyProvider(cadFilePath, true)) { dxfDate = properties.DxfDate; }

                            if (!string.IsNullOrEmpty(dxfDate)) data[i, checkDxfColumn] = dxfDate; else data[i, checkDxfColumn] = "Brak właściwości DXF";
                        }
                    }
                    else { data[i, checkDxfColumn] = "Brak DXF"; }
                }
                else { data[i, checkDxfColumn] = "-"; }
            }
        }

        public static void Shots(ExcelWorksheet worksheet, List<string> shotPaths, bool hasShots, string shotFolder, int typeNumber, int nameNumber, int imageNumber)
        {
            ExcelShapes shapes = null; ExcelRange range = null; ExcelRange rows = null; ExcelRange columns = null; ExcelRange firstRow = null;

            try
            {
                worksheet.Activate(); shapes = worksheet.Shapes; range = worksheet.UsedRange;

                rows = range.Rows; columns = range.Columns; firstRow = rows[1];

                List<int> numberRowToDelete = new List<int>(); int rowCount = rows.Count;

                Dictionary<string, string> shotDict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                if (!hasShots && shotPaths.Count > 0) { foreach (string sp in shotPaths) shotDict[Path.GetFileNameWithoutExtension(sp)] = sp; }
                else if (hasShots && shotPaths.Count == 0) { string[] folderShots = Directory.GetFiles(shotFolder, "*.jpg", SearchOption.TopDirectoryOnly); foreach (string sp in folderShots) shotDict[Path.GetFileNameWithoutExtension(sp)] = sp; }

                for (int i = 1; i <= rowCount; i++)
                {
                    ExcelRange row = null; ExcelRange cells = null; ExcelRange shotCell = null; ExcelApp excelApp = null; Microsoft.Office.Core.CommandBars cmdBars = null;

                    try
                    {
                        row = (ExcelRange)rows[i]; if (firstRow.Row == row.Row) continue;

                        string name = ExcelWrapper.GetValue(range, row.Row, nameNumber); if (name == null) continue;

                        cells = row.Cells; shotCell = (ExcelRange)cells[1, imageNumber]; shotCell.RowHeight = 120; shotCell.ColumnWidth = 20;

                        string type = ExcelWrapper.GetValue(range, row.Row, typeNumber); if (type == null) { numberRowToDelete.Add(row.Row); continue; }

                        if (shotDict.TryGetValue(name, out string matchedShotPath))
                        {
                            var picture = shapes.AddPicture(matchedShotPath, MsoTriState.msoFalse, MsoTriState.msoCTrue, (float)shotCell.Left, (float)shotCell.Top, (float)shotCell.Width, (float)shotCell.Height);

                            picture.Placement = ExcelXlPlacement.xlMoveAndSize; picture.Select();

                            excelApp = worksheet.Application; cmdBars = excelApp.CommandBars; cmdBars.ExecuteMso("PicturePlaceInCell");

                            CoreUtils.ReleaseCom(ref picture);
                        }
                    }
                    finally { CoreUtils.ReleaseCom(ref cmdBars); CoreUtils.ReleaseCom(ref excelApp); CoreUtils.ReleaseCom(ref shotCell); CoreUtils.ReleaseCom(ref cells); CoreUtils.ReleaseCom(ref row); }
                }

                var sortedIndices = numberRowToDelete.OrderByDescending(i => i).ToList();
                foreach (int rowIndex in sortedIndices)
                {
                    ExcelRange row = null; try { row = (ExcelRange)rows[rowIndex]; row.Delete(ExcelXlDeleteShiftDirection.xlShiftUp); } finally { CoreUtils.ReleaseCom(ref row); }
                }
            }
            finally { CoreUtils.ReleaseCom(ref firstRow); CoreUtils.ReleaseCom(ref columns); CoreUtils.ReleaseCom(ref rows); CoreUtils.ReleaseCom(ref range); CoreUtils.ReleaseCom(ref shapes); }
        }

        public static string GetShotPath(string path, SeWindow window)
        {
            string shotPath = path + ".jpg"; SeView view = null;

            try
            {
                view = window.View; view.Update(); view.Fit();

                view.SaveAsImage(Filename: shotPath, Width: window.UsableWidth, Height: window.UsableHeight, AltViewStyle: null, Resolution: 1, ColorDepth: 24, ImageQuality: SeImageQualityType.seImageQualityHigh, Invert: false);

                return shotPath;
            }
            finally { CoreUtils.ReleaseCom(ref view); }
        }
    }
}