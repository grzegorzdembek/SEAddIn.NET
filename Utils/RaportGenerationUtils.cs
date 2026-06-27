using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel; 

namespace SolidEdgeAdd_In.Utils
{
    public class RaportGenerationUtils
    {
        public static void Dxfs(string location, Excel.Worksheet worksheet, int typeNumber, int nameNumber)
        {
            Excel.Range range = null;
            Excel.Range rows = null;
            Excel.Range columns = null;
            Excel.Range firstRow = null;
            Excel.Range cells = null;
            try
            {
                range = worksheet.UsedRange;
                rows = worksheet.UsedRange.Rows;
                columns = worksheet.UsedRange.Columns;
                firstRow = rows[1];
                cells = range.Cells;

                int checkDxfColumn = columns.Count + 1;
                cells[1, checkDxfColumn] = "DXF:";

                foreach (Excel.Range row in rows)
                {
                    if (row.Row == firstRow.Row) continue;

                    string name = ExcelWrapper.GetValue(range, row.Row, nameNumber);
                    if (name == null) continue;

                    string type = ExcelWrapper.GetValue(range, row.Row, typeNumber);
                    if (type == null) continue;

                    if (type.Equals("Blacha", StringComparison.OrdinalIgnoreCase))
                    {
                        string[] dxfFiles = Directory.GetFiles(location, $"*{name}*.dxf", SearchOption.TopDirectoryOnly);

                        if (dxfFiles.Length > 0)
                        {
                            var cadFiles = Directory.GetFiles(location, $"*{name}*.*", SearchOption.TopDirectoryOnly)
                                                    .Where(f => f.EndsWith(".par", StringComparison.OrdinalIgnoreCase) ||
                                                                f.EndsWith(".psm", StringComparison.OrdinalIgnoreCase))
                                                    .ToArray();

                            if (cadFiles.Length > 0)
                            {
                                string cadFilePath = cadFiles[0];
                                string dxfDate = PropertyProvider.GetDxfDate(cadFilePath);

                                if (!string.IsNullOrEmpty(dxfDate))
                                {
                                    cells[row.Row, checkDxfColumn].Value = dxfDate;
                                }
                                else
                                {
                                    cells[row.Row, checkDxfColumn].Value = "Brak właściwości DXF"; 
                                }
                            }
                        }
                        else
                        {
                            cells[row.Row, checkDxfColumn].Value = "Brak DXF"; 
                        }
                    }
                    else
                    {
                        cells[row.Row, checkDxfColumn].Value = "-"; 
                    }
                }
            }
            finally
            {
                CoreUtils.ReleaseCom(ref cells);
                CoreUtils.ReleaseCom(ref firstRow);
                CoreUtils.ReleaseCom(ref columns);
                CoreUtils.ReleaseCom(ref rows);
                CoreUtils.ReleaseCom(ref range);
            }
        }

        public static void Shots(Excel.Worksheet worksheet, List<string> shotPaths, bool hasShots, string shotFolder, int typeNumber, int nameNumber, int imageNumber)
        {
            Excel.Shapes shapes = null;
            Excel.Range range = null;
            Excel.Range rows = null;
            Excel.Range columns = null;
            Excel.Range firstRow = null;
            try
            {
                worksheet.Activate();
                shapes = worksheet.Shapes;
                range = worksheet.UsedRange;
                rows = worksheet.UsedRange.Rows;
                columns = worksheet.UsedRange.Columns;
                firstRow = rows[1];           

                List<int> numberRowToDelete = new List<int>();

                if (!hasShots && shotPaths.Count > 0)
                {
                    foreach (Excel.Range row in rows)
                    {
                        if (firstRow.Row == row.Row) continue;

                        string name = ExcelWrapper.GetValue(range, row.Row, nameNumber);
                        if (name == null) continue;

                        dynamic shotCell = row.Cells[1, imageNumber];
                        shotCell.RowHeight = 120;
                        shotCell.ColumnWidth = 20;

                        string type = ExcelWrapper.GetValue(range, row.Row, typeNumber);
                        if (type == null)
                        {
                            numberRowToDelete.Add(row.Row);
                            continue;
                        }

                        foreach (string shotPath in shotPaths)
                        {
                            string jpgName = Path.GetFileNameWithoutExtension(shotPath);
                            if (string.Equals(name, jpgName, StringComparison.OrdinalIgnoreCase))
                            {
                                var picture = shapes.AddPicture(shotPath,
                                                            Microsoft.Office.Core.MsoTriState.msoFalse,
                                                            Microsoft.Office.Core.MsoTriState.msoCTrue,
                                                            (float)shotCell.Left,
                                                            (float)shotCell.Top,
                                                            (float)shotCell.Width,
                                                            (float)shotCell.Height);

                                picture.Placement = Excel.XlPlacement.xlMoveAndSize;
                                picture.Select();
                                worksheet.Application.CommandBars.ExecuteMso("PicturePlaceInCell");

                                CoreUtils.ReleaseCom(ref picture);
                                break;
                            }
                        }
                    }
                }

                if (hasShots && shotPaths.Count == 0)
                {
                    List<string> shotPathsFromFolder = Directory.GetFiles(shotFolder, "*.jpg", SearchOption.TopDirectoryOnly).ToList();
                    if (shotPathsFromFolder.Count == 0) return;

                    foreach (Excel.Range row in rows)
                    {
                        if (firstRow.Row == row.Row) continue;

                        string name = ExcelWrapper.GetValue(range, row.Row, nameNumber);
                        if (name == null) continue;

                        dynamic shotCell = row.Cells[1, imageNumber];

                        shotCell.RowHeight = 120;
                        shotCell.ColumnWidth = 20;

                        string type = ExcelWrapper.GetValue(range, row.Row, typeNumber);
                        if (type == null)
                        {
                            numberRowToDelete.Add(row.Row);
                            continue;
                        }

                        foreach (string shotPath in shotPathsFromFolder)
                        {
                            string shotName = Path.GetFileNameWithoutExtension(shotPath);
                            if (string.Equals(name, shotName, StringComparison.OrdinalIgnoreCase))
                            {
                                var picture = shapes.AddPicture(shotPath,
                                                            Microsoft.Office.Core.MsoTriState.msoFalse,
                                                            Microsoft.Office.Core.MsoTriState.msoCTrue,
                                                            (float)shotCell.Left,
                                                            (float)shotCell.Top,
                                                            (float)shotCell.Width,
                                                            (float)shotCell.Height);

                                picture.Placement = Excel.XlPlacement.xlMoveAndSize;
                                picture.Select();
                                worksheet.Application.CommandBars.ExecuteMso("PicturePlaceInCell");

                                CoreUtils.ReleaseCom(ref picture);
                                break;
                            }
                        }                     
                    }
                }

                var sortedIndices = numberRowToDelete.OrderByDescending(i => i).ToList();
                foreach (int rowIndex in sortedIndices)
                {
                    Microsoft.Office.Interop.Excel.Range row = rows[rowIndex];
                    row.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                }
            }
            finally
            {
                CoreUtils.ReleaseCom(ref firstRow);
                CoreUtils.ReleaseCom(ref columns);
                CoreUtils.ReleaseCom(ref rows);
                CoreUtils.ReleaseCom(ref range);
                CoreUtils.ReleaseCom(ref shapes);
            }
        }

        public static string GetDxfPath
           (string folder, string filePath, Dictionary<string, int> dict)
        {
            string baseName = Path.GetFileNameWithoutExtension(filePath);
            string dxfFileName =
            $"{PropertyProvider.GetThickness(filePath)}mm_{CoreUtils.GetCount(dict, filePath)}szt_{PropertyProvider.GetMaterial(filePath)}_{baseName}.dxf";
            string dxfFilePath = Path.Combine(folder, dxfFileName);
            return dxfFilePath;
        }

        public static string GetShotPath
            (string path, SolidEdgeFramework.Window window)
        {
            string shotPath = path + ".jpg";
            SolidEdgeFramework.View view = null;
            try
            {
                view = window.View;
                view.Update();
                view.Fit();
                view.SaveAsImage(Filename: shotPath,
                    Width: window.UsableWidth,
                    Height: window.UsableHeight,
                    AltViewStyle: null, Resolution: 1, ColorDepth: 24,
                    ImageQuality: SolidEdgeFramework.SeImageQualityType.seImageQualityHigh,
                    Invert: false);
                return shotPath;
            }
            finally { CoreUtils.ReleaseCom(ref view); }
        }
    }
}