namespace SolidEdgeAdd_In.Utils
{
    public class ExcelUtils
    {
        public static int GetColumnIndex(ExcelRange range, string columnName)
        {
            ExcelRange rows = null;
            ExcelRange firstRow = null;

            try
            {
                rows = range.Rows;
                firstRow = rows[1];
                object[,] rowData = (object[,])firstRow.Value2;

                return Enumerable.Range(1, rowData.GetLength(1)).FirstOrDefault(i => rowData[1, i]?.ToString() == columnName);
            }
            catch { return 0; }
            finally { Helpers.ReleaseCom(ref firstRow); Helpers.ReleaseCom(ref rows); }
        }

        public static void ProcessDataInMemory(object[,] data, int typeColIdx, int fileNameColIdx, int countColIdx, int dxfColIdx, int rowCount, int multiplier, Dictionary<string, FileData> occurrencesData)
        {
            data[1, dxfColIdx] = Constants.Properties.DxfDate;

            for (int i = 2; i <= rowCount; i++)
            {
                if (data[i, typeColIdx] == null || data[i, fileNameColIdx] == null) continue;

                string rawType = data[i, typeColIdx].ToString().Trim();
                string fileName = data[i, fileNameColIdx].ToString().Trim();

                string mappedType = rawType;
                if (rawType == Constants.PartTypes.Assembly) mappedType = Constants.Styles.Assembly;
                else if (rawType == Constants.PartTypes.Part) mappedType = Constants.Styles.Part;
                else if (rawType == Constants.PartTypes.Steelmaking) mappedType = Constants.Styles.Steelmaking;
                else if (rawType == Constants.PartTypes.Commercial) mappedType = Constants.Styles.Commercial;
                else if (rawType == Constants.PartTypes.Standard) mappedType = Constants.Styles.Standard;
                else if (rawType == Constants.PartTypes.SheetMetal) mappedType = Constants.Styles.SheetMetal;

                data[i, typeColIdx] = mappedType;

                if (data[i, countColIdx] != null && int.TryParse(data[i, countColIdx].ToString(), out int currentQty))
                {
                    data[i, countColIdx] = currentQty * multiplier;
                }

                if (mappedType.Equals(Constants.Styles.SheetMetal, StringComparison.OrdinalIgnoreCase))
                {
                    FileData filedata = occurrencesData.Values.FirstOrDefault(f => f.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase));

                    if (filedata != null)
                    {
                        data[i, dxfColIdx] = filedata.DxfDate;
                    }
                    else
                    {
                        data[i, dxfColIdx] = "Missing DXF Property";
                    }
                }
                else
                {
                    data[i, dxfColIdx] = "-";
                }
            }
        }

        public static void Styles(ExcelWorkbook workbook)
        {
            ExcelStyles styles = null;
            ExcelStyle assemblyStyle = null;
            ExcelStyle partStyle = null;
            ExcelStyle steelmakingStyle = null;
            ExcelStyle boughtStyle = null;
            ExcelStyle normalStyle = null;
            ExcelStyle sheetMetalStyle = null;

            try
            {
                styles = workbook.Styles;

                assemblyStyle = styles.Add(Constants.Styles.Assembly);
                assemblyStyle.Interior.Color = ColorTranslator.ToOle(Color.Khaki);
                assemblyStyle.IncludeAlignment = false;
                assemblyStyle.IncludeBorder = false;

                partStyle = styles.Add(Constants.Styles.Part);
                partStyle.Interior.Color = ColorTranslator.ToOle(Color.Lavender);
                partStyle.IncludeAlignment = false;
                partStyle.IncludeBorder = false;

                steelmakingStyle = styles.Add(Constants.Styles.Steelmaking);
                steelmakingStyle.Interior.Color = ColorTranslator.ToOle(Color.SaddleBrown);
                steelmakingStyle.IncludeAlignment = false;
                steelmakingStyle.IncludeBorder = false;

                boughtStyle = styles.Add(Constants.Styles.Commercial);
                boughtStyle.Interior.Color = ColorTranslator.ToOle(Color.Aquamarine);
                boughtStyle.IncludeAlignment = false;
                boughtStyle.IncludeBorder = false;

                normalStyle = styles.Add(Constants.Styles.Standard);
                normalStyle.Interior.Color = ColorTranslator.ToOle(Color.CadetBlue);
                normalStyle.IncludeAlignment = false;
                normalStyle.IncludeBorder = false;

                sheetMetalStyle = styles.Add(Constants.Styles.SheetMetal);
                sheetMetalStyle.Interior.Color = ColorTranslator.ToOle(Color.Azure);
                sheetMetalStyle.IncludeAlignment = false;
                sheetMetalStyle.IncludeBorder = false;
            }
            finally
            {
                Helpers.ReleaseCom(ref sheetMetalStyle);
                Helpers.ReleaseCom(ref normalStyle);
                Helpers.ReleaseCom(ref boughtStyle);
                Helpers.ReleaseCom(ref steelmakingStyle);
                Helpers.ReleaseCom(ref partStyle);
                Helpers.ReleaseCom(ref assemblyStyle);
                Helpers.ReleaseCom(ref styles);
            }
        }

        public static void Colors(ExcelWorksheet worksheet, int typeColumnIndex)
        {
            ExcelRange range = null;
            ExcelRange dataRange = null;
            ExcelRange columns = null;
            ExcelRange anchorCell = null;
            ExcelFormatConditions conditions = null;

            try
            {
                range = worksheet.UsedRange;
                int lastRow = range.Rows.Count;
                int lastCol = range.Columns.Count;

                if (lastRow < 2) { return; }

                dataRange = range.Range[range.Cells[2, 1], range.Cells[lastRow, lastCol]];
                columns = range.Columns;

                anchorCell = worksheet.Cells[2, typeColumnIndex];
                string address = anchorCell.Address[RowAbsolute: false, ColumnAbsolute: true];

                conditions = dataRange.FormatConditions;
                conditions.Delete();

                Rule(conditions, address, Constants.Styles.Assembly, Color.Khaki);
                Rule(conditions, address, Constants.Styles.Part, Color.Lavender);
                Rule(conditions, address, Constants.Styles.Steelmaking, Color.SaddleBrown);
                Rule(conditions, address, Constants.Styles.Commercial, Color.Aquamarine);
                Rule(conditions, address, Constants.Styles.Standard, Color.CadetBlue);
                Rule(conditions, address, Constants.Styles.SheetMetal, Color.Azure);
            }
            finally
            {
                Helpers.ReleaseCom(ref conditions);
                Helpers.ReleaseCom(ref anchorCell);
                Helpers.ReleaseCom(ref columns);
                Helpers.ReleaseCom(ref dataRange);
                Helpers.ReleaseCom(ref range);
            }
        }

        public static void Edit(ExcelWorksheet worksheet)
        {
            ExcelRange range = null;
            ExcelRange rows = null;
            ExcelRange columns = null;
            ExcelRange firstRow = null;
            ExcelRange cells = null;
            dynamic borders = null;
            dynamic interior = null;
            dynamic font = null;

            try
            {
                range = worksheet.UsedRange;
                rows = range.Rows;
                columns = range.Columns;
                firstRow = rows[1];
                cells = range.Cells;

                range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                range.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

                borders = range.Borders;
                borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                borders.Color = (int)Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack;

                firstRow.RowHeight = 25;
                interior = firstRow.Interior;
                interior.Color = (int)Microsoft.Office.Interop.Excel.XlRgbColor.rgbLightGray;
                font = firstRow.Font;
                font.Bold = true;

                columns.AutoFit();
            }
            finally
            {
                Helpers.ReleaseCom(ref font); Helpers.ReleaseCom(ref interior);
                Helpers.ReleaseCom(ref borders); Helpers.ReleaseCom(ref cells);
                Helpers.ReleaseCom(ref firstRow); Helpers.ReleaseCom(ref columns);
                Helpers.ReleaseCom(ref rows); Helpers.ReleaseCom(ref range);
            }
        }

        private static void Rule(ExcelFormatConditions conditions, string address, string criteria, Color color)
        {
            ExcelFormatCondition rule = null;
            dynamic interior = null;
            try
            {
                string formula = $"={address}=\"{criteria}\"";
                rule = (ExcelFormatCondition)conditions.Add(Microsoft.Office.Interop.Excel.XlFormatConditionType.xlExpression, Formula1: formula);

                interior = rule.Interior;
                interior.Color = ColorTranslator.ToOle(color);
                rule.StopIfTrue = false;
            }
            finally { Helpers.ReleaseCom(ref interior); Helpers.ReleaseCom(ref rule); }
        }
    }
}