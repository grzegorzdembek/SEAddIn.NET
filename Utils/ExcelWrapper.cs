namespace SolidEdgeAdd_In.Utils
{
    public class ExcelWrapper
    {
        public static int GetColumnNumber(ExcelRange range, string name)
        {
            ExcelRange rows = null;
            ExcelRange firstRow = null;

            try
            {
                rows = range.Rows;
                firstRow = rows[1];
                object[,] rowData = (object[,])firstRow.Value2;

                if (rowData != null)
                {
                    int colCount = rowData.GetLength(1);

                    for (int i = 1; i <= colCount; i++)
                    {
                        if (rowData[1, i]?.ToString() == name)
                        {
                            return i;
                        }
                    }
                }

                return 0;
            }
            catch
            {
                return 0;
            }
            finally
            {
                CoreUtils.ReleaseCom(ref firstRow);
                CoreUtils.ReleaseCom(ref rows);
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
                CoreUtils.ReleaseCom(ref sheetMetalStyle);
                CoreUtils.ReleaseCom(ref normalStyle);
                CoreUtils.ReleaseCom(ref boughtStyle);
                CoreUtils.ReleaseCom(ref steelmakingStyle);
                CoreUtils.ReleaseCom(ref partStyle);
                CoreUtils.ReleaseCom(ref assemblyStyle);
                CoreUtils.ReleaseCom(ref styles);
            }
        }

        public static void TypeMemory(object[,] data, int typeNumber, int rowCount)
        {
            for (int i = 2; i <= rowCount; i++)
            {
                object rawValue = data[i, typeNumber];

                if (rawValue == null)
                {
                    continue;
                }

                string value = rawValue.ToString().Trim();

                if (value == Constants.PartTypes.Assembly)
                {
                    data[i, typeNumber] = Constants.Styles.Assembly;
                }
                else if (value == Constants.PartTypes.Part)
                {
                    data[i, typeNumber] = Constants.Styles.Part;
                }
                else if (value == Constants.PartTypes.Steelmaking)
                {
                    data[i, typeNumber] = Constants.Styles.Steelmaking;
                }
                else if (value == Constants.PartTypes.Commercial)
                {
                    data[i, typeNumber] = Constants.Styles.Commercial;
                }
                else if (value == Constants.PartTypes.Standard)
                {
                    data[i, typeNumber] = Constants.Styles.Standard;
                }
                else if (value == Constants.PartTypes.SheetMetal)
                {
                    data[i, typeNumber] = Constants.Styles.SheetMetal;
                }
            }
        }

        public static void Colors(ExcelWorksheet worksheet, int typeNumber)
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

                if (lastRow < 2)
                {
                    return;
                }

                dataRange = range.Range[range.Cells[2, 1], range.Cells[lastRow, lastCol]];
                columns = range.Columns;

                anchorCell = worksheet.Cells[2, typeNumber];
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
                CoreUtils.ReleaseCom(ref conditions);
                CoreUtils.ReleaseCom(ref anchorCell);
                CoreUtils.ReleaseCom(ref columns);
                CoreUtils.ReleaseCom(ref dataRange);
                CoreUtils.ReleaseCom(ref range);
            }
        }

        public static void CountMemory(object[,] data, int multiplier, int countNumber, int rowCount)
        {
            for (int i = 3; i <= rowCount; i++)
            {
                object rawValue = data[i, countNumber];

                if (rawValue != null && int.TryParse(rawValue.ToString(), out int currentQty))
                {
                    data[i, countNumber] = currentQty * multiplier;
                }
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

                int colCount = columns.Count;
                columns.AutoFit();
            }
            finally
            {
                CoreUtils.ReleaseCom(ref font);
                CoreUtils.ReleaseCom(ref interior);
                CoreUtils.ReleaseCom(ref borders);
                CoreUtils.ReleaseCom(ref cells);
                CoreUtils.ReleaseCom(ref firstRow);
                CoreUtils.ReleaseCom(ref columns);
                CoreUtils.ReleaseCom(ref rows);
                CoreUtils.ReleaseCom(ref range);
            }
        }

        public static string GetValue(ExcelRange range, int row, int column)
        {
            ExcelRange cells = null;
            string value = null;

            try
            {
                cells = range.Cells;
                object objectValue = cells[row, column].Value;

                if (objectValue != null)
                {
                    value = objectValue?.ToString()?.Trim();
                }

                return value;
            }
            catch
            {
                return null;
            }
            finally
            {
                CoreUtils.ReleaseCom(ref cells);
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
            finally
            {
                CoreUtils.ReleaseCom(ref interior);
                CoreUtils.ReleaseCom(ref rule);
            }
        }
    }
}