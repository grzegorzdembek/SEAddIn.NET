using Excel = Microsoft.Office.Interop.Excel; 
using System.Drawing;
using System;

namespace SolidEdgeAdd_In.Utils
{
    public class ExcelWrapper
    {
        public static int GetColumnNumber(Excel.Range range, string name)
        {
            Excel.Range firstRow = null;
            Excel.Range cells = null;
            try
            {
                firstRow = range.Rows[1];
                cells = firstRow.Cells;

                foreach (Excel.Range cell in cells)
                {
                    string value = cell.Value?.ToString();

                    if (value == name)
                    {
                        int index = cell.Column;
                        return index;
                    }
                }
                return 0;
            }
            catch { return 0; }
            finally
            {
                CoreUtils.ReleaseCom(ref cells);
                CoreUtils.ReleaseCom(ref firstRow);
            }
        }

        public static void Styles(Excel.Workbook workbook)
        {
            Excel.Styles styles = null;
            Excel.Style assemblyStyle = null;
            Excel.Style partStyle = null;
            Excel.Style steelmakingStyle = null;
            Excel.Style boughtStyle = null;
            Excel.Style normalStyle = null;
            Excel.Style sheetMetalStyle = null;

            try
            {
                styles = workbook.Styles;

                assemblyStyle = styles.Add("Złożenie");
                assemblyStyle.Interior.Color = ColorTranslator.ToOle(Color.Khaki);
                assemblyStyle.IncludeAlignment = false;
                assemblyStyle.IncludeBorder = false;

                partStyle = styles.Add("Część");
                partStyle.Interior.Color = ColorTranslator.ToOle(Color.Lavender);
                partStyle.IncludeAlignment = false;
                partStyle.IncludeBorder = false;

                steelmakingStyle = styles.Add("Hutnicze");
                steelmakingStyle.Interior.Color = ColorTranslator.ToOle(Color.SaddleBrown);
                steelmakingStyle.IncludeAlignment = false;
                steelmakingStyle.IncludeBorder = false;

                boughtStyle = styles.Add("Handlowe");
                boughtStyle.Interior.Color = ColorTranslator.ToOle(Color.Aquamarine);
                boughtStyle.IncludeAlignment = false;
                boughtStyle.IncludeBorder = false;

                normalStyle = styles.Add("Normalia");
                normalStyle.Interior.Color = ColorTranslator.ToOle(Color.CadetBlue);
                normalStyle.IncludeAlignment = false;
                normalStyle.IncludeBorder = false;

                sheetMetalStyle = styles.Add("Blacha");
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

        public static void Type(Excel.Worksheet worksheet, int typeNumber)
        {
            Excel.Range range = null;
            Excel.Range rows = null;
            Excel.Range columns = null;
            Excel.Range firstRow = null;
            try
            {
                range = worksheet.UsedRange;
                rows = range.Rows;
                columns = range.Columns;
                firstRow = rows[1];

                if (typeNumber == 0) return;
                foreach (Excel.Range row in rows)
                {
                    if (row.Row == firstRow.Row) continue;

                    string value = GetValue(range, row.Row, typeNumber); if (value == null) continue;
                    var cell = row.Cells[1, typeNumber];

                    if (value == "A") { cell.Value = "Złożenie"; continue; }
                    if (value == "C") { cell.Value = "Część"; continue; }
                    if (value == "K") { cell.Value = "Hutnicze"; continue; }
                    if (value == "H") { cell.Value = "Handlowe"; continue; }
                    if (value == "N") { cell.Value = "Normalia"; continue; }
                    if (value == "B") { cell.Value = "Blacha"; continue; }
                }
            }
            finally
            {
                CoreUtils.ReleaseCom(ref firstRow);
                CoreUtils.ReleaseCom(ref columns);
                CoreUtils.ReleaseCom(ref rows);
                CoreUtils.ReleaseCom(ref range);
            }
        }

        public static void Colors(Excel.Worksheet worksheet, int typeNumber)
        {
            Excel.Range range = null;
            Excel.Range dataRange = null;
            Excel.Range columns = null;
            Excel.FormatConditions conditions = null;
            try
            {
                range = worksheet.UsedRange;
                int lastRow = range.Rows.Count;
                int lastCol = range.Columns.Count;

                if (lastRow < 2) return;

                dataRange = range.Range[range.Cells[2, 1], range.Cells[lastRow, lastCol]];
                columns = worksheet.UsedRange.Columns;

                Excel.Range anchorCell = worksheet.Cells[2, typeNumber];
                string address = anchorCell.Address[RowAbsolute: false, ColumnAbsolute: true];

                conditions = dataRange.FormatConditions;
                conditions.Delete();

                Rule(conditions, address, "Złożenie", Color.Khaki);
                Rule(conditions, address, "Część", Color.Lavender);
                Rule(conditions, address, "Hutnicze", Color.SaddleBrown);
                Rule(conditions, address, "Handlowe", Color.Aquamarine);
                Rule(conditions, address, "Normalia", Color.CadetBlue);
                Rule(conditions, address, "Blacha", Color.Azure);
            }
            finally
            {
                CoreUtils.ReleaseCom(ref conditions);
                CoreUtils.ReleaseCom(ref columns);
                CoreUtils.ReleaseCom(ref dataRange);
                CoreUtils.ReleaseCom(ref range);
            }
        }

        public static void Count(Excel.Worksheet worksheet, int multiplier, int countNumber)
        {
            Excel.Range range = null;
            Excel.Range rows = null;
            try
            {
                range = worksheet.UsedRange;
                rows = range.Rows;
            
                int rowCount = rows.Count;
                for (int i = 3; i <= rowCount; i++)
                {
                    Excel.Range cell = null;
                    try
                    {
                        cell = (Excel.Range)range.Cells[i, countNumber];
                        object cellValue = cell.Value;

                        if (cellValue != null && int.TryParse(cellValue.ToString(), out int currentQty))
                        {
                            int newQty = currentQty * multiplier;
                            cell.Value = newQty;
                        }
                    }
                    finally
                    {
                        CoreUtils.ReleaseCom(ref cell);
                    }
                }
            }
            finally
            {
                CoreUtils.ReleaseCom(ref rows);
                CoreUtils.ReleaseCom(ref range);
            }
        }

        public static void Edit(Excel.Worksheet worksheet)
        {
            Excel.Range range = null;
            Excel.Range rows = null;
            Excel.Range columns = null;
            Excel.Range firstRow = null;
            Excel.Range cells = null;
            try
            {
                range = worksheet.UsedRange;
                rows = range.Rows;
                columns = range.Columns;
                firstRow = rows[1];
                cells = range.Cells;
                range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                range.Borders.Weight = Excel.XlBorderWeight.xlThin;
                range.Borders.Color = (int)Excel.XlRgbColor.rgbBlack;
                firstRow.RowHeight = 25;
                firstRow.Interior.Color = (int)Excel.XlRgbColor.rgbLightGray;
                firstRow.Font.Bold = true;
                foreach (Excel.Range column in columns) { column.AutoFit(); }
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

        public static string GetValue(Excel.Range range, int row, int column)
        {
            Excel.Range cells = null;
            string value = null;
            try
            {
                cells = range.Cells; object objectValue = cells[row, column].Value;
                if (objectValue != null) value = objectValue?.ToString()?.Trim();
                return value;
            }
            catch { return null; }
            finally { CoreUtils.ReleaseCom(ref cells); }
        }

        private static void Rule(Excel.FormatConditions conditions, string address, string criteria, Color color)
        {
            Excel.FormatCondition rule = null;
            try
            {
                string formula = $"={address}=\"{criteria}\"";
                rule = (Excel.FormatCondition)conditions.Add(Excel.XlFormatConditionType.xlExpression, Formula1: formula);
                rule.Interior.Color = ColorTranslator.ToOle(color);
                rule.StopIfTrue = false;
            }
            finally { CoreUtils.ReleaseCom(ref rule); }
        }
    }

}
