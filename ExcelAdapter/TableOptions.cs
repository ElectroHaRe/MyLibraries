using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Drawing;

namespace ExcelAdapter
{
    public class TableOptions
    {
        public Borders HeaderCellBorders;
        public Borders HeaderBorders;
        public Borders ValueCellBorders;
        public Borders TableBorders;

        public ExcelHorizontalAlignment HorizontalAligment;
        public ExcelVerticalAlignment VerticalAligment;

        public Action<ExcelRangeBase> HeaderCellOperations;
        public Action<ExcelRangeBase> ValueCellOperations;

        public TableOptions()
        {
            HeaderBorders.Top = HeaderBorders.Left =
                HeaderBorders.Right = new BorderStyle(ExcelBorderStyle.Medium, Color.Black);
            HeaderBorders.Bottom = new BorderStyle(ExcelBorderStyle.Double, Color.Black);

            HeaderCellBorders.Top = HeaderCellBorders.Bottom = HeaderCellBorders.Left =
                HeaderCellBorders.Right = new BorderStyle(ExcelBorderStyle.Thin, Color.Gray);

            ValueCellBorders = HeaderCellBorders;
            TableBorders = HeaderBorders;

            HorizontalAligment = ExcelHorizontalAlignment.Center;
            VerticalAligment = ExcelVerticalAlignment.Center;
        }

        public void ConfigurateTable(ExcelWorksheet worksheet, int rowCount, int columnCount, Action<double> progressChanged = null)
        {
            double cellsCount = rowCount * columnCount + 1;
            int counter = 0;

            var headerRange = worksheet.Cells[1, 1, 1, columnCount];
            var valueCellRange = worksheet.Cells[2, 1, rowCount, columnCount];
            var tableRange = worksheet.Cells[1, 1, rowCount, columnCount];

            SetBorders(headerRange, HeaderBorders);
            SetBorders(tableRange, TableBorders);

            foreach (var cell in headerRange)
            {
                SetBorders(cell, HeaderCellBorders);
                HeaderCellOperations?.Invoke(cell);
                progressChanged?.Invoke(++counter * 100 / cellsCount);
            }

            foreach (var cell in valueCellRange)
            {
                SetBorders(cell, ValueCellBorders);
                ValueCellOperations?.Invoke(cell);
                progressChanged?.Invoke(++counter * 100 / cellsCount);
            }

            tableRange.Style.HorizontalAlignment = HorizontalAligment;
            tableRange.Style.VerticalAlignment = VerticalAligment;

            progressChanged?.Invoke(++counter * 100 / cellsCount);
        }

        public static void SetBorders(ExcelRangeBase range, Borders borders)
        {
            range.Style.Border.Top.Style = borders.Top.LineStyle;
            range.Style.Border.Top.Color.SetColor(borders.Top.LineColor);

            range.Style.Border.Bottom.Style = borders.Bottom.LineStyle;
            range.Style.Border.Bottom.Color.SetColor(borders.Bottom.LineColor);

            range.Style.Border.Left.Style = borders.Left.LineStyle;
            range.Style.Border.Left.Color.SetColor(borders.Left.LineColor);

            range.Style.Border.Right.Style = borders.Right.LineStyle;
            range.Style.Border.Right.Color.SetColor(borders.Right.LineColor);
        }
    }
}
