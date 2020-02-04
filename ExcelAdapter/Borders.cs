using System.Drawing;
using OfficeOpenXml.Style;

namespace ExcelAdapter
{
    public struct BorderStyle
    {
        public ExcelBorderStyle LineStyle;
        public Color LineColor;

        public BorderStyle(ExcelBorderStyle lineStyle, Color lineColor) =>
            (LineStyle, LineColor) = (lineStyle, lineColor);
    }

    public struct Borders
    {
        public BorderStyle Top;
        public BorderStyle Bottom;
        public BorderStyle Left;
        public BorderStyle Right;
    }
}
