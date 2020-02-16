using OfficeOpenXml;
using SimpleTable;
using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelAdapter
{
    public static class TableExtension
    {
        /// <summary>
        /// Saves the table in excel format
        /// </summary>
        /// <param name="table">Source table (the sheet in the workbook will have a table name)</param>
        /// <param name="savePath">Save file path (includes file name and its extension)</param>
        /// <param name="progressChanged">Called when conversion progress and table retention change(double (0,100])</param>
        public static void ToExcel(this Table table, string savePath, Action<double> progressChanged = null) =>
            table.ToExcel(savePath, new TableOptions(), progressChanged);

        /// <summary>
        /// Saves the table in excel format
        /// </summary>
        /// <param name="table">Source table (the sheet in the workbook will have a table name)</param>
        /// <param name="savePath">Save file path (includes file name and its extension)</param>
        /// <param name="tableOptions">Some working sheet markup parameters</param>
        /// <param name="progressChanged">Called when conversion progress and table retention change(double (0,100])</param>
        public static void ToExcel(this Table table, string savePath, TableOptions tableOptions, Action<double> progressChanged = null)
        {
            var excelPackage = new ExcelPackage(new FileInfo(savePath));
            table.ToExcel(excelPackage, tableOptions, (currentProgress) => progressChanged?.Invoke(currentProgress * 0.99));
            excelPackage.Save();
            progressChanged?.Invoke(1);
        }

        /// <summary>
        /// Saves the table in excel format
        /// </summary>
        /// <param name="table">Source table (the sheet in the workbook will have a table name)</param>
        /// <param name="excelPackage"></param>
        /// <param name="progressChanged">Called when conversion progress and table retention change(double (0,100])</param>
        /// <returns></returns>
        public static ExcelWorksheet ToExcel(this Table table, ExcelPackage excelPackage, Action<double> progressChanged = null) =>
            table.ToExcel(excelPackage, new TableOptions(), progressChanged);

        /// <summary>
        /// Saves the table in excel format
        /// </summary>
        /// <param name="table">Source table (the sheet in the workbook will have a table name)</param>
        /// <param name="excelPackage"></param>
        /// <param name="tableOptions">Some working sheet markup parameters</param>
        /// <param name="progressChanged">Called when conversion progress and table retention change(double (0,100])</param>
        /// <returns></returns>
        public static ExcelWorksheet ToExcel(this Table table, ExcelPackage excelPackage, TableOptions tableOptions, Action<double> progressChanged = null)
        {
            var workSheet = excelPackage.Workbook.Worksheets.Add(table.Name);
            workSheet.Cells.LoadFromArrays((List<object[]>)table).AutoFitColumns();
            tableOptions.ConfigurateSheet(workSheet, table.RowCount + 1, table.ColumnCount, progressChanged);
            return workSheet;
        }

        /// <summary>
        /// Saves tables in excel format into one workbook
        /// </summary>
        /// <param name="tables">Source tables (sheets in the workbook will have their names. Workbook does not support sheets of the same name)</param>
        /// <param name="savePath">Save file path (includes file name and its extension)</param>
        /// <param name="progressChanged">Called when the conversion progress and table writing to excelPackage(double (0,100]) are changed</param>
        public static void ToExcel(this Table[] tables, string savePath, Action<double> progressChanged = null) =>
            tables.ToExcel(savePath, new TableOptions(), progressChanged);

        /// <summary>
        /// Saves tables in excel format into one workbook
        /// </summary>
        /// <param name="tables">Source tables (sheets in the workbook will have their names. Workbook does not support sheets of the same name)</param>
        /// <param name="savePath">Save file path (includes file name and its extension)</param>
        /// <param name="tableOptions">Some working sheet markup parameters</param>
        /// <param name="progressChanged">Called when the conversion progress and table writing to excelPackage(double (0,100]) are changed</param>
        public static void ToExcel(this Table[] tables, string savePath, TableOptions tableOptions, Action<double> progressChanged = null)
        {
            var excelPackage = new ExcelPackage(new FileInfo(savePath));
            tables.ToExcel(excelPackage, tableOptions, (currentProgress) => progressChanged?.Invoke(currentProgress * 0.99));
            excelPackage.Save();
            progressChanged?.Invoke(1);
        }

        /// <summary>
        /// Saves tables in excel format into one workbook
        /// </summary>
        /// <param name="tables">Source tables (sheets in the workbook will have their names. Workbook does not support sheets of the same name)</param>
        /// <param name="excelPackage"></param>
        /// <param name="progressChanged">Called when conversion progress and table retention change(double (0,100])</param>
        /// <returns></returns>
        public static ExcelWorksheets ToExcel(this Table[] tables, ExcelPackage excelPackage, Action<double> progressChanged = null) =>
            tables.ToExcel(excelPackage, new TableOptions(), progressChanged);

        /// <summary>
        /// Saves tables in excel format into one workbook
        /// </summary>
        /// <param name="tables">Source tables (sheets in the workbook will have their names. Workbook does not support sheets of the same name)</param>
        /// <param name="excelPackage"></param>
        /// <param name="tableOptions">Some working sheet markup parameters</param>
        /// <param name="progressChanged">Called when conversion progress and table retention change(double (0,100])</param>
        /// <returns></returns>
        public static ExcelWorksheets ToExcel(this Table[] tables, ExcelPackage excelPackage, TableOptions tableOptions, Action<double> progressChanged = null)
        {
            double size = 0;
            double coeff = 0;
            double progress = 0;

            foreach (var table in tables)
            {
                size += (table.RowCount + 1) * table.ColumnCount;
            }

            foreach (var table in tables)
            {
                coeff = ((table.RowCount + 1) * table.ColumnCount) / size;
                table.ToExcel(excelPackage, tableOptions,
                    (progressForTable) =>
                    progressChanged?.Invoke(calculateProgress(progress, progressForTable, coeff)));
                progress += 0.99 * coeff;
            }

            progressChanged?.Invoke(1);
            return excelPackage.Workbook.Worksheets;

            static double calculateProgress(double currentProgress, double progressForTable, double coeff) =>
                currentProgress + progressForTable * coeff;
        }
    }
}
