using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace SimpleTable
{
    public class Table : IEnumerable<Row>
    {
        public readonly string Name;

        private List<Column> columns = new List<Column>();
        private List<Row> rows = new List<Row>();

        public int ColumnCount => columns.Count;
        public int RowCount => rows.Count;

        public ReadOnlyCollection<Column> Columns => columns.AsReadOnly();

        public Row this[int rowIndex] => rowIndex >= 0 && rowIndex < RowCount ? rows[rowIndex]
                                         : throw new IndexOutOfRangeException($"Count = {RowCount} , rowIndex = {rowIndex}");

        public Column this[string columnName] => columns.Find((item) => item.Name == columnName) ??
                                                  throw new ArgumentException($"Column with name '{columnName}' does not exists");

        public Cell this[int rowIndex, int columnIndex] => this[rowIndex][columnIndex];

        public Cell this[int rowIndex, string columnName] => this[rowIndex][columnName];

        public int this[Row row] => rows.FindIndex((item) => item == row);

        public Table(string name) => this.Name = name;

        public Row FindRow(Predicate<Row> match) => rows.Find(match);

        public int FindRowIndex(Predicate<Row> match) => rows.FindIndex(match);

        public Row CreateRow() => new Row(columns.ToArray());

        /// <summary>
        /// Inserts a copy of the row
        /// </summary>
        /// <returns>Number of inserted values. If it's 0, the row is not inserted</returns>
        public int InsertRow(Row row)
        {
            return InsertRow(row, RowCount);
        }

        /// <summary>
        /// Inserts a copy of the row in the required position
        /// </summary>
        /// <returns>Number of inserted values. If it's 0, the row is not inserted</returns>
        public int InsertRow(Row row, int index)
        {
            var newRow = CreateRow();
            int counter = 0;

            foreach (var column in columns)
            {
                if (row.Contains(column.Name))
                {
                    newRow[column.Name].Value = row[column.Name].Value;
                    counter++;
                }
            }

            if (counter != 0)
                rows.Insert(index, newRow);

            return counter;
        }

        public static explicit operator List<object[]>(Table table)
        {
            List<object[]> array = new List<object[]>();

            var columnNames = from column in table.columns
                              select column.Name as object;
            var rowValues = from row in table.rows
                            select (object[])row;

            array.Add(columnNames.ToArray());
            array.AddRange(rowValues);

            return array;
        }

        public bool Contains(string columnName) => columns.Any((item) => item.Name == columnName);

        public Column GetColumn(int index) => index >= 0 && index < ColumnCount ? columns[index]
                                                : throw new IndexOutOfRangeException($"Count = {ColumnCount} , rowIndex = {index}");

        public int GetColumnPosition(string columnName) => columns.FindIndex((item) => item.Name == columnName);

        public void SetColumnPosition(string columnName, int newIndex)
        {
            var column = this[columnName];

            if (GetColumn(newIndex) == column)
                return;

            columns.Remove(column);
            columns.Insert(newIndex, column);

            var oldRows = rows;
            rows = new List<Row>();

            foreach (var oldRow in oldRows)
            {
                InsertRow(oldRow);
            }
        }

        public void RemoveRow(int index) => rows.Remove(this[index]);

        public bool RemoveRow(Row row) => rows.Remove(row);

        public void AddColumn<T>(string name, T defaultValue)
        {
            if (Contains(name))
                throw new ArgumentException($"Column with name '{name}' already exists");

            var newColumn = new Column<T>(name, defaultValue);
            columns.Add(newColumn);

            foreach (var row in rows)
            {
                row.AddCell(newColumn.CreateCell());
            }
        }

        public void RemoveColumn(string name)
        {
            var column = this[name];

            foreach (var row in rows)
            {
                row.RemoveCell(name);
            }

            columns.Remove(column);
        }

        public IEnumerator<Row> GetEnumerator() => rows.GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
