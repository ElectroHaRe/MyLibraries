using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace SimpleTable
{
    public class Row : IEnumerable<Cell>
    {
        private Dictionary<string, Cell> cells = new Dictionary<string, Cell>();

        public int Count => cells.Count;

        public Cell this[string columnName]
        {
            get
            {
                if (!Contains(columnName))
                    throw new ArgumentException($"Cell with columnName '{columnName}' does not exists");

                return cells[columnName];
            }
        }

        public Cell this[int index]
        {
            get
            {
                if (index < 0 || index >= cells.Count)
                    throw new IndexOutOfRangeException($"Cells count = {cells.Count} , index = '{index}'");

                return cells.Values.ToList()[index];
            }
        }

        public Row(params Column[] columns)
        {
            foreach (var column in columns)
            {
                if (Contains(column.Name))
                    throw new ArgumentException($"Cell with columnName '{column.Name}' already exists");

                cells.Add(column.Name, column.CreateCell());
            }
        }

        public bool Contains(string columnName) => cells.ContainsKey(columnName);

        public Column GetColumn(string columnName)
            => Contains(columnName) ? cells[columnName].Column
               : throw new ArgumentException($"Cell with columnName '{columnName}' does not exists");

        internal void RemoveCell(string columnName)
        {
            if (!Contains(columnName))
                throw new ArgumentException($"Cell with columnName '{columnName}' does not exists");

            cells.Remove(columnName);
        }

        internal void AddCell(Cell cell)
        {
            if (Contains(cell.Column.Name))
                throw new ArgumentException($"Cell with columnName '{cell.Column.Name}' already exists");

            cells.Add(cell.Column.Name, cell);
        }

        public static explicit operator object[](Row row) => (from cell in row.cells select cell.Value.Value).ToArray();

        public IEnumerator<Cell> GetEnumerator()
        {
            foreach (var cell in cells)
            {
                yield return cell.Value;
            }
        }

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
