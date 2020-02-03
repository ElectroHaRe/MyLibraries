using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace SimpleTable
{
    public class Row : IEnumerable
    {
        private Dictionary<string, ICell> cells = new Dictionary<string, ICell>();

        public int Count => cells.Count;

        public object this[string columnName]
        {
            get
            {
                if (!cells.ContainsKey(columnName))
                    throw new ArgumentException($"Cell with columnName '{columnName}' does not exists");

                return cells[columnName].Value;
            }

            set
            {
                if (!cells.ContainsKey(columnName))
                    throw new ArgumentException($"Cell with columnName '{columnName}' does not exists");

                cells[columnName].Value = value;
            }
        }

        public object this[int index]
        {
            get
            {
                if (index < 0 || index >= cells.Count)
                    throw new IndexOutOfRangeException($"Cells count = {cells.Count} , index = '{index}'");

                return cells.Values.ToList()[index].Value;
            }

            set
            {
                if (index < 0 || index >= cells.Count)
                    throw new IndexOutOfRangeException($"Cells count = {cells.Count} , index = '{index}'");

                this[cells.Keys.ToList()[index]] = value;
            }
        }

        public Row(params IColumn[] columns)
        {
            foreach (var column in columns)
            {
                if (cells.ContainsKey(column.Name))
                    throw new ArgumentException($"Cell with columnName '{column.Name}' already exists");

                cells.Add(column.Name, column.CreateCell());
            }
        }

        public bool Contains(string columnName) => cells.ContainsKey(columnName);

        public IColumn GetColumn(string columnName)
            => cells.ContainsKey(columnName) ? cells[columnName].Column
               : throw new ArgumentException($"Cell with columnName '{columnName}' does not exists");

        internal void RemoveCell(string columnName)
        {
            if (cells.ContainsKey(columnName))
                throw new ArgumentException($"Cell with columnName '{columnName}' does not exists");

            cells.Remove(columnName);
        }

        internal void AddCell(ICell cell)
        {
            if (cells.ContainsKey(cell.Column.Name))
                throw new ArgumentException($"Cell with columnName '{cell.Column.Name}' already exists");

            cells.Add(cell.Column.Name, cell);
        }

        public object[] GetValues() => (from cell in cells select cell.Value.Value).ToArray();

        public IEnumerator GetEnumerator() => GetValues().GetEnumerator();
    }
}
