using System;

namespace SimpleTable
{
    public interface ICell
    {
        object Value { get; set; }
        IColumn Column { get; }
    }

    internal struct Cell<T> : ICell
    {
        private T value;
        private readonly Column<T> column;

        public T Value { get => value; set => this.value = value; }
        public Column<T> Column => column;

        object ICell.Value
        {
            get => value;
            set
            {
                if (value is T)
                    this.value = (T)value;
                else
                    throw new ArgumentException($"Value type does not match cell type | Value = {value}", $"value = {value}");
            }              
        }

        IColumn ICell.Column => Column;

        internal Cell(Column<T> column) => (this.column, value) = (column, column.DefaultValue);

        internal Cell(Column<T> column, T value) => (this.column, this.value) = (column, value);
    }
}
