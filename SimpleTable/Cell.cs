namespace SimpleTable
{
    public interface ICell
    {
        object Value { get; set; }
        IColumn Column { get; }
    }

    internal class Cell<T> : ICell
    {
        private readonly Column<T> column;

        public T Value { get; set; }
        public Column<T> Column => column;

        object ICell.Value
        {
            get => Value;
            set => this.Value = (T)value;
        }

        IColumn ICell.Column => Column;

        internal Cell(Column<T> column) => (this.column, Value) = (column, column.DefaultValue);

        internal Cell(Column<T> column, T value) => (this.column, this.Value) = (column, value);
    }
}
