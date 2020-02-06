namespace SimpleTable
{
    public abstract class Cell
    {
        public virtual object Value { get; set; }
        public virtual Column Column { get; }
    }

    internal class Cell<T> : Cell
    {
        private readonly Column<T> column;
        private T value;

        public override object Value { get => value; set => this.value = (T)value; }
        public override Column Column => column;

        internal Cell(Column<T> column) => (this.column, Value) = (column, column.DefaultValue);

        internal Cell(Column<T> column, T value) => (this.column, this.Value) = (column, value);

        internal T GetTypefiedValue() => value;

        internal Column<T> GetTypefiedColumn() => column;
    }
}
