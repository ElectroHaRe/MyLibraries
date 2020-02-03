using System;

namespace SimpleTable
{
    public interface IColumn
    {
        string Name { get; }
        Type ValueType { get; }
        object DefaultValue { get; }

        ICell CreateCell();
    }

    internal struct Column<T> : IColumn
    {
        private readonly string name;
        private readonly T defaultValue;

        public string Name => Name;
        public Type ValueType => typeof(T);
        internal T DefaultValue => defaultValue;
        object IColumn.DefaultValue => defaultValue;

        internal Column(string name, T defaultValue) => (this.name, this.defaultValue) = (name, defaultValue);

        public override string ToString() => $"{name} ({ValueType})";

        internal Cell<T> CreateCell() => new Cell<T>(this);

        ICell IColumn.CreateCell() => CreateCell();
    }
}