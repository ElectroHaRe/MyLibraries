using System;

namespace SimpleTable
{
    public abstract class Column
    {
        public virtual string Name { get; protected set; }

        public virtual Type ValueType => DefaultValue.GetType();

        public virtual object DefaultValue { get; protected set; }

        public override string ToString() => $"{Name} ({ValueType})";

        public abstract Cell CreateCell();
    }

    internal class Column<T> : Column
    {
        internal Column(string name, T defaultValue)
        {
            Name = name;
            DefaultValue = defaultValue;
        }

        public override Cell CreateCell() => CreateTypefiedCell();

        internal T GetTypefiedDefaultValue() => (T)DefaultValue;

        internal Cell<T> CreateTypefiedCell() => new Cell<T>(this);
    }
}