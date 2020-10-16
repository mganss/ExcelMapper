using System;

namespace Ganss.Excel
{
    /// <summary>
    /// Attribute that specifies the mapping of a property to a column in an Excel file.
    /// </summary>
    /// <seealso cref="System.Attribute" />
    [AttributeUsage(AttributeTargets.Property, Inherited = true, AllowMultiple = true)]
    public sealed class ColumnAttribute : Attribute
    {
        readonly string name = null;
        readonly int index = -1;
        readonly ColumnInfoDirections direction;

        /// <summary>
        /// Initializes a new instance of the <see cref="ColumnAttribute"/> class.
        /// </summary>
        /// <param name="name">The name of the column.</param>
        /// <param name="direction">mapping direction</param>
        public ColumnAttribute(string name, ColumnInfoDirections direction = ColumnInfoDirections.Both)
        {
            this.name = name;
            this.direction = direction;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ColumnAttribute"/> class.
        /// </summary>
        /// <param name="index">The index of the column.</param>
        /// <param name="direction">mapping direction</param>
        public ColumnAttribute(int index, ColumnInfoDirections direction = ColumnInfoDirections.Both)
        {
            this.index = index;
            this.direction = direction;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ColumnAttribute"/> class.
        /// </summary>
        /// <param name="direction">mapping direction</param>
        public ColumnAttribute(ColumnInfoDirections direction)
        {
            this.direction = direction;
        }

        /// <summary>
        /// Gets the direction of the column.
        /// </summary>
        /// <value>
        /// The name of the column.
        /// </value>
        public ColumnInfoDirections Direction => direction;

        /// <summary>
        /// Gets the name of the column.
        /// </summary>
        /// <value>
        /// The name of the column.
        /// </value>
        public string Name => name;

        /// <summary>
        /// Gets the index of the column.
        /// </summary>
        /// <value>
        /// The index of the column.
        /// </value>
        public int Index => index;

    }
}
