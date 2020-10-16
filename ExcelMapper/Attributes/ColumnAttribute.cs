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
        readonly MappingDirections directions;

        /// <summary>
        /// Initializes a new instance of the <see cref="ColumnAttribute"/> class.
        /// </summary>
        /// <param name="name">The name of the column.</param>
        /// <param name="directions">mapping direction</param>
        public ColumnAttribute(string name, MappingDirections directions = MappingDirections.Both)
        {
            this.name = name;
            this.directions = directions;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ColumnAttribute"/> class.
        /// </summary>
        /// <param name="index">The index of the column.</param>
        /// <param name="directions">mapping direction</param>
        public ColumnAttribute(int index, MappingDirections directions = MappingDirections.Both)
        {
            this.index = index;
            this.directions = directions;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ColumnAttribute"/> class.
        /// </summary>
        /// <param name="directions">mapping direction</param>
        public ColumnAttribute(MappingDirections directions)
        {
            this.directions = directions;
        }

        /// <summary>
        /// Gets the direction of the column.
        /// </summary>
        /// <value>
        /// The name of the column.
        /// </value>
        public MappingDirections Directions => directions;

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
