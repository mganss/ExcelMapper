using System;

namespace Ganss.Excel
{
    /// <summary>
    /// Attribute that specifies the mapping direction of a property to a column in an Excel file.
    /// </summary>
    /// <seealso cref="System.Attribute" />
    [AttributeUsage(AttributeTargets.Property, Inherited = true, AllowMultiple = false)]
    public sealed class ToExcelOnlyAttribute : Attribute
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ToExcelOnlyAttribute"/> class.
        /// </summary>
        public ToExcelOnlyAttribute()
        { }
    }
}
