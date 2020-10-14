using System;

namespace Ganss.Excel
{
    /// <summary>
    /// Attribute that specifies the mapping direction of a property to a column in an Excel file.
    /// </summary>
    /// <seealso cref="System.Attribute" />
    [AttributeUsage(AttributeTargets.Property, Inherited = true, AllowMultiple = false)]
    public sealed class FromExcelOnlyAttribute : Attribute
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="FromExcelOnlyAttribute"/> class.
        /// </summary>
        public FromExcelOnlyAttribute()
        { }
    }
}
