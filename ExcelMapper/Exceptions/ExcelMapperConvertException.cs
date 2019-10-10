using System;

namespace Ganss.Excel.Exceptions
{
    /// <summary>
    /// Represents an error that occurs when Excel Mapper is unable to convert a field to the expected type.
    /// </summary>
    public class ExcelMapperConvertException : Exception
    {
        public ExcelMapperConvertException(string message) : base(message) { }

        public ExcelMapperConvertException(string message, Exception inner) : base(message, inner) { }
    }
}