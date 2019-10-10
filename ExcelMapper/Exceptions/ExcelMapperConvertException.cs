using System;

namespace Ganss.Excel.Exceptions
{
    /// <summary>
    /// Represents an error that occurs when Excel Mapper is unable to convert a field to the expected type.
    /// </summary>
    public class ExcelMapperConvertException : Exception
    {
        public ExcelMapperConvertException() { }

        public ExcelMapperConvertException(string message) : base(message) { }

        public ExcelMapperConvertException(string message, Exception innerException) : base(message, innerException) { }

        public ExcelMapperConvertException(object cellValue, Type targetType, int line, int column) : base(FormatMessage(cellValue, targetType, line, column)) { }

        public ExcelMapperConvertException(object cellValue, Type targetType, int line, int column, Exception innerException) 
            : base(FormatMessage(cellValue, targetType, line, column), innerException) { }

        private static string FormatMessage(object cellValue, Type targetType, int line, int column)
            => $"Unable to convert \"{(string.IsNullOrWhiteSpace(cellValue.ToString()) ? "<EMPTY>" : cellValue)}\" from [L:{line}]:[C:{column}] to {targetType}.";
    }
}