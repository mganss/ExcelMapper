using Ganss.Excel.Exceptions;
using System.ComponentModel;

namespace Ganss.Excel;

/// <summary>
/// Provides data for the <see cref="ExcelMapper.ErrorParsingCell"/> event.
/// Event handler can allow parsing to continue by setting <see cref="CancelEventArgs.Cancel"/> to true,
/// cancelling the exception.
/// </summary>
/// <remarks>
/// Initializes a new instance of the <see cref="ParsingErrorEventArgs"/> class.
/// </remarks>
/// <param name="error">The error captured.</param>
public class ParsingErrorEventArgs(ExcelMapperConvertException error) : CancelEventArgs
{
    /// <summary>
    /// The error captured
    /// </summary>
    public ExcelMapperConvertException Error { get; private set; } = error;
}
