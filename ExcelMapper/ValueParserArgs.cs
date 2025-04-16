using NPOI.SS.UserModel;

namespace Ganss.Excel;

/// <summary>
/// Encapsulation of arguments passed to a Value Parser function.
/// </summary>
public class ValueParserArgs
{
    /// <summary>
    /// Gets or sets the cell.
    /// </summary>
    public ICell Cell { get; set; }

    /// <summary>
    /// Gets or sets the column name.
    /// </summary>
    public string ColumnName { get; set; }

    /// <summary>
    /// Gets or sets the cell value.
    /// </summary>
    public object CellValue { get; set; }
}
