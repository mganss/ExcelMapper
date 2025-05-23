﻿using NPOI.SS.UserModel;

namespace Ganss.Excel;

/// <summary>
/// Encapsulation of arguments passed to a Value Converter function.
/// </summary>
public class ValueConverterArgs
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
