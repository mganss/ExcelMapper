using NPOI.SS.UserModel;
using System;

namespace Ganss.Excel;

/// <summary>
/// Provides data for the <see cref="ExcelMapper.Saving"/> event.
/// </summary>
/// <remarks>
/// Initializes a new instance of the <see cref="SavingEventArgs"/> class.
/// </remarks>
/// <param name="sheet">The sheet that is being saved.</param>
public class SavingEventArgs(ISheet sheet) : EventArgs
{
    /// <summary>
    /// Gets or sets the sheet.
    /// </summary>
    /// <value>
    /// The sheet.
    /// </value>
    public ISheet Sheet { get; private set; } = sheet;
}
