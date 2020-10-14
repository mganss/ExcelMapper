using System;

namespace Ganss.Excel
{

    /// <summary>
    /// Data direction
    /// </summary>
    [Flags]
    public enum ColumnInfoDirections
    {
        /// <summary>
        /// From Excel to Object
        /// </summary>
        Cell2Prop = 1 << 0,
        /// <summary>
        /// From Object to Excel
        /// </summary>
        Prop2Cell = 1 << 1,
        /// <summary>
        /// Both directions
        /// </summary>
        Both = Cell2Prop | Prop2Cell,
    }
}
