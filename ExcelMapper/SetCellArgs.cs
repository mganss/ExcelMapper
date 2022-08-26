using System;
using System.Collections.Generic;
using NPOI.SS.UserModel;

namespace Ganss.Excel
{
    /// <summary>
    /// set excell cell parameter
    /// </summary>
    /// <typeparam name="TEntity"></typeparam>
    public class SetCellArgs<TEntity> : EventArgs where TEntity : class
    {
        /// <summary>
        /// excel cell
        /// </summary>
        public ICell Cell { get; }

        /// <summary>
        /// collection row data
        /// </summary>
        public TEntity Data { get; }

        /// <summary>
        /// property value
        /// </summary>
        public object Value { get; }

        /// <summary>
        /// set excell cell parameter
        /// </summary>
        /// <param name="cell">excel cell</param>
        /// <param name="data">collection row data</param>
        /// <param name="value">excel cell value</param>
        public SetCellArgs(ICell cell, TEntity data, object value)
        {
            Cell = cell;
            Data = data;
            Value = value;
        }
    }
}
