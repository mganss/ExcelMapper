using System;
using NPOI.SS.UserModel;

namespace Ganss.Excel
{
    /// <summary>
    /// set excell cell parameter
    /// </summary>
    /// <typeparam name="TEntity"></typeparam>
    /// <typeparam name="TValue"></typeparam>
    public class SetCellArgs<TEntity, TValue> : EventArgs where TEntity : class
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
        public TValue Value { get; }

        /// <summary>
        /// set excell cell parameter
        /// </summary>
        /// <param name="cell">excel cell</param>
        /// <param name="data">collection row data</param>
        /// <param name="value">excel cell value</param>
        public SetCellArgs(ICell cell, TEntity data, TValue value)
        {
            Cell = cell;
            Data = data;
            Value = value;
        }
    }

    /// <summary>
    /// set excell cell parameter
    /// </summary>
    /// <typeparam name="TEntity"></typeparam>
    public class SetCellArgs<TEntity> : SetCellArgs<TEntity, object> where TEntity : class
    {

        /// <summary>
        /// set excell cell parameter
        /// </summary>
        /// <param name="cell">excel cell</param>
        /// <param name="data">collection row data</param>
        /// <param name="value">excel cell value</param>
        public SetCellArgs(ICell cell, TEntity data, object value) : base(cell, data, value)
        {
        }
    }
}
