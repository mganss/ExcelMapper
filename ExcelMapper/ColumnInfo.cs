using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq.Expressions;
using System.Reflection;
using System.Text.Json;

namespace Ganss.Excel
{
    /// <summary>
    /// Describes the mapping of a property to a cell in an Excel sheet.
    /// </summary>
    public class ColumnInfo
    {
        private PropertyInfo property;

        /// <summary>
        /// Gets or sets the property.
        /// </summary>
        /// <value>
        /// The property.
        /// </value>
        public PropertyInfo Property
        {
            get { return property; }
            set
            {
                property = value;
                if (property != null)
                {
                    var underlyingType = Nullable.GetUnderlyingType(property.PropertyType);
                    IsNullable = underlyingType != null;
                    PropertyType = underlyingType ?? property.PropertyType;
                }
                else
                {
                    PropertyType = null;
                    IsNullable = false;
                }
            }
        }

        /// <summary>
        /// Gets a value indicating whether the property is nullable.
        /// </summary>
        /// <value>
        /// <c>true</c> if the property is nullable; otherwise, <c>false</c>.
        /// </value>
        public bool IsNullable { get; private set; }

        /// <summary>
        /// Gets the type of the property.
        /// </summary>
        /// <value>
        /// The type of the property.
        /// </value>
        public Type PropertyType { get; private set; }

        /// <summary>
        /// Gets or sets the cell setter.
        /// </summary>
        /// <value>
        /// The cell setter.
        /// </value>
        public Action<ICell, object> SetCell { get; set; }

        /// <summary>
        /// Gets or sets the property setter.
        /// </summary>
        /// <value>
        /// The property setter.
        /// </value>
        public Func<object, object> SetProp { get; set; }

        /// <summary>
        /// Gets or sets the builtin format.
        /// </summary>
        /// <value>
        /// The builtin format.
        /// </value>
        public short BuiltinFormat { get; set; }

        /// <summary>
        /// Gets or sets the custom format.
        /// </summary>
        /// <value>
        /// The custom format.
        /// </value>
        public string CustomFormat { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether to map the formula result.
        /// </summary>
        /// <value>
        /// <c>true</c> if the formula result will be mapped; otherwise, <c>false</c>.
        /// </value>
        public bool FormulaResult { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether to serialize as JSON.
        /// </summary>
        /// <value>
        /// <c>true</c> if the property will be serialized as JSON; otherwise, <c>false</c>.
        /// </value>
        public bool Json { get; set; }

        static readonly HashSet<Type> NumericTypes = new HashSet<Type>
        {
            typeof(decimal),
            typeof(byte), typeof(sbyte),
            typeof(short), typeof(ushort),
            typeof(int), typeof(uint),
            typeof(long), typeof(ulong),
            typeof(float), typeof(double)
        };

        Action<ICell, object> GenerateCellSetter()
        {
            if (PropertyType == typeof(DateTime))
            {
                return (c, o) =>
                {
                    if (o == null)
                        c.SetCellValue((string)null);
                    else
                        c.SetCellValue((DateTime)o);
                };
            }
            else if (PropertyType == typeof(bool))
            {
                if (IsNullable)
                    return (c, o) =>
                    {
                        if (o == null)
                            c.SetCellValue((string)null);
                        else
                            c.SetCellValue((bool)o);
                    };
                else
                    return (c, o) => c.SetCellValue((bool)o);
            }
            else if (NumericTypes.Contains(PropertyType))
            {
                if (IsNullable)
                    return (c, o) =>
                    {
                        if (o == null)
                            c.SetCellValue((string)null);
                        else
                            c.SetCellValue(Convert.ToDouble(o));
                    };
                else
                    return (c, o) => c.SetCellValue(Convert.ToDouble(o));
            }
            else
            {
                return (c, o) =>
                {
                    if (o == null)
                        c.SetCellValue((string)null);
                    else if (!Json)
                        c.SetCellValue(o.ToString());
                    else
                        c.SetCellValue(JsonSerializer.Serialize(o));
                };
            }
        }

        /// <summary>Sets the column style.</summary>
        /// <param name="sheet">The sheet.</param>
        /// <param name="columnIndex">Index of the column.</param>
        public void SetColumnStyle(ISheet sheet, int columnIndex)
        {
            if (BuiltinFormat != 0 || CustomFormat != null)
            {
                var wb = sheet.Workbook;
                var cs = wb.CreateCellStyle();
                if (CustomFormat != null)
                    cs.DataFormat = wb.CreateDataFormat().GetFormat(CustomFormat);
                else
                    cs.DataFormat = BuiltinFormat;
                sheet.SetDefaultColumnStyle(columnIndex, cs);
            }
        }

        /// <summary>Sets the cell style.</summary>
        /// <param name="c">The cell.</param>
        public void SetCellStyle(ICell c)
        {
            if (BuiltinFormat != 0 || CustomFormat != null)
                c.CellStyle = c.Sheet.GetColumnStyle(c.ColumnIndex);
        }

        /// <summary>
        /// Sets the property of the specified object to the specified value.
        /// </summary>
        /// <param name="o">The object whose property to set.</param>
        /// <param name="val">The value.</param>
        public virtual void SetProperty(object o, object val)
        {
            object v;
            if (SetProp != null)
                v = SetProp(val);
            else if (IsNullable && (val == null || (val is string s && s.Length == 0)))
                v = null;
            else
                v = Convert.ChangeType(val, PropertyType, CultureInfo.InvariantCulture);

            Property.SetValue(o, v, null);
        }

        /// <summary>
        /// Gets the property value of the specified object.
        /// </summary>
        /// <param name="o">The o.</param>
        /// <returns></returns>
        public object GetProperty(object o)
        {
            return Property.GetValue(o, null);
        }

        /// <summary>Specifies a method to use when setting the cell value from an object.</summary>
        /// <param name="setCell">The method to use when setting the cell value from an object.</param>
        /// <returns>The <see cref="ColumnInfo"/> object.</returns>
        public ColumnInfo SetCellUsing(Action<ICell, object> setCell)
        {
            SetCell = setCell;
            return this;
        }

        /// <summary>Specifies a method to use when setting the property value from the cell value.</summary>
        /// <param name="setProp">The method to use when setting the property value from the cell value.</param>
        /// <returns>The <see cref="ColumnInfo"/> object.</returns>
        public ColumnInfo SetPropertyUsing(Func<object, object> setProp)
        {
            SetProp = setProp;
            return this;
        }

        /// <summary>Selects formula results to be mapped instead of the formula itself.</summary>
        /// <returns>The <see cref="ColumnInfo"/> object.</returns>
        public ColumnInfo AsFormulaResult()
        {
            FormulaResult = true;
            return this;
        }

        /// <summary>Selects the property to be serialized as JSON.</summary>
        /// <returns>The <see cref="ColumnInfo"/> object.</returns>
        public ColumnInfo AsJson()
        {
            Json = true;
            return this;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ColumnInfo"/> class.
        /// </summary>
        /// <param name="propertyInfo">The property information.</param>
        public ColumnInfo(PropertyInfo propertyInfo)
        {
            Property = propertyInfo;
            SetCell = GenerateCellSetter();
            if (PropertyType == typeof(DateTime))
                BuiltinFormat = 0x16; // "m/d/yy h:mm"
        }

    }

    /// <summary>
    /// Describes the mapping of a property to a cell in an Excel sheet.
    /// </summary>
    public class ColumnInfo<T> : ColumnInfo
    {
        /// <summary>
        /// Next mapping in the setProp() chain
        /// </summary>
        internal ColumnInfo<T> NextSetProp;

        /// <summary>
        /// Initializes a new instance of the <see cref="ColumnInfo{T}"/> class.
        /// </summary>
        /// <param name="propertyInfo">The property information.</param>
        public ColumnInfo(PropertyInfo propertyInfo) : base(propertyInfo)
        { }

        /// <summary>Specifies a method to use when setting the cell value from an object.</summary>
        /// <param name="setCell">The method to use when setting the cell value from an object.</param>
        /// <returns>The <see cref="ColumnInfo{T}"/> object.</returns>
        public new ColumnInfo<T> SetCellUsing(Action<ICell, object> setCell)
        {
            SetCell = setCell;
            return this;
        }

        /// <summary>Specifies a method to use when setting the property value from the cell value.</summary>
        /// <param name="setProp">The method to use when setting the property value from the cell value.</param>
        /// <returns>The <see cref="ColumnInfo{T}"/> object.</returns>
        public new ColumnInfo<T> SetPropertyUsing(Func<object, object> setProp)
        {
            SetProp = setProp;
            return this;
        }

        /// <summary>Specifies a method to use when setting the <paramref name="propertyExpression"/> value from the cell value.</summary>
        /// <param name="propertyExpression">The property expression.</param>
        /// <param name="setProp">The method to use when setting the property value from the cell value.</param>
        /// <returns>The <see cref="ColumnInfo{T}"/> object.</returns>
        public ColumnInfo<T> ThenSetPropertyUsing(Expression<Func<T, object>> propertyExpression, Func<object, object> setProp)
        {
            var prop = ExcelMapper.GetPropertyInfo(propertyExpression);
            NextSetProp = new ColumnInfo<T>(prop);
            NextSetProp.SetProp = setProp;
            return NextSetProp;
        }

        /// <summary>
        /// Sets the property of the specified object to the specified value.
        /// </summary>
        /// <param name="o">The object whose property to set.</param>
        /// <param name="val">The value.</param>
        public override void SetProperty(object o, object val)
        {
            base.SetProperty(o, val);

            // Chain all the the linked set properties
            if (this.NextSetProp != null)
                this.NextSetProp.SetProperty(o, val);
        }
    }

}
