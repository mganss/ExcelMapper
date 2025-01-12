using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text.Json;
using Ganss.Excel.Exceptions;

namespace Ganss.Excel
{
    /// <summary>
    /// Describes the mapping of a property to a cell in an Excel sheet.
    /// </summary>
    public class ColumnInfo
    {
        /// <summary>
        /// Gets or sets the mapping directions.
        /// </summary>
        public MappingDirections Directions { get; internal set; } = MappingDirections.Both;

        private PropertyInfo property;
        private bool isSubType;

        /// <summary>
        /// Gets or sets the default cell setter.
        /// </summary>
        protected Action<ICell, object> defaultCellSetter;
        protected Action<object, ICell, object> customCellSetter;

        /// <summary>
        /// Sets the property type.
        /// </summary>
        /// <param name="propertyType">The property type.</param>
        internal void SetPropertyType(Type propertyType)
        {
            if (propertyType.IsValueType)
            {
                var underlyingType = Nullable.GetUnderlyingType(propertyType);
                IsNullable = underlyingType != null;
                PropertyType = underlyingType ?? propertyType;
            }
            else
            {
                IsNullable = true;
                PropertyType = propertyType;
            }

            isSubType = !PropertyType.IsPrimitive
                && !PropertyType.IsEnum
                && PropertyType != typeof(decimal)
                && PropertyType != typeof(string)
                && !DateTypes.Contains(PropertyType.FullName)
                && PropertyType != typeof(Guid)
                && PropertyType != typeof(byte[])
                && PropertyType != typeof(object);
        }

        /// <summary>
        /// Gets or sets the property name.
        /// </summary>
        public string Name { get; protected set; }

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
                    SetPropertyType(property.PropertyType);
                    Name = property.Name;
                }
                else
                {
                    PropertyType = null;
                    IsNullable = false;
                }
            }
        }

        /// <summary>
        /// Gets a value indicating whether the mapped property has a nested type.
        /// </summary>
        public bool IsSubType
        {
            get
            {
                return isSubType && !Json
                    && SetProp == null;
            }
        }

        /// <summary>
        /// Gets a value indicating whether the property is nullable.
        /// </summary>
        /// <value>
        /// <c>true</c> if the property is nullable; otherwise, <c>false</c>.
        /// </value>
        public bool IsNullable { get; protected set; }

        /// <summary>
        /// Gets the type of the property.
        /// </summary>
        /// <value>
        /// The type of the property.
        /// </value>
        public Type PropertyType { get; protected set; }

        /// <summary>
        /// Gets or sets the cell setter.
        /// </summary>
        /// <value>
        /// The cell setter.
        /// </value>
        public Action<object, ICell, object> SetCell => customCellSetter != null ? customCellSetter : (entity, cell, value) => defaultCellSetter(cell, value);

        /// <summary>
        /// Gets or sets the property setter.
        /// </summary>
        /// <value>
        /// The property setter.
        /// </value>
        public Func<object, object, ICell, object> SetProp { get; set; }

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
        /// Gets or sets a value indicating whether to save the property as a formula cell. Only needed when saving.
        /// </summary>
        /// <value>
        /// <c>true</c> if the cell will contain a formula; otherwise, <c>false</c>.
        /// </value>
        public bool Formula { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether to serialize as JSON.
        /// </summary>
        /// <value>
        /// <c>true</c> if the property will be serialized as JSON; otherwise, <c>false</c>.
        /// </value>
        public bool Json { get; set; }

        /// <summary>
        /// Types that map a numeric column.
        /// </summary>
        protected static readonly HashSet<Type> NumericTypes =
        [
            typeof(decimal),
            typeof(byte), typeof(sbyte),
            typeof(short), typeof(ushort),
            typeof(int), typeof(uint),
            typeof(long), typeof(ulong),
            typeof(float), typeof(double)
        ];

        /// <summary>
        /// Types that map to a DateTime column.
        /// </summary>
        protected static readonly HashSet<string> DateTypes =
        [
            typeof(DateTime).FullName,
            typeof(DateTimeOffset).FullName,
            "System.DateOnly"
        ];

        static DateTime DateOnlyToDateTime(object date)
        {
            var dateOnly = date.GetType();
            var day = (int)dateOnly.GetProperty("Day").GetValue(date);
            var month = (int)dateOnly.GetProperty("Month").GetValue(date);
            var year = (int)dateOnly.GetProperty("Year").GetValue(date);

            return new DateTime(year, month, day, 0, 0, 0, DateTimeKind.Unspecified);
        }

        /// <summary>
        /// Generates the cell setter.
        /// </summary>
        /// <returns>The cell setter.</returns>
        protected Action<ICell, object> GenerateCellSetter()
        {
            if (DateTypes.Contains(PropertyType.FullName))
            {
                return (c, o) =>
                {
                    if (o == null)
                        c.SetCellValue((string)null);
                    else if (o is DateTimeOffset dt)
                        c.SetCellValue(dt.DateTime);
                    else if (o.GetType().FullName == "System.DateOnly")
                        c.SetCellValue(DateOnlyToDateTime(o));
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
            else if (PropertyType == typeof(byte[]))
            {
                return (c, o) =>
                {
                    if (o == null)
                        c.SetCellValue((string)null);
                    else
                        c.SetCellValue(System.Text.Encoding.UTF8.GetString((byte[])o));
                };
            }
            else
            {
                return (c, o) =>
                {
                    if (o == null)
                        c.SetCellValue((string)null);
                    else if (Formula)
                        c.SetCellFormula(o.ToString());
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
            {
                c.CellStyle = c.Sheet.GetColumnStyle(c.ColumnIndex);
                //When it is a dynamic type, after initialization, if the first row is null and the dataformat is not set, the dataformat needs to be repaired again
                if (c.CellStyle.DataFormat == 0)
                {
                    SetColumnStyle(c.Sheet, c.ColumnIndex);
                    c.CellStyle = c.Sheet.GetColumnStyle(c.ColumnIndex);
                }
            }
        }

        private object ParseEnum(Type t, string s)
        {
            var name = Array.Find(Enum.GetNames(t), n => n.Equals(s, StringComparison.OrdinalIgnoreCase));
            return name == null
                ? throw new ArgumentException($"Did not find a matching enum name for {s} in enum type {t.Name}.", nameof(s))
                : Enum.Parse(t, name);
        }

        /// <summary>
        /// Computes value that can be assigned to property from cell value.
        /// </summary>
        /// <param name="o">The object which contains the property.</param>
        /// <param name="val">The value.</param>
        /// <param name="cell">The cell where the value originates from.</param>
        /// <returns>Value that can be assigned to property.</returns>
        public virtual object GetPropertyValue(object o, object val, ICell cell)
        {
            object v;

            if (SetProp != null)
                v = SetProp(o, val, cell);
            else if (IsNullable && val == null)
                v = null;
            else if (val is string s && s.Length == 0)
                v = string.Empty;
            else if (val is string g && PropertyType == typeof(Guid))
                v = Guid.Parse(g);
            else if (val is string es && PropertyType.IsEnum)
                v = ParseEnum(PropertyType, es);
            else if (val is string && PropertyType == typeof(byte[]))
                v = System.Text.Encoding.UTF8.GetBytes(val as string);
            else if (val is DateTime d && PropertyType == typeof(DateTimeOffset))
                v = new DateTimeOffset(d);
            else if (val is DateTime && PropertyType.FullName == "System.DateOnly")
                v = PropertyType.GetMethod("FromDateTime").Invoke(null, [val]);
            else
                v = Convert.ChangeType(val, PropertyType, CultureInfo.InvariantCulture);

            return v;
        }

        /// <summary>
        /// Sets the property of the specified object to the specified value.
        /// </summary>
        /// <param name="o">The object whose property to set.</param>
        /// <param name="val">The value.</param>
        /// <param name="cell">The cell where the value originates from.</param>
        public virtual void SetProperty(object o, object val, ICell cell)
        {
            var v = GetPropertyValue(o, val, cell);
            Property.SetValue(o, v, null);
        }

        /// <summary>
        /// Gets the property value of the specified object.
        /// </summary>
        /// <param name="o">The object from which to get the property value.</param>
        /// <returns>The property value.</returns>
        public virtual object GetProperty(object o)
        {
            return Property.GetValue(o, null);
        }

        /// <summary>Specifies a method to use when setting the cell value from an object.</summary>
        /// <param name="setCell">The method to use when setting the cell value from an object.</param>
        /// <returns>The <see cref="ColumnInfo"/> object.</returns>
        public ColumnInfo SetCellUsing<TEntity>(Action<SetCellArgs<TEntity>> setCell) where TEntity : class
        {
            customCellSetter = (entity, cell, value) => setCell(new SetCellArgs<TEntity>(cell, (TEntity)entity, value));
            return this;
        }

        /// <summary>Specifies a method to use when setting the cell value from an object.</summary>
        /// <param name="setCell">The method to use when setting the cell value from an object.</param>
        /// <returns>The <see cref="ColumnInfo"/> object.</returns>
        public ColumnInfo SetCellUsing<TEntity, TValue>(Action<SetCellArgs<TEntity, TValue>> setCell) where TEntity : class
        {
            customCellSetter = (entity, cell, value) => setCell(new SetCellArgs<TEntity, TValue>(cell, (TEntity)entity, (TValue)value));
            return this;
        }

        /// <summary>Specifies a method to use when setting the cell value from an object.</summary>
        /// <param name="setCell">The method to use when setting the cell value from an object.</param>
        /// <returns>The <see cref="ColumnInfo"/> object.</returns>
        [Obsolete("This method is obsolete. Use SetCellUsing<TEntity>(Action<SetCellArgs<TEntity>>) instead.")]
        public ColumnInfo SetCellUsing(Action<ICell, object> setCell)
        {
            customCellSetter = (entity, cell, value) => setCell(cell, value);
            return this;
        }

        /// <summary>Specifies a method to use when setting the cell value from an object.</summary>
        /// <param name="setCell">The method to use when setting the cell value from an object.</param>
        /// <returns>The <see cref="ColumnInfo"/> object.</returns>
        [Obsolete("This method is obsolete. Use SetCellUsing<TEntity, TValue>(Action<SetCellArgs<TEntity, TValue>>) instead.")]
        public ColumnInfo SetCellUsing<T>(Action<ICell, T> setCell)
        {
            customCellSetter = (entity, cell, value) => setCell(cell, (T)value);
            return this;
        }

        /// <summary>Specifies a method to use when setting the property value from the cell value.</summary>
        /// <param name="setProp">The method to use when setting the property value from the cell value.</param>
        /// <returns>The <see cref="ColumnInfo"/> object.</returns>
        public ColumnInfo SetPropertyUsing<TEntity>(Func<SetPropertyArgs<TEntity>, object> setProp) where TEntity : class
        {
            SetProp = (entity, value, cell) => setProp(new SetPropertyArgs<TEntity>(cell, (TEntity)entity, value));
            return this;
        }

        /// <summary>Specifies a method to use when setting the property value from the cell value.</summary>
        /// <param name="setProp">The method to use when setting the property value from the cell value.</param>
        /// <returns>The <see cref="ColumnInfo"/> object.</returns>
        [Obsolete("This method is obsolete. Use SetPropertyUsing<TEntity>(Func<SetPropertyArgs<TEntity>, object>) instead.")]
        public ColumnInfo SetPropertyUsing(Func<object, object> setProp)
        {
            SetProp = (entity, value, cell) => setProp(value);
            return this;
        }

        /// <summary>Specifies a method to use when setting the property value from the cell value.</summary>
        /// <param name="setProp">The method to use when setting the property value from the cell value.</param>
        /// <returns>The <see cref="ColumnInfo"/> object.</returns>
        [Obsolete("This method is obsolete. Use SetPropertyUsing<TEntity>(Func<SetPropertyArgs<TEntity>, object>) instead.")]
        public ColumnInfo SetPropertyUsing(Func<object, ICell, object> setProp)
        {
            SetProp = (entity, value, cell) => setProp(value, cell);
            return this;
        }

        /// <summary>Specifies a method to use when setting the property value from the cell value.</summary>
        /// <param name="setProp">The method to use when setting the property value from the cell value.</param>
        /// <returns>The <see cref="ColumnInfo"/> object.</returns>
        [Obsolete("This method is obsolete. Use SetPropertyUsing<TEntity>(Func<SetPropertyArgs<TEntity>, object>) instead.")]
        public ColumnInfo SetPropertyUsing(Func<object, object, ICell, object> setProp)
        {
            SetProp = setProp;
            return this;
        }

        /// <summary>Specifies a method to use when setting the property value from the cell value.</summary>
        /// <param name="setProp">The method to use when setting the property value from the cell value.</param>
        /// <returns>The <see cref="ColumnInfo"/> object.</returns>
        [Obsolete("This method is obsolete. Use SetPropertyUsing<TEntity>(Func<SetPropertyArgs<TEntity>, object>) instead.")]
        public ColumnInfo SetPropertyUsing<T>(Func<T, object, object> setProp)
        {
            SetProp = (entity, value, cell) => setProp((T)entity, value);
            return this;
        }

        /// <summary>Specifies a method to use when setting the property value from the cell value.</summary>
        /// <param name="setProp">The method to use when setting the property value from the cell value.</param>
        /// <returns>The <see cref="ColumnInfo"/> object.</returns>
        [Obsolete("This method is obsolete. Use SetPropertyUsing<TEntity>(Func<SetPropertyArgs<TEntity>, object>) instead.")]
        public ColumnInfo SetPropertyUsing<T>(Func<T, object, ICell, object> setProp)
        {
            SetProp = (entity, value, cell) => setProp((T)entity, value, cell);
            return this;
        }

        /// <summary>Selects formula results to be mapped instead of the formula itself.</summary>
        /// <returns>The <see cref="ColumnInfo"/> object.</returns>
        public ColumnInfo AsFormulaResult()
        {
            FormulaResult = true;
            return this;
        }

        /// <summary>Selects the property to be saved as a formula.</summary>
        /// <returns>The <see cref="ColumnInfo"/> object.</returns>
        public ColumnInfo AsFormula()
        {
            Formula = true;
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
        /// <param name="direction">Data direction</param>
        public ColumnInfo(PropertyInfo propertyInfo, MappingDirections direction = MappingDirections.Both)
        {
            Property = propertyInfo;
            Directions = direction;
            defaultCellSetter = GenerateCellSetter();
            if (DateTypes.Contains(PropertyType.FullName))
                BuiltinFormat = 0x16; // "m/d/yy h:mm"
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ColumnInfo"/> class.
        /// </summary>
        protected ColumnInfo() { }

        /// <summary>Selects the property to be unidirectional from Excel to Object.</summary>
        /// <returns>The <see cref="ColumnInfo"/> object.</returns>
        public ColumnInfo FromExcelOnly()
        {
            Directions = MappingDirections.ExcelToObject;
            return this;
        }

        /// <summary>Selects the property to be unidirectional from Excel to Object.</summary>
        /// <returns>The <see cref="ColumnInfo"/> object.</returns>
        public ColumnInfo ToExcelOnly()
        {
            Directions = MappingDirections.ObjectToExcel;
            return this;
        }

        internal void ChangeSetterType(Type newType)
        {
            SetPropertyType(newType);
            defaultCellSetter = GenerateCellSetter();
            if (DateTypes.Contains(PropertyType.FullName))
                BuiltinFormat = 0x16; // "m/d/yy h:mm"
            else
            {
                BuiltinFormat = 0;
            }
        }
    }

    /// <summary>
    /// Describes the mapping of an <see cref="ExpandoObject"/>'s property to a cell.
    /// </summary>
    public class DynamicColumnInfo : ColumnInfo
    {
        /// <summary>
        /// Gets or sets the column index.
        /// </summary>
        public int Index { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="DynamicColumnInfo"/> class.
        /// </summary>
        /// <param name="index">The column index.</param>
        /// <param name="name">The column name.</param>
        public DynamicColumnInfo(int index, string name)
        {
            Index = index;
            Name = name;
            FormulaResult = true;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DynamicColumnInfo"/> class.
        /// </summary>
        /// <param name="name">The column name.</param>
        /// <param name="t">The type of the column.</param>
        public DynamicColumnInfo(string name, Type t)
        {
            Name = name;
            SetPropertyType(t);
            defaultCellSetter = GenerateCellSetter();
            if (DateTypes.Contains(PropertyType.FullName))
                BuiltinFormat = 0x16; // "m/d/yy h:mm"
        }

        /// <summary>
        /// Sets the property of the specified object to the specified value.
        /// </summary>
        /// <param name="o">The object whose property to set.</param>
        /// <param name="val">The value.</param>
        /// <param name="cell">The cell where the value originates from.</param>
        public override void SetProperty(object o, object val, ICell cell)
        {
            var expando = (IDictionary<string, object>)o;
            if (!string.IsNullOrEmpty(Name))
                expando.Add(Name, val);
        }

        /// <summary>
        /// Gets the property value of the specified object.
        /// </summary>
        /// <param name="o">The o.</param>
        /// <returns>The property value.</returns>
        public override object GetProperty(object o)
        {
            return ((IDictionary<string, object>)o)[Name];
        }
    }
}
