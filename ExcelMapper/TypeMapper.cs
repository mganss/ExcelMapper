using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;

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

        static HashSet<Type> NumericTypes = new HashSet<Type>
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
                    else
                        c.SetCellValue(o.ToString());
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

        private void SetCellFormat(ICell c, short defaultFormat = 0)
        {
            var wb = c.Row.Sheet.Workbook;
            var cs = wb.CreateCellStyle();
            if (CustomFormat != null)
                cs.DataFormat = wb.CreateDataFormat().GetFormat(CustomFormat);
            else
                cs.DataFormat = BuiltinFormat != 0 ? BuiltinFormat : defaultFormat;
            c.CellStyle = cs;
        }

        /// <summary>
        /// Sets the property of the specified object to the specified value.
        /// </summary>
        /// <param name="o">The object whose property to set.</param>
        /// <param name="val">The value.</param>
        public void SetProperty(object o, object val)
        {
            object v;
            if (SetProp != null)
                v = SetProp(val);
            else if (IsNullable && (val == null || (val as string) == ""))
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
    /// Maps a <see cref="Type"/>'s properties to columns in an Excel sheet.
    /// </summary>
    public class TypeMapper
    {
        Type Type { get; set; }

        /// <summary>
        /// Gets or sets the columns by name.
        /// </summary>
        /// <value>
        /// The dictionary of columns by name.
        /// </value>
        public Dictionary<string, ColumnInfo> ColumnsByName { get; set; } = new Dictionary<string, ColumnInfo>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Gets or sets the columns by index.
        /// </summary>
        /// <value>
        /// The dictionary of columns by index.
        /// </value>
        public Dictionary<int, ColumnInfo> ColumnsByIndex { get; set; } = new Dictionary<int, ColumnInfo>();

        /// <summary>
        /// Creates a <see cref="TypeMapper"/> object from the specified type.
        /// </summary>
        /// <param name="type">The type.</param>
        /// <returns>A <see cref="TypeMapper"/> object.</returns>
        public static TypeMapper Create(Type type)
        {
            var typeMapper = new TypeMapper { Type = type };
            typeMapper.Analyze();
            return typeMapper;
        }

        void Analyze()
        {
            foreach (var prop in Type.GetProperties(BindingFlags.Instance | BindingFlags.Public))
            {
                if (!(Attribute.GetCustomAttribute(prop, typeof(IgnoreAttribute)) is IgnoreAttribute ignoreAttribute))
                {
                    var ci = new ColumnInfo(prop);

                    if (Attribute.GetCustomAttribute(prop, typeof(ColumnAttribute)) is ColumnAttribute columnAttribute)
                    {
                        if (!string.IsNullOrEmpty(columnAttribute.Name))
                            ColumnsByName[columnAttribute.Name] = ci;
                        else
                            ColumnsByName[prop.Name] = ci;

                        if (columnAttribute.Index > 0)
                            ColumnsByIndex[columnAttribute.Index - 1] = ci;
                    }
                    else
                        ColumnsByName[prop.Name] = ci;

                    if (Attribute.GetCustomAttribute(prop, typeof(DataFormatAttribute)) is DataFormatAttribute dataFormatAttribute)
                    {
                        ci.BuiltinFormat = dataFormatAttribute.BuiltinFormat;
                        ci.CustomFormat = dataFormatAttribute.CustomFormat;
                    }
                }
            }
        }

        /// <summary>
        /// Gets the <see cref="ColumnInfo"/> for the specified column name.
        /// </summary>
        /// <param name="name">The column name.</param>
        /// <returns>A <see cref="ColumnInfo"/> object or null if no <see cref="ColumnInfo"/> exists for the specified column name.</returns>
        public ColumnInfo GetColumnByName(string name)
        {
            ColumnsByName.TryGetValue(name, out ColumnInfo col);
            return col;
        }

        /// <summary>
        /// Gets the <see cref="ColumnInfo"/> for the specified column index.
        /// </summary>
        /// <param name="index">The column index.</param>
        /// <returns>A <see cref="ColumnInfo"/> object or null if no <see cref="ColumnInfo"/> exists for the specified column index.</returns>
        public ColumnInfo GetColumnByIndex(int index)
        {
            ColumnsByIndex.TryGetValue(index, out ColumnInfo col);
            return col;
        }
    }
}
