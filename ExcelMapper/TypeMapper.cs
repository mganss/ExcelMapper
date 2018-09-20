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
                    {
                        var d = (DateTime)o;
                        c.SetCellValue(d);

                        if (BuiltinFormat != 0 || CustomFormat != null || c.CellStyle.DataFormat == 0)
                            SetCellFormat(c, 0x16); // "m/d/yy h:mm"
                    }
                };
            }
            else if (PropertyType == typeof(bool))
            {
                if (IsNullable)
                    return (c, o) =>
                    {
                        if (o == null) c.SetCellValue((string)null); else c.SetCellValue((bool)o);
                    };
                else
                    return (c, o) => c.SetCellValue((bool)o);
            }
            else if (NumericTypes.Contains(PropertyType))
            {
                if (IsNullable)
                    return (c, o) =>
                    {
                        if (o == null) c.SetCellValue((string)null);
                        else
                        {
                            c.SetCellValue(Convert.ToDouble(o));
                            if (BuiltinFormat != 0 || CustomFormat != null)
                                SetCellFormat(c);
                        }
                    };
                else
                    return (c, o) =>
                    {
                        c.SetCellValue(Convert.ToDouble(o));
                        if (BuiltinFormat != 0 || CustomFormat != null)
                            SetCellFormat(c);
                    };
            }
            else
            {
                return (c, o) =>
                {
                    if (o == null) c.SetCellValue((string)null); else c.SetCellValue(o.ToString());
                };
            }
        }

        private void SetCellFormat(ICell c, short defaultFormat = 0)
        {
            var wb = c.Row.Sheet.Workbook;
            var cs = wb.CreateCellStyle();
            cs.DataFormat = CustomFormat != null ? wb.CreateDataFormat().GetFormat(CustomFormat) : BuiltinFormat != 0 ? BuiltinFormat : defaultFormat;
            c.CellStyle = cs;
        }

        /// <summary>
        /// Sets the property of the specified object to the specified value.
        /// </summary>
        /// <param name="o">The object whose property to set.</param>
        /// <param name="val">The value.</param>
        public void SetProperty(object o, object val)
        {
            var v = IsNullable && (val == null || (val as string) == "") ? null : Convert.ChangeType(val, PropertyType, CultureInfo.InvariantCulture);
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

        /// <summary>
        /// Initializes a new instance of the <see cref="ColumnInfo"/> class.
        /// </summary>
        /// <param name="propertyInfo">The property information.</param>
        public ColumnInfo(PropertyInfo propertyInfo)
        {
            Property = propertyInfo;
            SetCell = GenerateCellSetter();
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
