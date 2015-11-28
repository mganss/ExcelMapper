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
        /// <summary>
        /// Gets or sets the property.
        /// </summary>
        /// <value>
        /// The property.
        /// </value>
        public PropertyInfo Property { get; set; }

        /// <summary>
        /// Gets or sets the cell setter.
        /// </summary>
        /// <value>
        /// The cell setter.
        /// </value>
        public Action<ICell, object> SetCell { get; set; }

        static HashSet<Type> NumericTypes = new HashSet<Type>
        {
            typeof(decimal),
            typeof(byte), typeof(sbyte),
            typeof(short), typeof(ushort),
            typeof(int), typeof(uint),
            typeof(long), typeof(ulong),
            typeof(float), typeof(double)
        };

        Action<ICell, object> GenerateCellSetter(Type type)
        {
            if (type == typeof(DateTime))
            {
                return (c, o) => c.SetCellValue(Convert.ToDateTime(o));
            }
            else if (type == typeof(bool))
            {
                return (c, o) => c.SetCellValue(Convert.ToBoolean(o));
            }
            else if (NumericTypes.Contains(type))
            {
                return (c, o) => c.SetCellValue(Convert.ToDouble(o));
            }
            else
            {
                return (c, o) => c.SetCellValue(o.ToString());
            }
        }

        /// <summary>
        /// Sets the property of the specified object to the specified value.
        /// </summary>
        /// <param name="o">The object whose property to set.</param>
        /// <param name="val">The value.</param>
        public void SetProperty(object o, object val)
        {
            var v = Convert.ChangeType(val, Property.PropertyType, CultureInfo.InvariantCulture);
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
            SetCell = GenerateCellSetter(propertyInfo.PropertyType);
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
                var columnAttribute = Attribute.GetCustomAttribute(prop, typeof(ColumnAttribute)) as ColumnAttribute;
                if (columnAttribute != null)
                {
                    if (!string.IsNullOrEmpty(columnAttribute.Name))
                        ColumnsByName[columnAttribute.Name] = new ColumnInfo(prop);
                    else
                        ColumnsByIndex[columnAttribute.Index - 1] = new ColumnInfo(prop);
                }
                else
                    ColumnsByName[prop.Name] = new ColumnInfo(prop);
            }
        }

        /// <summary>
        /// Gets the <see cref="ColumnInfo"/> for the specified column name.
        /// </summary>
        /// <param name="name">The column name.</param>
        /// <returns>A <see cref="ColumnInfo"/> object.</returns>
        /// <exception cref="System.ArgumentException">No property exists for the specified column name</exception>
        public ColumnInfo GetColumnByName(string name)
        {
            ColumnInfo col;
            if (!ColumnsByName.TryGetValue(name, out col))
                throw new ArgumentException($"No property for column name {name} on type {Type.Name}");
            return col;
        }

        /// <summary>
        /// Gets the <see cref="ColumnInfo"/> for the specified column index.
        /// </summary>
        /// <param name="index">The column index.</param>
        /// <returns>A <see cref="ColumnInfo"/> object.</returns>
        /// <exception cref="System.ArgumentException">No property for the specified column index</exception>
        public ColumnInfo GetColumnByIndex(int index)
        {
            ColumnInfo col;
            if (!ColumnsByIndex.TryGetValue(index, out col))
                throw new ArgumentException($"No property for column index {index} on type {Type.Name}");
            return col;
        }
    }
}
