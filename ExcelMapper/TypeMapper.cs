using System;
using System.Collections.Generic;
using System.Reflection;

namespace Ganss.Excel
{

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
        public Dictionary<string, List<ColumnInfo>> ColumnsByName { get; set; } = new Dictionary<string, List<ColumnInfo>>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Gets or sets the columns by index.
        /// </summary>
        /// <value>
        /// The dictionary of columns by index.
        /// </value>
        public Dictionary<int, List<ColumnInfo>> ColumnsByIndex { get; set; } = new Dictionary<int, List<ColumnInfo>>();

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
                if (!(Attribute.GetCustomAttribute(prop, typeof(IgnoreAttribute)) is IgnoreAttribute))
                {
                    var ci = new ColumnInfo(prop); // Both direction

                    if (Attribute.GetCustomAttribute(prop, typeof(ColumnAttribute)) is ColumnAttribute columnAttribute)
                    {
                        if (!string.IsNullOrEmpty(columnAttribute.Name))
                        {
                            if (!ColumnsByName.ContainsKey(columnAttribute.Name))
                                ColumnsByName.Add(columnAttribute.Name, new List<ColumnInfo>());

                            ColumnsByName[columnAttribute.Name].Add(ci);
                        }
                        else if (!ColumnsByName.ContainsKey(prop.Name))
                            ColumnsByName.Add(prop.Name, new List<ColumnInfo>() { ci });

                        if (columnAttribute.Index > 0)
                        {
                            var idx = columnAttribute.Index - 1;
                            if (!ColumnsByIndex.ContainsKey(idx))
                                ColumnsByIndex.Add(idx, new List<ColumnInfo>());

                            ColumnsByIndex[idx].Add(ci);
                        }
                    }
                    else if (!ColumnsByName.ContainsKey(prop.Name))
                        ColumnsByName.Add(prop.Name, new List<ColumnInfo>() { ci });

                    if (Attribute.GetCustomAttribute(prop, typeof(FromExcelOnlyAttribute)) is FromExcelOnlyAttribute)
                        ci.Direction = ColumnInfoDirections.Cell2Prop;

                    if (Attribute.GetCustomAttribute(prop, typeof(ToExcelOnlyAttribute)) is ToExcelOnlyAttribute)
                        ci.Direction = ColumnInfoDirections.Prop2Cell;

                    if (Attribute.GetCustomAttribute(prop, typeof(DataFormatAttribute)) is DataFormatAttribute dataFormatAttribute)
                    {
                        ci.BuiltinFormat = dataFormatAttribute.BuiltinFormat;
                        ci.CustomFormat = dataFormatAttribute.CustomFormat;
                    }

                    if (Attribute.GetCustomAttribute(prop, typeof(FormulaResultAttribute)) is FormulaResultAttribute)
                        ci.FormulaResult = true;

                    if (Attribute.GetCustomAttribute(prop, typeof(JsonAttribute)) is JsonAttribute)
                        ci.Json = true;
                }
            }
        }

        /// <summary>
        /// Gets the <see cref="ColumnInfo"/> for the specified column name.
        /// </summary>
        /// <param name="name">The column name.</param>
        /// <returns>A <see cref="ColumnInfo"/> object or null if no <see cref="ColumnInfo"/> exists for the specified column name.</returns>
        public List<ColumnInfo> GetColumnByName(string name)
        {
            ColumnsByName.TryGetValue(name, out List<ColumnInfo> col);
            return col;
        }

        /// <summary>
        /// Gets the <see cref="ColumnInfo"/> for the specified column index.
        /// </summary>
        /// <param name="index">The column index.</param>
        /// <returns>A <see cref="ColumnInfo"/> object or null if no <see cref="ColumnInfo"/> exists for the specified column index.</returns>
        public List<ColumnInfo> GetColumnByIndex(int index)
        {
            ColumnsByIndex.TryGetValue(index, out List<ColumnInfo> col);
            return col;
        }
    }
}
