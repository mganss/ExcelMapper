using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
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
        /// Gets or sets the Before Mapping action.
        /// </summary>
        internal ActionInvoker BeforeMappingActionInvoker { get; set; }

        /// <summary>
        /// Gets or sets the After Mapping action.
        /// </summary>
        internal ActionInvoker AfterMappingActionInvoker { get; set; }

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

        /// <summary>
        /// Creates a <see cref="TypeMapper"/> object from a list of cells.
        /// </summary>
        /// <param name="columns">The cells.</param>
        /// <returns>A <see cref="TypeMapper"/> object.</returns>
        public static TypeMapper Create(IEnumerable<ICell> columns)
        {
            var typeMapper = new TypeMapper();

            foreach (var col in columns)
            {
                var index = col.ColumnIndex;
                var name = col.StringCellValue;
                var columnInfo = new DynamicColumnInfo(index, name);
                typeMapper.ColumnsByIndex.Add(index, new List<ColumnInfo> { columnInfo });
                if (!typeMapper.ColumnsByName.TryGetValue(name, out var columnInfos))
                    typeMapper.ColumnsByName.Add(name, new List<ColumnInfo> { columnInfo });
                else
                    columnInfos.Add(columnInfo);
            }

            return typeMapper;
        }

        /// <summary>
        /// Creates a <see cref="TypeMapper"/> object from an <see cref="ExpandoObject"/> object.
        /// </summary>
        /// <param name="o">The <see cref="ExpandoObject"/> object.</param>
        /// <returns>A <see cref="TypeMapper"/> object.</returns>
        public static TypeMapper Create(ExpandoObject o)
        {
            var typeMapper = new TypeMapper();
            var l = o.ToList();

            for (int i = 0; i < o.Count(); i++)
            {
                var prop = l[i];
                var name = prop.Key;
                var columnInfo = new DynamicColumnInfo(prop.Key, prop.Value.GetType());

                if (!int.TryParse(prop.Key, out _))
                {
                    if ((i % 2) == 0 || !int.TryParse(l[i - 1].Key, out var ix))
                        ix = i;
                    typeMapper.ColumnsByIndex.Add(ix, new List<ColumnInfo> { columnInfo });
                    if (!typeMapper.ColumnsByName.TryGetValue(name, out var columnInfos))
                        typeMapper.ColumnsByName.Add(name, new List<ColumnInfo> { columnInfo });
                    else
                        columnInfos.Add(columnInfo);
                }
            }

            return typeMapper;
        }

        void Analyze()
        {
            foreach (var prop in Type.GetProperties(BindingFlags.Instance | BindingFlags.Public))
            {
                if (!(Attribute.GetCustomAttribute(prop, typeof(IgnoreAttribute)) is IgnoreAttribute))
                {
                    var ci = new ColumnInfo(prop); // Both direction

                    var attribs = Attribute.GetCustomAttributes(prop, typeof(ColumnAttribute)).Cast<ColumnAttribute>();
                    if (attribs.Any())
                    {
                        foreach (var columnAttribute in attribs)
                        {
                            ci = new ColumnInfo(prop); // Both direction
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

                            ci.Directions = columnAttribute.Directions;
                        }
                    }
                    else if (!ColumnsByName.ContainsKey(prop.Name))
                        ColumnsByName.Add(prop.Name, new List<ColumnInfo>() { ci });

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
