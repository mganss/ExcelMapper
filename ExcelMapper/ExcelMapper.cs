using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;

namespace Ganss.Excel
{
    /// <summary>
    /// Map objects to Excel files.
    /// </summary>
    public class ExcelMapper
    {
        /// <summary>
        /// Gets or sets the <see cref="TypeMapper"/> factory.
        /// Default is a static <see cref="Ganss.Excel.TypeMapperFactory"/> object that caches <see cref="TypeMapper"/>s statically across <see cref="ExcelMapper"/> instances.
        /// </summary>
        /// <value>
        /// The <see cref="TypeMapper"/> factory.
        /// </value>
        public ITypeMapperFactory TypeMapperFactory { get; set; } = DefaultTypeMapperFactory;

        /// <summary>
        /// Gets or sets a value indicating whether the Excel file contains a header row of column names. Default is <c>true</c>.
        /// </summary>
        /// <value>
        ///   <c>true</c> if the Excel file contains a header row; otherwise, <c>false</c>.
        /// </value>
        public bool HeaderRow { get; set; } = true;

        /// <summary>
        /// Gets or sets a value indicating whether to track objects read from the Excel file. Default is true.
        /// If object tracking is enabled, the <see cref="ExcelMapper"/> object keeps track of objects it yields through the Fetch() methods.
        /// You can then modify these objects and save them back to an Excel file without having to specify the list of objects to save.
        /// </summary>
        /// <value>
        ///   <c>true</c> if object tracking is enabled; otherwise, <c>false</c>.
        /// </value>
        public bool TrackObjects { get; set; } = true;

        /// <summary>
        /// Gets or sets a value indicating whether to skip blank rows when reading from Excel files. Default is true.
        /// </summary>
        /// <value>
        ///   <c>true</c> if blank lines are skipped; otherwise, <c>false</c>.
        /// </value>
        public bool SkipBlankRows { get; set; } = true;

        Dictionary<string, Dictionary<int, object>> Objects { get; set; } = new Dictionary<string, Dictionary<int, object>>();
        IWorkbook Workbook { get; set; }

        static readonly TypeMapperFactory DefaultTypeMapperFactory = new TypeMapperFactory();

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelMapper"/> class.
        /// </summary>
        public ExcelMapper() { }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelMapper"/> class.
        /// </summary>
        /// <param name="file">The path to the Excel file.</param>
        public ExcelMapper(string file)
        {
            Workbook = WorkbookFactory.Create(file);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelMapper"/> class.
        /// </summary>
        /// <param name="stream">The stream the Excel file is read from.</param>
        public ExcelMapper(Stream stream)
        {
            Workbook = WorkbookFactory.Create(stream);
        }

        /// <summary>
        /// Fetches objects from the specified sheet name.
        /// </summary>
        /// <typeparam name="T">The type of objects the Excel file is mapped to.</typeparam>
        /// <param name="file">The path to the Excel file.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <returns>The objects read from the Excel file.</returns>
        public IEnumerable<T> Fetch<T>(string file, string sheetName) where T : new()
        {
            Workbook = WorkbookFactory.Create(file);
            return Fetch<T>(sheetName);
        }

        /// <summary>
        /// Fetches objects from the specified sheet index.
        /// </summary>
        /// <typeparam name="T">The type of objects the Excel file is mapped to.</typeparam>
        /// <param name="file">The path to the Excel file.</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <returns>The objects read from the Excel file.</returns>
        public IEnumerable<T> Fetch<T>(string file, int sheetIndex) where T : new()
        {
            Workbook = WorkbookFactory.Create(file);
            return Fetch<T>(sheetIndex);
        }

        /// <summary>
        /// Fetches objects from the specified sheet name.
        /// </summary>
        /// <typeparam name="T">The type of objects the Excel file is mapped to.</typeparam>
        /// <param name="stream">The stream the Excel file is read from.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <returns>The objects read from the Excel file.</returns>
        public IEnumerable<T> Fetch<T>(Stream stream, string sheetName) where T : new()
        {
            Workbook = WorkbookFactory.Create(stream);
            return Fetch<T>(sheetName);
        }

        /// <summary>
        /// Fetches objects from the specified sheet index.
        /// </summary>
        /// <typeparam name="T">The type of objects the Excel file is mapped to.</typeparam>
        /// <param name="stream">The stream the Excel file is read from.</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <returns>The objects read from the Excel file.</returns>
        public IEnumerable<T> Fetch<T>(Stream stream, int sheetIndex) where T : new()
        {
            Workbook = WorkbookFactory.Create(stream);
            return Fetch<T>(sheetIndex);
        }

        /// <summary>
        /// Fetches objects from the specified sheet name.
        /// </summary>
        /// <typeparam name="T">The type of objects the Excel file is mapped to.</typeparam>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <returns>The objects read from the Excel file.</returns>
        public IEnumerable<T> Fetch<T>(string sheetName) where T : new()
        {
            var sheet = Workbook.GetSheet(sheetName);
            return Fetch<T>(sheet);
        }

        /// <summary>
        /// Fetches objects from the specified sheet index.
        /// </summary>
        /// <typeparam name="T">The type of objects the Excel file is mapped to.</typeparam>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <returns>The objects read from the Excel file.</returns>
        public IEnumerable<T> Fetch<T>(int sheetIndex = 0) where T : new()
        {
            var sheet = Workbook.GetSheetAt(sheetIndex);
            return Fetch<T>(sheet);
        }

        IEnumerable<T> Fetch<T>(ISheet sheet) where T : new()
        {
            var typeMapper = TypeMapperFactory.Create(typeof(T));
            var columns = sheet.GetRow(0).Cells
                .Where(c => !HeaderRow || (c.CellType == CellType.String && !string.IsNullOrWhiteSpace(c.StringCellValue)))
                .Select(c => new { c.ColumnIndex, ColumnInfo = HeaderRow ? typeMapper.GetColumnByName(c.StringCellValue) : typeMapper.GetColumnByIndex(c.ColumnIndex) })
                .Where(c => c.ColumnInfo != null)
                .ToDictionary(c => c.ColumnIndex, c => c.ColumnInfo);
            var i = HeaderRow ? 1 : 0;
            IRow row = null;

            if (TrackObjects) Objects[sheet.SheetName] = new Dictionary<int, object>();

            while ((row = sheet.GetRow(i)) != null)
            {
                if (!SkipBlankRows || row.Cells.Any(c => c.CellType != CellType.Blank))
                {
                    var o = new T();

                    foreach (var col in columns)
                    {
                        var cell = row.GetCell(col.Key);
                        if (cell != null)
                            col.Value.SetProperty(o, GetCellValue(cell, col.Value));
                    }

                    if (TrackObjects) Objects[sheet.SheetName][i] = o;

                    yield return o;
                }

                i++;
            }
        }

        /// <summary>
        /// Saves the specified objects to the specified Excel file.
        /// </summary>
        /// <typeparam name="T">The type of objects to save.</typeparam>
        /// <param name="file">The path to the Excel file.</param>
        /// <param name="objects">The objects to save.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="xlsx">if set to <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
        public void Save<T>(string file, IEnumerable<T> objects, string sheetName, bool xlsx = true)
        {
            using (var fs = File.Open(file, FileMode.Create, FileAccess.Write))
                Save(fs, objects, sheetName, xlsx);
        }

        /// <summary>
        /// Saves the specified objects to the specified Excel file.
        /// </summary>
        /// <typeparam name="T">The type of objects to save.</typeparam>
        /// <param name="file">The path to the Excel file.</param>
        /// <param name="objects">The objects to save.</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <param name="xlsx">if set to <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
        public void Save<T>(string file, IEnumerable<T> objects, int sheetIndex = 0, bool xlsx = true)
        {
            using (var fs = File.Open(file, FileMode.Create, FileAccess.Write))
                Save(fs, objects, sheetIndex, xlsx);
        }

        /// <summary>
        /// Saves the specified objects to the specified stream.
        /// </summary>
        /// <typeparam name="T">The type of objects to save.</typeparam>
        /// <param name="stream">The stream to save the objects to.</param>
        /// <param name="objects">The objects to save.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="xlsx">if set to <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
        public void Save<T>(Stream stream, IEnumerable<T> objects, string sheetName, bool xlsx = true)
        {
            if (Workbook == null)
                Workbook = xlsx ? (IWorkbook)new XSSFWorkbook() : (IWorkbook)new HSSFWorkbook();
            var sheet = Workbook.GetSheet(sheetName);
            if (sheet == null) sheet = Workbook.CreateSheet(sheetName);
            Save(stream, sheet, objects);
        }

        /// <summary>
        /// Saves the specified objects to the specified stream.
        /// </summary>
        /// <typeparam name="T">The type of objects to save.</typeparam>
        /// <param name="stream">The stream to save the objects to.</param>
        /// <param name="objects">The objects to save.</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <param name="xlsx">if set to <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
        public void Save<T>(Stream stream, IEnumerable<T> objects, int sheetIndex = 0, bool xlsx = true)
        {
            if (Workbook == null)
                Workbook = xlsx ? (IWorkbook)new XSSFWorkbook() : (IWorkbook)new HSSFWorkbook();
            ISheet sheet;
            if (Workbook.NumberOfSheets > sheetIndex)
                sheet = Workbook.GetSheetAt(sheetIndex);
            else
                sheet = Workbook.CreateSheet();
            Save(stream, sheet, objects);
        }

        /// <summary>
        /// Saves tracked objects to the specified Excel file.
        /// </summary>
        /// <param name="file">The path to the Excel file.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="xlsx">if set to <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
        public void Save(string file, string sheetName, bool xlsx = true)
        {
            using (var fs = File.Open(file, FileMode.Create, FileAccess.Write))
                Save(fs, sheetName, xlsx);
        }

        /// <summary>
        /// Saves tracked objects to the specified Excel file.
        /// </summary>
        /// <param name="file">The path to the Excel file.</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <param name="xlsx">if set to <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
        public void Save(string file, int sheetIndex = 0, bool xlsx = true)
        {
            using (var fs = File.Open(file, FileMode.Create, FileAccess.Write))
                Save(fs, sheetIndex, xlsx);
        }

        /// <summary>
        /// Saves tracked objects to the specified stream.
        /// </summary>
        /// <param name="stream">The stream to save the objects to.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="xlsx">if set to <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
        public void Save(Stream stream, string sheetName, bool xlsx = true)
        {
            if (Workbook == null)
                Workbook = xlsx ? (IWorkbook)new XSSFWorkbook() : (IWorkbook)new HSSFWorkbook();
            var sheet = Workbook.GetSheet(sheetName);
            if (sheet == null) sheet = Workbook.CreateSheet(sheetName);
            Save(stream, sheet);
        }

        /// <summary>
        /// Saves tracked objects to the specified stream.
        /// </summary>
        /// <param name="stream">The stream to save the objects to.</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <param name="xlsx">if set to <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
        public void Save(Stream stream, int sheetIndex = 0, bool xlsx = true)
        {
            if (Workbook == null)
                Workbook = xlsx ? (IWorkbook)new XSSFWorkbook() : (IWorkbook)new HSSFWorkbook();
            var sheet = Workbook.GetSheetAt(sheetIndex);
            if (sheet == null) sheet = Workbook.CreateSheet();
            Save(stream, sheet);
        }

        void Save(Stream stream, ISheet sheet)
        {
            var objects = Objects[sheet.SheetName];
            var typeMapper = TypeMapperFactory.Create(objects.First().Value.GetType());
            var columnsByIndex = GetColumns(sheet, typeMapper);

            foreach (var o in objects)
            {
                var i = o.Key;
                var row = sheet.GetRow(i);
                if (row == null) row = sheet.CreateRow(i);

                foreach (var col in columnsByIndex)
                {
                    var cell = row.GetCell(col.Key, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    col.Value.SetCell(cell, col.Value.GetProperty(o.Value));
                }
            }

            Workbook.Write(stream);
        }

        void Save<T>(Stream stream, ISheet sheet, IEnumerable<T> objects)
        {
            var typeMapper = TypeMapperFactory.Create(typeof(T));
            var columnsByIndex = GetColumns(sheet, typeMapper);
            var i = HeaderRow ? 1 : 0;

            foreach (var o in objects)
            {
                var row = sheet.GetRow(i);
                if (row == null) row = sheet.CreateRow(i);

                foreach (var col in columnsByIndex)
                {
                    var cell = row.GetCell(col.Key, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    col.Value.SetCell(cell, col.Value.GetProperty(o));
                }

                i++;
            }

            if (SkipBlankRows)
            {
                while (i <= sheet.LastRowNum)
                {
                    var row = sheet.GetRow(i);
                    while (row.Cells.Any())
                        row.RemoveCell(row.GetCell(row.FirstCellNum));
                    i++;
                }
            }

            Workbook.Write(stream);
        }

        Dictionary<int, ColumnInfo> GetColumns(ISheet sheet, TypeMapper typeMapper)
        {
            Dictionary<int, ColumnInfo> columnsByIndex;
            if (HeaderRow)
            {
                var columnsByName = typeMapper.ColumnsByName;
                var headerRow = sheet.GetRow(0);

                if (headerRow == null)
                {
                    var j = 0;
                    columnsByIndex = new Dictionary<int, ColumnInfo>();
                    headerRow = sheet.CreateRow(0);

                    foreach (var getter in columnsByName)
                    {
                        var cell = headerRow.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                        cell.SetCellValue(getter.Key);
                        columnsByIndex[j] = getter.Value;
                        j++;
                    }
                }
                else
                {
                    columnsByIndex = headerRow.Cells
                        .Where(c => c.CellType == CellType.String && !string.IsNullOrWhiteSpace(c.StringCellValue))
                        .Select(c => new { c.ColumnIndex, ColumnInfo = HeaderRow ? typeMapper.GetColumnByName(c.StringCellValue) : typeMapper.GetColumnByIndex(c.ColumnIndex) })
                        .Where(c => c.ColumnInfo != null)
                        .ToDictionary(c => c.ColumnIndex, c => c.ColumnInfo);
                }
            }
            else
            {
                columnsByIndex = typeMapper.ColumnsByIndex;
            }

            return columnsByIndex;
        }

        object GetCellValue(ICell cell, ColumnInfo targetColumn)
        {
            var cellType = cell.CellType == CellType.Formula && targetColumn.PropertyType != typeof(string) ? cell.CachedFormulaResultType : cell.CellType;

            switch (cellType)
            {
                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(cell))
                        return cell.DateCellValue;
                    else
                        return cell.NumericCellValue;
                case CellType.Formula:
                    return cell.CellFormula;
                case CellType.Boolean:
                    return cell.BooleanCellValue;
                case CellType.Error:
                    return cell.ErrorCellValue;
                case CellType.Unknown:
                case CellType.Blank:
                case CellType.String:
                default:
                    return cell.StringCellValue;
            }
        }

        /// <summary>
        /// Adds a mapping from a column name to a property.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="columnName">Name of the column.</param>
        /// <param name="propertyExpression">The property expression.</param>
        public void AddMapping<T>(string columnName, Expression<Func<T,object>> propertyExpression)
        {
            var typeMapper = TypeMapperFactory.Create(typeof(T));
            var prop = GetPropertyInfo(propertyExpression);
            typeMapper.ColumnsByName[columnName] = new ColumnInfo(prop);
        }

        /// <summary>
        /// Adds a mapping from a column index to a property.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="columnIndex">Index of the column.</param>
        /// <param name="propertyExpression">The property expression.</param>
        public void AddMapping<T>(int columnIndex, Expression<Func<T,object>> propertyExpression)
        {
            var typeMapper = TypeMapperFactory.Create(typeof(T));
            var prop = GetPropertyInfo(propertyExpression);
            typeMapper.ColumnsByIndex[columnIndex] = new ColumnInfo(prop);
        }

        static PropertyInfo GetPropertyInfo<T>(Expression<Func<T, object>> propertyExpression)
        {
            var exp = (LambdaExpression)propertyExpression;
            var mExp = (exp.Body.NodeType == ExpressionType.MemberAccess) ?
                (MemberExpression)exp.Body :
                (MemberExpression)((UnaryExpression)exp.Body).Operand;
            return (PropertyInfo)mExp.Member;
        }

        /// <summary>
        /// Adds a mapping from a column name to a property.
        /// </summary>
        /// <param name="t">The type that contains the property to map to.</param>
        /// <param name="columnName">Name of the column.</param>
        /// <param name="propertyName">Name of the property.</param>
        public void AddMapping(Type t, string columnName, string propertyName)
        {
            var typeMapper = TypeMapperFactory.Create(t);
            var prop = t.GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public);
            typeMapper.ColumnsByName[columnName] = new ColumnInfo(prop);
        }

        /// <summary>
        /// Adds a mapping from a column name to a property.
        /// </summary>
        /// <param name="t">The type that contains the property to map to.</param>
        /// <param name="columnIndex">Index of the column.</param>
        /// <param name="propertyName">Name of the property.</param>
        public void AddMapping(Type t, int columnIndex, string propertyName)
        {
            var typeMapper = TypeMapperFactory.Create(t);
            var prop = t.GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public);
            typeMapper.ColumnsByIndex[columnIndex] = new ColumnInfo(prop);
        }

        /// <summary>
        /// Ignores a property.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="propertyExpression">The property expression.</param>
        public void Ignore<T>(Expression<Func<T, object>> propertyExpression)
        {
            var typeMapper = TypeMapperFactory.Create(typeof(T));
            var prop = GetPropertyInfo(propertyExpression);
            var kvp = typeMapper.ColumnsByName.FirstOrDefault(c => c.Value.Property == prop);
            if (kvp.Key != null) typeMapper.ColumnsByName.Remove(kvp.Key);
        }

        /// <summary>
        /// Ignores a property.
        /// </summary>
        /// <param name="t">The type that contains the property to map to.</param>
        /// <param name="propertyName">Name of the property.</param>
        public void Ignore(Type t, string propertyName)
        {
            var typeMapper = TypeMapperFactory.Create(t);
            var prop = t.GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public);
            var kvp = typeMapper.ColumnsByName.FirstOrDefault(c => c.Value.Property == prop);
            if (kvp.Key != null) typeMapper.ColumnsByName.Remove(kvp.Key);
        }
    }
}
