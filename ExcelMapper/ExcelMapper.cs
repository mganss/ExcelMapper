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
using Ganss.Excel.Exceptions;
using System.Threading.Tasks;
using System.Globalization;
using System.Text.Json;
using NPOI.Util;
using System.Dynamic;

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
        /// Gets or sets the row number of the header row. Default is 0.
        /// The header row may be outside of the range of <see cref="MinRowNumber"/> and <see cref="MaxRowNumber"/>.
        /// </summary>
        /// <value>
        /// The header row number.
        /// </value>
        public int HeaderRowNumber { get; set; } = 0;

        /// <summary>
        /// Gets or sets the minimum row number of the rows that may contain data. Default is 0.
        /// </summary>
        /// <value>
        /// The minimum row number.
        /// </value>
        public int MinRowNumber { get; set; } = 0;

        /// <summary>
        /// Gets or sets the inclusive maximum row number of the rows that may contain data. Default is <see cref="int.MaxValue"/>.
        /// </summary>
        /// <value>
        /// The maximum row number.
        /// </value>
        public int MaxRowNumber { get; set; } = int.MaxValue;

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

        /// <summary>
        /// Gets or sets the <see cref="DataFormatter"/> object to use when formatting cell values.
        /// </summary>
        /// <value>
        /// The <see cref="DataFormatter"/> object to use when formatting cell values.
        /// </value>
        public DataFormatter DataFormatter { get; set; } = new DataFormatter(CultureInfo.InvariantCulture);

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
        /// <param name="workbook">The workbook.</param>
        public ExcelMapper(IWorkbook workbook)
        {
            Workbook = workbook;
        }

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
        /// <param name="valueParser">Allow value parsing</param>
        /// <returns>The objects read from the Excel file.</returns>
        public IEnumerable<T> Fetch<T>(string file, string sheetName, Func<string, object, object> valueParser = null)
            where T : new()
        {
            return Fetch(file, typeof(T), sheetName, valueParser).OfType<T>();
        }

        /// <summary>
        /// Fetches objects from the specified sheet name.
        /// </summary>
        /// <param name="type">The type of objects the Excel file is mapped to.</param>
        /// <param name="file">The path to the Excel file.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="valueParser">Allow value parsing</param>
        /// <returns>The objects read from the Excel file.</returns>
        public IEnumerable Fetch(string file, Type type, string sheetName, Func<string, object, object> valueParser = null)
        {
            Workbook = WorkbookFactory.Create(file);
            return Fetch(type, sheetName, valueParser);
        }

        /// <summary>
        /// Fetches objects from the specified sheet name.
        /// </summary>
        /// <param name="file">The path to the Excel file.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="valueParser">Allow value parsing</param>
        /// <returns>The objects read from the Excel file.</returns>
        public IEnumerable<dynamic> Fetch(string file, string sheetName, Func<string, object, object> valueParser = null)
        {
            Workbook = WorkbookFactory.Create(file);
            return Fetch(sheetName, valueParser);
        }

        /// <summary>
        /// Fetches objects from the specified sheet index.
        /// </summary>
        /// <typeparam name="T">The type of objects the Excel file is mapped to.</typeparam>
        /// <param name="file">The path to the Excel file.</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <param name="valueParser">Allow value parsing</param>
        /// <returns>The objects read from the Excel file.</returns>
        public IEnumerable<T> Fetch<T>(string file, int sheetIndex, Func<string, object, object> valueParser = null)
            where T : new()
        {
            return Fetch(file, typeof(T), sheetIndex, valueParser).OfType<T>();
        }

        /// <summary>
        /// Fetches objects from the specified sheet index.
        /// </summary>
        /// <param name="type">The type of objects the Excel file is mapped to.</param>
        /// <param name="file">The path to the Excel file.</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <param name="valueParser">Allow value parsing</param>
        /// <returns>The objects read from the Excel file.</returns>
        public IEnumerable Fetch(string file, Type type, int sheetIndex, Func<string, object, object> valueParser = null)
        {
            Workbook = WorkbookFactory.Create(file);
            return Fetch(type, sheetIndex, valueParser);
        }

        /// <summary>
        /// Fetches objects from the specified sheet index.
        /// </summary>
        /// <param name="file">The path to the Excel file.</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <param name="valueParser">Allow value parsing</param>
        /// <returns>The objects read from the Excel file.</returns>
        public IEnumerable<dynamic> Fetch(string file, int sheetIndex, Func<string, object, object> valueParser = null)
        {
            Workbook = WorkbookFactory.Create(file);
            return Fetch(sheetIndex, valueParser);
        }

        /// <summary>
        /// Fetches objects from the specified sheet name.
        /// </summary>
        /// <typeparam name="T">The type of objects the Excel file is mapped to.</typeparam>
        /// <param name="stream">The stream the Excel file is read from.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="valueParser">Allow value parsing</param>
        /// <returns>The objects read from the Excel file.</returns>
        public IEnumerable<T> Fetch<T>(Stream stream, string sheetName, Func<string, object, object> valueParser = null)
            where T : new()
        {
            return Fetch(stream, typeof(T), sheetName, valueParser).OfType<T>();
        }

        /// <summary>
        /// Fetches objects from the specified sheet name.
        /// </summary>
        /// <param name="type">The type of objects the Excel file is mapped to.</param>
        /// <param name="stream">The stream the Excel file is read from.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="valueParser">Allow value parsing</param>
        /// <returns>The objects read from the Excel file.</returns>
        public IEnumerable Fetch(Stream stream, Type type, string sheetName, Func<string, object, object> valueParser = null)
        {
            Workbook = WorkbookFactory.Create(stream);
            return Fetch(type, sheetName, valueParser);
        }

        /// <summary>
        /// Fetches objects from the specified sheet name.
        /// </summary>
        /// <param name="stream">The stream the Excel file is read from.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="valueParser">Allow value parsing</param>
        /// <returns>The objects read from the Excel file.</returns>
        public IEnumerable<dynamic> Fetch(Stream stream, string sheetName, Func<string, object, object> valueParser = null)
        {
            Workbook = WorkbookFactory.Create(stream);
            return Fetch(sheetName, valueParser);
        }

        /// <summary>
        /// Fetches objects from the specified sheet index.
        /// </summary>
        /// <typeparam name="T">The type of objects the Excel file is mapped to.</typeparam>
        /// <param name="stream">The stream the Excel file is read from.</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <param name="valueParser">Allow value parsing</param>
        /// <returns>The objects read from the Excel file.</returns>
        public IEnumerable<T> Fetch<T>(Stream stream, int sheetIndex, Func<string, object, object> valueParser = null)
            where T : new()
        {
            return Fetch(stream, typeof(T), sheetIndex, valueParser).OfType<T>();
        }

        /// <summary>
        /// Fetches objects from the specified sheet index.
        /// </summary>
        /// <param name="type">The type of objects the Excel file is mapped to.</param>
        /// <param name="stream">The stream the Excel file is read from.</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <param name="valueParser">Allow value parsing</param>
        /// <returns>The objects read from the Excel file.</returns>
        public IEnumerable Fetch(Stream stream, Type type, int sheetIndex, Func<string, object, object> valueParser = null)
        {
            Workbook = WorkbookFactory.Create(stream);
            return Fetch(type, sheetIndex, valueParser);
        }

        /// <summary>
        /// Fetches objects from the specified sheet index.
        /// </summary>
        /// <param name="stream">The stream the Excel file is read from.</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <param name="valueParser">Allow value parsing</param>
        /// <returns>The objects read from the Excel file.</returns>
        public IEnumerable<dynamic> Fetch(Stream stream, int sheetIndex, Func<string, object, object> valueParser = null)
        {
            Workbook = WorkbookFactory.Create(stream);
            return Fetch(sheetIndex, valueParser);
        }

        /// <summary>
        /// Fetches objects from the specified sheet name.
        /// </summary>
        /// <typeparam name="T">The type of objects the Excel file is mapped to.</typeparam>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="valueParser">Allow value parsing</param>
        /// <returns>The objects read from the Excel file.</returns>
        /// <exception cref="System.ArgumentOutOfRangeException">Thrown when a sheet is not found</exception>
        public IEnumerable<T> Fetch<T>(string sheetName, Func<string, object, object> valueParser = null)
            where T : new()
        {
            return Fetch(typeof(T), sheetName, valueParser).OfType<T>();
        }

        /// <summary>
        /// Fetches objects from the specified sheet name.
        /// </summary>
        /// <param name="type">The type of objects the Excel file is mapped to.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="valueParser">Allow value parsing</param>
        /// <returns>The objects read from the Excel file.</returns>
        /// <exception cref="System.ArgumentOutOfRangeException">Thrown when a sheet is not found</exception>
        public IEnumerable Fetch(Type type, string sheetName, Func<string, object, object> valueParser = null)
        {
            PrimitiveCheck(type);

            var sheet = Workbook.GetSheet(sheetName);
            if (sheet == null)
            {
                throw new ArgumentOutOfRangeException(nameof(sheetName), sheetName, "Sheet not found");
            }
            return Fetch(sheet, type, valueParser);
        }

        /// <summary>
        /// Fetches objects from the specified sheet name.
        /// </summary>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="valueParser">Allow value parsing</param>
        /// <returns>The objects read from the Excel file.</returns>
        /// <exception cref="System.ArgumentOutOfRangeException">Thrown when a sheet is not found</exception>
        public IEnumerable<dynamic> Fetch(string sheetName, Func<string, object, object> valueParser = null)
        {
            var sheet = Workbook.GetSheet(sheetName);
            if (sheet == null)
            {
                throw new ArgumentOutOfRangeException(nameof(sheetName), sheetName, "Sheet not found");
            }
            return Fetch(sheet, valueParser);
        }

        /// <summary>
        /// Fetches objects from the specified sheet index.
        /// </summary>
        /// <typeparam name="T">The type of objects the Excel file is mapped to.</typeparam>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <param name="valueParser">Allow value parsing</param>
        /// <returns>The objects read from the Excel file.</returns>
        public IEnumerable<T> Fetch<T>(int sheetIndex = 0, Func<string, object, object> valueParser = null) where T : new()
        {
            var sheet = Workbook.GetSheetAt(sheetIndex);
            return Fetch<T>(sheet, valueParser);
        }

        /// <summary>
        /// Fetches objects from the specified sheet index.
        /// </summary>
        /// <param name="type">The type of objects the Excel file is mapped to</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <param name="valueParser">Allow value parsing</param>
        /// <returns>The objects read from the Excel file.</returns>
        public IEnumerable Fetch(Type type, int sheetIndex = 0, Func<string, object, object> valueParser = null)
        {
            PrimitiveCheck(type);

            var sheet = Workbook.GetSheetAt(sheetIndex);
            return Fetch(sheet, type, valueParser);
        }

        /// <summary>
        /// Fetches objects from the specified sheet index.
        /// </summary>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <param name="valueParser">Allow value parsing</param>
        /// <returns>The objects read from the Excel file.</returns>
        public IEnumerable<dynamic> Fetch(int sheetIndex = 0, Func<string, object, object> valueParser = null)
        {
            var sheet = Workbook.GetSheetAt(sheetIndex);
            return Fetch(sheet, valueParser);
        }

        IEnumerable<T> Fetch<T>(ISheet sheet, Func<string, object, object> valueParser = null) where T : new()
        {
            return Fetch(sheet, typeof(T), valueParser).OfType<T>();
        }

        IEnumerable Fetch(ISheet sheet, Type type, Func<string, object, object> valueParser = null)
        {
            var firstRow = sheet.GetRow(HeaderRow ? HeaderRowNumber : MinRowNumber);

            if (firstRow == null)
                yield break;

            var cells = Enumerable.Range(0, firstRow.LastCellNum).Select(i => firstRow.GetCell(i, MissingCellPolicy.CREATE_NULL_AS_BLANK));
            var firstRowCells = cells
                .Where(c => !HeaderRow || (c.CellType == CellType.String && !string.IsNullOrWhiteSpace(c.StringCellValue)));
            var typeMapper = type != null ? TypeMapperFactory.Create(type) : TypeMapper.Create(firstRowCells);
            var columns = firstRowCells
                .Select(c => new
                {
                    c.ColumnIndex,
                    ColumnInfo = GetColumnInfo(typeMapper, c)
                })
                .Where(c => c.ColumnInfo?.Any() ?? false)
                .ToDictionary(c => c.ColumnIndex, c => c.ColumnInfo);
            var i = MinRowNumber;
            IRow row = null;

            if (TrackObjects) Objects[sheet.SheetName] = new Dictionary<int, object>();

            var objInstanceIdx = 0;

            while (i <= MaxRowNumber && (row = sheet.GetRow(i)) != null)
            {
                // optionally skip header row and blank rows
                if ((!HeaderRow || i != HeaderRowNumber) && (!SkipBlankRows || row.Cells.Any(c => !IsCellBlank(c))))
                {
                    var o = type != null ? Activator.CreateInstance(type) : new ExpandoObject();

                    typeMapper?.BeforeMappingActionInvoker?.Invoke(o, objInstanceIdx);

                    foreach (var col in columns)
                    {
                        var cell = row.GetCell(col.Key);

                        if (cell != null && (!SkipBlankRows || !IsCellBlank(cell)))
                        {
                            foreach (var ci in col.Value.Where(c => c.Directions.HasFlag(MappingDirections.ExcelToObject)))
                            {
                                var cellValue = GetCellValue(cell, ci);
                                try
                                {
                                    if (valueParser != null)
                                        cellValue = valueParser(string.IsNullOrWhiteSpace(ci.Name) ? col.Key.ToString() : ci.Name, cellValue);

                                    ci.SetProperty(o, cellValue, cell);
                                }
                                catch (Exception e)
                                {
                                    throw new ExcelMapperConvertException(cellValue, ci.PropertyType, i, col.Key, e);
                                }
                            }
                        }
                    }

                    if (TrackObjects) Objects[sheet.SheetName][i] = o;

                    typeMapper?.AfterMappingActionInvoker?.Invoke(o, objInstanceIdx);

                    objInstanceIdx++;

                    yield return o;
                }

                i++;
            }
        }

        IEnumerable<dynamic> Fetch(ISheet sheet, Func<string, object, object> valueParser = null) =>
            Fetch(sheet, type: null, valueParser).Cast<dynamic>();

        /// <summary>
        /// Fetches objects from the specified sheet name using async I/O.
        /// </summary>
        /// <typeparam name="T">The type of objects the Excel file is mapped to.</typeparam>
        /// <param name="file">The path to the Excel file.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="valueParser">Allow value parsing</param>
        /// <returns>The objects read from the Excel file.</returns>
        public async Task<IEnumerable<T>> FetchAsync<T>(string file, string sheetName, Func<string, object, object> valueParser = null)
            where T : new()
        {
            return (await FetchAsync(file, typeof(T), sheetName, valueParser)).OfType<T>();
        }

        /// <summary>
        /// Fetches dynamic objects from the specified sheet name using async I/O.
        /// </summary>
        /// <param name="file">The path to the Excel file.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="valueParser">Allow value parsing</param>
        /// <returns>The objects read from the Excel file.</returns>
        public async Task<IEnumerable<dynamic>> FetchAsync(string file, string sheetName, Func<string, object, object> valueParser = null)
        {
            using var ms = await ReadAsync(file);
            return Fetch(ms, sheetName, valueParser);
        }

        /// <summary>
        /// Fetches objects from the specified sheet name using async I/O.
        /// </summary>
        /// <param name="type">The type of objects the Excel file is mapped to.</param>
        /// <param name="file">The path to the Excel file.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="valueParser">Allow value parsing</param>
        /// <returns>The objects read from the Excel file.</returns>
        public async Task<IEnumerable> FetchAsync(string file, Type type, string sheetName, Func<string, object, object> valueParser = null)
        {
            using var ms = await ReadAsync(file);
            return Fetch(ms, type, sheetName, valueParser);
        }

        /// <summary>
        /// Fetches objects from the specified sheet index using async I/O.
        /// </summary>
        /// <typeparam name="T">The type of objects the Excel file is mapped to.</typeparam>
        /// <param name="file">The path to the Excel file.</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <param name="valueParser">Allow value parsing</param>
        /// <returns>The objects read from the Excel file.</returns>
        public async Task<IEnumerable<T>> FetchAsync<T>(string file, int sheetIndex = 0, Func<string, object, object> valueParser = null)
            where T : new()
        {
            using var ms = await ReadAsync(file);
            return Fetch(ms, typeof(T), sheetIndex, valueParser).OfType<T>();
        }

        /// <summary>
        /// Fetches dynamic objects from the specified sheet index using async I/O.
        /// </summary>
        /// <param name="file">The path to the Excel file.</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <param name="valueParser">Allow value parsing</param>
        /// <returns>The objects read from the Excel file.</returns>
        public async Task<IEnumerable<dynamic>> FetchAsync(string file, int sheetIndex = 0, Func<string, object, object> valueParser = null)
        {
            using var ms = await ReadAsync(file);
            return Fetch(ms, sheetIndex, valueParser).Cast<dynamic>();
        }

        /// <summary>
        /// Fetches objects from the specified sheet index using async I/O.
        /// </summary>
        /// <param name="type">The type of objects the Excel file is mapped to.</param>
        /// <param name="file">The path to the Excel file.</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <param name="valueParser">Allow value parsing</param>
        /// <returns>The objects read from the Excel file.</returns>
        public async Task<IEnumerable> FetchAsync(string file, Type type, int sheetIndex = 0, Func<string, object, object> valueParser = null)
        {
            using var ms = await ReadAsync(file);
            return Fetch(ms, type, sheetIndex, valueParser);
        }

        /// <summary>
        /// Fetches objects from the specified sheet name using async I/O.
        /// </summary>
        /// <typeparam name="T">The type of objects the Excel file is mapped to.</typeparam>
        /// <param name="stream">The stream the Excel file is read from.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="valueParser">Allow value parsing</param>
        /// <returns>The objects read from the Excel file.</returns>
        public async Task<IEnumerable<T>> FetchAsync<T>(Stream stream, string sheetName, Func<string, object, object> valueParser = null)
            where T : new()
        {
            using var ms = await ReadAsync(stream);
            return Fetch(ms, typeof(T), sheetName, valueParser).OfType<T>();
        }

        /// <summary>
        /// Fetches dynamic objects from the specified sheet name using async I/O.
        /// </summary>
        /// <param name="stream">The stream the Excel file is read from.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="valueParser">Allow value parsing</param>
        /// <returns>The objects read from the Excel file.</returns>
        public async Task<IEnumerable<dynamic>> FetchAsync(Stream stream, string sheetName, Func<string, object, object> valueParser = null)
        {
            using var ms = await ReadAsync(stream);
            return Fetch(ms, sheetName, valueParser).Cast<dynamic>();
        }

        /// <summary>
        /// Fetches objects from the specified sheet name using async I/O.
        /// </summary>
        /// <param name="type">The type of objects the Excel file is mapped to.</param>
        /// <param name="stream">The stream the Excel file is read from.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="valueParser">Allow value parsing</param>
        /// <returns>The objects read from the Excel file.</returns>
        public async Task<IEnumerable> FetchAsync(Stream stream, Type type, string sheetName, Func<string, object, object> valueParser = null)
        {
            using var ms = await ReadAsync(stream);
            return Fetch(ms, type, sheetName, valueParser);
        }

        /// <summary>
        /// Fetches objects from the specified sheet index using async I/O.
        /// </summary>
        /// <typeparam name="T">The type of objects the Excel file is mapped to.</typeparam>
        /// <param name="stream">The stream the Excel file is read from.</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <param name="valueParser">Allow value parsing</param>
        /// <returns>The objects read from the Excel file.</returns>
        public async Task<IEnumerable<T>> FetchAsync<T>(Stream stream, int sheetIndex = 0, Func<string, object, object> valueParser = null)
            where T : new()
        {
            using var ms = await ReadAsync(stream);
            return Fetch(ms, typeof(T), sheetIndex, valueParser).OfType<T>();
        }

        /// <summary>
        /// Fetches dynamic objects from the specified sheet index using async I/O.
        /// </summary>
        /// <param name="stream">The stream the Excel file is read from.</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <param name="valueParser">Allow value parsing</param>
        /// <returns>The objects read from the Excel file.</returns>
        public async Task<IEnumerable<dynamic>> FetchAsync(Stream stream, int sheetIndex = 0, Func<string, object, object> valueParser = null)
        {
            using var ms = await ReadAsync(stream);
            return Fetch(ms, sheetIndex, valueParser).Cast<dynamic>();
        }

        /// <summary>
        /// Fetches objects from the specified sheet index using async I/O.
        /// </summary>
        /// <param name="type">The type of objects the Excel file is mapped to.</param>
        /// <param name="stream">The stream the Excel file is read from.</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <param name="valueParser">Allow value parsing</param>
        /// <returns>The objects read from the Excel file.</returns>
        public async Task<IEnumerable> FetchAsync(Stream stream, Type type, int sheetIndex = 0, Func<string, object, object> valueParser = null)
        {
            using var ms = await ReadAsync(stream);
            return Fetch(ms, type, sheetIndex, valueParser);
        }

        static async Task<Stream> ReadAsync(string file)
        {
            using var fs = new FileStream(file, FileMode.Open, FileAccess.Read);
            var ms = new MemoryStream();
            await fs.CopyToAsync(ms);
            return ms;
        }

        static async Task<Stream> ReadAsync(Stream stream)
        {
            var ms = new MemoryStream();
            await stream.CopyToAsync(ms);
            return ms;
        }

        private static bool IsCellBlank(ICell cell)
        {
            return cell.CellType switch
            {
                CellType.String => string.IsNullOrWhiteSpace(cell.StringCellValue),
                CellType.Blank => true,
                _ => false,
            };
        }

        List<ColumnInfo> GetColumnInfo(TypeMapper typeMapper, ICell cell)
        {
            var colByIndex = typeMapper.GetColumnByIndex(cell.ColumnIndex);

            if (!HeaderRow || colByIndex != null)
                return colByIndex;

            var name = cell.StringCellValue;
            var colByName = typeMapper.GetColumnByName(name);

            // map column by name only if it hasn't been mapped to another property by index
            if (colByName != null
                && !typeMapper.ColumnsByIndex.SelectMany(ci => ci.Value).Any(c => c.Property.Name.Equals(name, StringComparison.OrdinalIgnoreCase)))
                return colByName;

            return new List<ColumnInfo>();
        }

        /// <summary>
        /// Saves the specified objects to the specified Excel file.
        /// </summary>
        /// <typeparam name="T">The type of objects to save.</typeparam>
        /// <param name="file">The path to the Excel file.</param>
        /// <param name="objects">The objects to save.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="xlsx">if set to <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
        /// <param name="valueConverter">converter receiving property name and value</param>
        public void Save<T>(string file, IEnumerable<T> objects, string sheetName, bool xlsx = true, Func<string, object, object> valueConverter = null)
        {
            using var fs = File.Open(file, FileMode.Create, FileAccess.Write);
            Save(fs, objects, sheetName, xlsx, valueConverter);
        }

        /// <summary>
        /// Saves the specified objects to the specified Excel file.
        /// </summary>
        /// <typeparam name="T">The type of objects to save.</typeparam>
        /// <param name="file">The path to the Excel file.</param>
        /// <param name="objects">The objects to save.</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <param name="xlsx">if set to <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
        /// <param name="valueConverter">converter receiving property name and value</param>
        public void Save<T>(string file, IEnumerable<T> objects, int sheetIndex = 0, bool xlsx = true, Func<string, object, object> valueConverter = null)
        {
            using var fs = File.Open(file, FileMode.Create, FileAccess.Write);
            Save(fs, objects, sheetIndex, xlsx, valueConverter);
        }

        /// <summary>
        /// Saves the specified objects to the specified stream.
        /// </summary>
        /// <typeparam name="T">The type of objects to save.</typeparam>
        /// <param name="stream">The stream to save the objects to.</param>
        /// <param name="objects">The objects to save.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="xlsx">if set to <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
        /// <param name="valueConverter">converter receiving property name and value</param>
        public void Save<T>(Stream stream, IEnumerable<T> objects, string sheetName, bool xlsx = true, Func<string, object, object> valueConverter = null)
        {
            if (Workbook == null)
                Workbook = xlsx ? (IWorkbook)new XSSFWorkbook() : (IWorkbook)new HSSFWorkbook();
            var sheet = Workbook.GetSheet(sheetName);
            if (sheet == null) sheet = Workbook.CreateSheet(sheetName);
            Save(stream, sheet, objects, valueConverter);
        }

        /// <summary>
        /// Saves the specified objects to the specified stream.
        /// </summary>
        /// <typeparam name="T">The type of objects to save.</typeparam>
        /// <param name="stream">The stream to save the objects to.</param>
        /// <param name="objects">The objects to save.</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <param name="xlsx">if set to <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
        /// <param name="valueConverter">converter receiving property name and value</param>
        public void Save<T>(Stream stream, IEnumerable<T> objects, int sheetIndex = 0, bool xlsx = true, Func<string, object, object> valueConverter = null)
        {
            if (Workbook == null)
                Workbook = xlsx ? (IWorkbook)new XSSFWorkbook() : (IWorkbook)new HSSFWorkbook();
            ISheet sheet;
            if (Workbook.NumberOfSheets > sheetIndex)
                sheet = Workbook.GetSheetAt(sheetIndex);
            else
                sheet = Workbook.CreateSheet();
            Save(stream, sheet, objects, valueConverter);
        }

        /// <summary>
        /// Saves tracked objects to the specified Excel file.
        /// </summary>
        /// <param name="file">The path to the Excel file.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="xlsx">if set to <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
        /// <param name="valueConverter">converter receiving property name and value</param>
        public void Save(string file, string sheetName, bool xlsx = true, Func<string, object, object> valueConverter = null)
        {
            using var fs = File.Open(file, FileMode.Create, FileAccess.Write);
            Save(fs, sheetName, xlsx, valueConverter);
        }

        /// <summary>
        /// Saves tracked objects to the specified Excel file.
        /// </summary>
        /// <param name="file">The path to the Excel file.</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <param name="xlsx">if set to <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
        /// <param name="valueConverter">converter receiving property name and value</param>
        public void Save(string file, int sheetIndex = 0, bool xlsx = true, Func<string, object, object> valueConverter = null)
        {
            using var fs = File.Open(file, FileMode.Create, FileAccess.Write);
            Save(fs, sheetIndex, xlsx, valueConverter);
        }

        /// <summary>
        /// Saves tracked objects to the specified stream.
        /// </summary>
        /// <param name="stream">The stream to save the objects to.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="xlsx">if set to <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
        /// <param name="valueConverter">converter receiving property name and value</param>
        public void Save(Stream stream, string sheetName, bool xlsx = true, Func<string, object, object> valueConverter = null)
        {
            if (Workbook == null)
                Workbook = xlsx ? (IWorkbook)new XSSFWorkbook() : (IWorkbook)new HSSFWorkbook();
            var sheet = Workbook.GetSheet(sheetName);
            if (sheet == null) sheet = Workbook.CreateSheet(sheetName);
            Save(stream, sheet, valueConverter);
        }

        /// <summary>
        /// Saves tracked objects to the specified stream.
        /// </summary>
        /// <param name="stream">The stream to save the objects to.</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <param name="xlsx">if set to <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
        /// <param name="valueConverter">converter receiving property name and value</param>
        public void Save(Stream stream, int sheetIndex = 0, bool xlsx = true, Func<string, object, object> valueConverter = null)
        {
            if (Workbook == null)
                Workbook = xlsx ? (IWorkbook)new XSSFWorkbook() : (IWorkbook)new HSSFWorkbook();
            var sheet = Workbook.GetSheetAt(sheetIndex);
            if (sheet == null) sheet = Workbook.CreateSheet();
            Save(stream, sheet, valueConverter);
        }

        void Save(Stream stream, ISheet sheet, Func<string, object, object> valueConverter = null)
        {
            var objects = Objects[sheet.SheetName];
            var typeMapper = TypeMapperFactory.Create(objects.First().Value);
            var columnsByIndex = typeMapper.ColumnsByIndex;
            var columnsByName = typeMapper.ColumnsByName;

            PrepareColumnsForSaving(ref columnsByIndex, ref columnsByName);

            GetColumns(sheet, typeMapper, ref columnsByIndex, ref columnsByName);

            SetColumnStyles(sheet, columnsByIndex);

            foreach (var o in objects)
            {
                var i = o.Key;
                var row = sheet.GetRow(i);
                if (row == null) row = sheet.CreateRow(i);

                foreach (var col in columnsByIndex)
                {
                    var cell = row.GetCell(col.Key, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    foreach (var ci in col.Value.Where(c => c.Directions.HasFlag(MappingDirections.ObjectToExcel)))
                    {
                        SetCell(valueConverter, o.Value, cell, ci);
                    }
                }
            }

            Workbook.Write(stream);
        }

        void Save<T>(Stream stream, ISheet sheet, IEnumerable<T> objects, Func<string, object, object> valueConverter = null)
        {
            var firstObject = objects.FirstOrDefault();
            var typeMapper = firstObject is ExpandoObject ? TypeMapperFactory.Create(firstObject) : TypeMapperFactory.Create(typeof(T));
            var columnsByIndex = typeMapper.ColumnsByIndex;
            var columnsByName = typeMapper.ColumnsByName;
            var i = MinRowNumber;

            PrepareColumnsForSaving(ref columnsByIndex, ref columnsByName);

            GetColumns(sheet, typeMapper, ref columnsByIndex, ref columnsByName);

            SetColumnStyles(sheet, columnsByIndex);

            foreach (var o in objects)
            {
                if (i > MaxRowNumber) break;

                if (HeaderRow && i == HeaderRowNumber)
                    i++;

                var row = sheet.GetRow(i);
                if (row == null) row = sheet.CreateRow(i);

                foreach (var col in columnsByIndex)
                {
                    var cell = row.GetCell(col.Key, MissingCellPolicy.CREATE_NULL_AS_BLANK);

                    foreach (var ci in col.Value.Where(c => c.Directions.HasFlag(MappingDirections.ObjectToExcel)))
                    {
                        SetCell(valueConverter, o, cell, ci);
                    }
                }

                i++;
            }

            if (SkipBlankRows)
            {
                while (i <= sheet.LastRowNum && i <= MaxRowNumber)
                {
                    var row = sheet.GetRow(i);
                    while (row.Cells.Any())
                        row.RemoveCell(row.GetCell(row.FirstCellNum));
                    i++;
                }
            }

            Workbook.Write(stream);
        }

        private static void SetCell<T>(Func<string, object, object> valueConverter, T objInstance, ICell cell, ColumnInfo ci)
        {
            Type oldType = null;
            object val = ci.GetProperty(objInstance);
            if (valueConverter != null)
            {
                val = valueConverter(ci.Name, val);
                var newType = val?.GetType() ?? oldType;
                if (newType != ci.PropertyType)
                {
                    oldType = ci.PropertyType;
                    ci.ChangeSetterType(newType);
                }
            }
            ci.SetCellStyle(cell);
            ci.SetCell(cell, val);
            if (oldType != null)
                ci.ChangeSetterType(oldType);
        }

        private static void PrepareColumnsForSaving(ref Dictionary<int, List<ColumnInfo>> columnsByIndex, ref Dictionary<string, List<ColumnInfo>> columnsByName)
        {
            // All columns with <see cref="MappingDirections.ExcelToObject"/> direction only should not be saved
            columnsByName = columnsByName.Where(kvp => !kvp.Value.All(ci => ci.Directions == MappingDirections.ExcelToObject))
                .ToDictionary(kvp => kvp.Key, kvp => kvp.Value);

            columnsByIndex = columnsByIndex.Where(kvp => !kvp.Value.All(ci => ci.Directions == MappingDirections.ExcelToObject))
                .ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
        }

        /// <summary>
        /// Saves the specified objects to the specified Excel file using async I/O.
        /// </summary>
        /// <typeparam name="T">The type of objects to save.</typeparam>
        /// <param name="file">The path to the Excel file.</param>
        /// <param name="objects">The objects to save.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="xlsx">if set to <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
        /// <param name="valueConverter">converter receiving property name and value</param>
        public async Task SaveAsync<T>(string file, IEnumerable<T> objects, string sheetName, bool xlsx = true, Func<string, object, object> valueConverter = null)
        {
            var ms = new MemoryStream();
            Save(ms, objects, sheetName, xlsx, valueConverter);
            await SaveAsync(file, ms.ToArray());
        }

        /// <summary>
        /// Saves the specified objects to the specified Excel file using async I/O.
        /// </summary>
        /// <typeparam name="T">The type of objects to save.</typeparam>
        /// <param name="file">The path to the Excel file.</param>
        /// <param name="objects">The objects to save.</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <param name="xlsx">if set to <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
        /// <param name="valueConverter">converter receiving property name and value</param>
        public async Task SaveAsync<T>(string file, IEnumerable<T> objects, int sheetIndex = 0, bool xlsx = true, Func<string, object, object> valueConverter = null)
        {
            var ms = new MemoryStream();
            Save(ms, objects, sheetIndex, xlsx, valueConverter);
            await SaveAsync(file, ms.ToArray());
        }

        /// <summary>
        /// Saves the specified objects to the specified stream using async I/O.
        /// </summary>
        /// <typeparam name="T">The type of objects to save.</typeparam>
        /// <param name="stream">The stream to save the objects to.</param>
        /// <param name="objects">The objects to save.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="xlsx">if set to <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
        /// <param name="valueConverter">converter receiving property name and value</param>
        public async Task SaveAsync<T>(Stream stream, IEnumerable<T> objects, string sheetName, bool xlsx = true, Func<string, object, object> valueConverter = null)
        {
            var ms = new MemoryStream();
            Save(ms, objects, sheetName, xlsx, valueConverter);
            await SaveAsync(stream, ms);
        }

        /// <summary>
        /// Saves the specified objects to the specified stream using async I/O.
        /// </summary>
        /// <typeparam name="T">The type of objects to save.</typeparam>
        /// <param name="stream">The stream to save the objects to.</param>
        /// <param name="objects">The objects to save.</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <param name="xlsx">if set to <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
        /// <param name="valueConverter">converter receiving property name and value</param>
        public async Task SaveAsync<T>(Stream stream, IEnumerable<T> objects, int sheetIndex = 0, bool xlsx = true, Func<string, object, object> valueConverter = null)
        {
            var ms = new MemoryStream();
            Save(ms, objects, sheetIndex, xlsx, valueConverter);
            await SaveAsync(stream, ms);
        }

        /// <summary>
        /// Saves tracked objects to the specified Excel file using async I/O.
        /// </summary>
        /// <param name="file">The path to the Excel file.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="xlsx">if set to <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
        /// <param name="valueConverter">converter receiving property name and value</param>
        public async Task SaveAsync(string file, string sheetName, bool xlsx = true, Func<string, object, object> valueConverter = null)
        {
            var ms = new MemoryStream();
            Save(ms, sheetName, xlsx, valueConverter);
            await SaveAsync(file, ms.ToArray());
        }

        /// <summary>
        /// Saves tracked objects to the specified Excel file using async I/O.
        /// </summary>
        /// <param name="file">The path to the Excel file.</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <param name="xlsx">if set to <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
        /// <param name="valueConverter">converter receiving property name and value</param>
        public async Task SaveAsync(string file, int sheetIndex = 0, bool xlsx = true, Func<string, object, object> valueConverter = null)
        {
            var ms = new MemoryStream();
            Save(ms, sheetIndex, xlsx, valueConverter);
            await SaveAsync(file, ms.ToArray());
        }

        /// <summary>
        /// Saves tracked objects to the specified stream using async I/O.
        /// </summary>
        /// <param name="stream">The stream to save the objects to.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="xlsx">if set to <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
        /// <param name="valueConverter">converter receiving property name and value</param>
        public async Task SaveAsync(Stream stream, string sheetName, bool xlsx = true, Func<string, object, object> valueConverter = null)
        {
            var ms = new MemoryStream();
            Save(ms, sheetName, xlsx, valueConverter);
            await SaveAsync(stream, ms);
        }

        /// <summary>
        /// Saves tracked objects to the specified stream using async I/O.
        /// </summary>
        /// <param name="stream">The stream to save the objects to.</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <param name="xlsx">if set to <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
        /// <param name="valueConverter">converter receiving property name and value</param>
        public async Task SaveAsync(Stream stream, int sheetIndex = 0, bool xlsx = true, Func<string, object, object> valueConverter = null)
        {
            var ms = new MemoryStream();
            Save(ms, sheetIndex, xlsx, valueConverter);
            await SaveAsync(stream, ms);
        }

        static async Task SaveAsync(string file, byte[] buf)
        {
            using var fs = new FileStream(file, FileMode.OpenOrCreate, FileAccess.Write);
            await fs.WriteAsync(buf, 0, buf.Length);
        }

        static async Task SaveAsync(Stream stream, MemoryStream ms)
        {
            var buf = ms.ToArray();
            await stream.WriteAsync(buf, 0, buf.Length);
        }

        static void SetColumnStyles(ISheet sheet, Dictionary<int, List<ColumnInfo>> columnsByIndex)
        {
            foreach (var col in columnsByIndex)
                col.Value.Where(c => c.Directions.HasFlag(MappingDirections.ObjectToExcel))
                    .ToList().ForEach(ci => ci.SetColumnStyle(sheet, col.Key));
        }

        Dictionary<int, string> GetColumns(ISheet sheet, TypeMapper typeMapper
            , ref Dictionary<int, List<ColumnInfo>> columnsByIndex
            , ref Dictionary<string, List<ColumnInfo>> columnsByName
        )
        {
            Dictionary<int, string> ColIndexNameMap = new Dictionary<int, string>();
            if (HeaderRow)
            {
                var headerRow = sheet.GetRow(HeaderRowNumber);
                var hasColumnsByIndex = columnsByIndex.Any();

                if (headerRow == null)
                {
                    var j = 0;
                    headerRow = sheet.CreateRow(HeaderRowNumber);

                    foreach (var getter in columnsByName)
                    {
                        var columnIndex = j;

                        if (hasColumnsByIndex)
                        {
                            columnIndex = (
                                from kvpi in columnsByIndex
                                from kvpci in kvpi.Value
                                join gci in getter.Value on kvpci.Name equals gci.Name
                                select kvpi
                            ).First().Key;
                        }

                        var cell = headerRow.GetCell(columnIndex, MissingCellPolicy.CREATE_NULL_AS_BLANK);

                        if (!hasColumnsByIndex)
                            columnsByIndex[j] = getter.Value;

                        ColIndexNameMap.Add(columnIndex, getter.Key);

                        cell.SetCellValue(getter.Key);

                        j++;
                    }
                }
                else if (!hasColumnsByIndex)
                {
                    columnsByIndex = headerRow.Cells
                        .Where(c => c.CellType == CellType.String && !string.IsNullOrWhiteSpace(c.StringCellValue))
                        .Select(c =>
                        {
                            var val = new { c.ColumnIndex, ColumnInfo = typeMapper.GetColumnByName(c.StringCellValue), ColumnName = c.StringCellValue };
                            ColIndexNameMap.Add(val.ColumnIndex, val.ColumnName);
                            return val;
                        })
                        .Where(c => c.ColumnInfo != null)
                        .ToDictionary(c => c.ColumnIndex, c => c.ColumnInfo);
                }
            }
            return ColIndexNameMap;
        }

        object GetCellValue(ICell cell, ColumnInfo targetColumn)
        {
            var formulaResult = cell.CellType == CellType.Formula && (targetColumn.PropertyType != typeof(string) || targetColumn.FormulaResult);
            var cellType = formulaResult ? cell.CachedFormulaResultType : cell.CellType;

            switch (cellType)
            {
                case CellType.Numeric:
                    if (!formulaResult && targetColumn.PropertyType == typeof(string))
                    {
                        return DataFormatter.FormatCellValue(cell);
                    }
                    else if (DateUtil.IsCellDateFormatted(cell))
                    {
                        // temporary workaround for https://github.com/tonyqus/npoi/issues/412
                        LocaleUtil.SetUserTimeZone(TimeZone.CurrentTimeZone);
                        return cell.DateCellValue;
                    }
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
                    if (targetColumn.Json)
                        return JsonSerializer.Deserialize(cell.StringCellValue, targetColumn.PropertyType);
                    else
                        return cell.StringCellValue;
            }
        }

        static PropertyInfo GetPropertyInfo<T>(Expression<Func<T, object>> propertyExpression)
        {
            var exp = (LambdaExpression)propertyExpression;
            var mExp = (exp.Body.NodeType == ExpressionType.MemberAccess) ?
                (MemberExpression)exp.Body :
                (MemberExpression)((UnaryExpression)exp.Body).Operand;
            return (PropertyInfo)mExp.Member;
        }

        static void PrimitiveCheck(Type type)
        {
            if (type.IsPrimitive || typeof(string).Equals(type) || typeof(object).Equals(type) || Nullable.GetUnderlyingType(type) != null)
            {
                throw new ArgumentException($"{type.Name} can not be used to map an excel because it is a primitive type");
            }
        }

        /// <summary>
        /// Action to call after an object is mapped
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="action"></param>
        /// <returns></returns>
        public ExcelMapper AddAfterMapping<T>(Action<T, int> action)
        {
            var typeMapper = TypeMapperFactory.Create(typeof(T));
            typeMapper.AfterMappingActionInvoker = ActionInvoker.CreateInstance(action);
            return this;
        }

        /// <summary>
        /// Action to call before an object is mapped
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="action"></param>
        /// <returns></returns>
        public ExcelMapper AddBeforeMapping<T>(Action<T, int> action)
        {
            var typeMapper = TypeMapperFactory.Create(typeof(T));
            typeMapper.BeforeMappingActionInvoker = ActionInvoker.CreateInstance(action);
            return this;
        }

        /// <summary>
        /// Adds a mapping from a column name to a property.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="columnName">Name of the column.</param>
        /// <param name="propertyExpression">The property expression.</param>
        public ColumnInfo AddMapping<T>(string columnName, Expression<Func<T, object>> propertyExpression)
        {
            var typeMapper = TypeMapperFactory.Create(typeof(T));
            var prop = GetPropertyInfo(propertyExpression);

            if (!typeMapper.ColumnsByName.ContainsKey(columnName))
                typeMapper.ColumnsByName.Add(columnName, new List<ColumnInfo>());

            var columnInfo = typeMapper.ColumnsByName[columnName].FirstOrDefault(ci => ci.Property.Name == prop.Name);
            if (columnInfo is null)
            {
                columnInfo = new ColumnInfo(prop);
                typeMapper.ColumnsByName[columnName].Add(columnInfo);
            }

            return columnInfo;
        }

        /// <summary>
        /// Adds a mapping from a column index to a property.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="columnIndex">Index of the column.</param>
        /// <param name="propertyExpression">The property expression.</param>
        public ColumnInfo AddMapping<T>(int columnIndex, Expression<Func<T, object>> propertyExpression)
        {
            var typeMapper = TypeMapperFactory.Create(typeof(T));
            var prop = GetPropertyInfo(propertyExpression);

            if (!typeMapper.ColumnsByIndex.ContainsKey(columnIndex))
                typeMapper.ColumnsByIndex.Add(columnIndex, new List<ColumnInfo>());

            var columnInfo = typeMapper.ColumnsByIndex[columnIndex].FirstOrDefault(ci => ci.Property.Name == prop.Name);
            if (columnInfo is null)
            {
                columnInfo = new ColumnInfo(prop);
                typeMapper.ColumnsByIndex[columnIndex].Add(columnInfo);
            }

            return columnInfo;
        }

        /// <summary>
        /// Adds a mapping from a column name to a property.
        /// </summary>
        /// <param name="t">The type that contains the property to map to.</param>
        /// <param name="columnName">Name of the column.</param>
        /// <param name="propertyName">Name of the property.</param>
        public ColumnInfo AddMapping(Type t, string columnName, string propertyName)
        {
            var typeMapper = TypeMapperFactory.Create(t);
            var prop = t.GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public);

            if (!typeMapper.ColumnsByName.ContainsKey(columnName))
                typeMapper.ColumnsByName.Add(columnName, new List<ColumnInfo>());

            var columnInfo = typeMapper.ColumnsByName[columnName].FirstOrDefault(ci => ci.Property.Name == prop.Name);
            if (columnInfo is null)
            {
                columnInfo = new ColumnInfo(prop);
                typeMapper.ColumnsByName[columnName].Add(columnInfo);
            }

            return columnInfo;
        }

        /// <summary>
        /// Adds a mapping from a column name to a property.
        /// </summary>
        /// <param name="t">The type that contains the property to map to.</param>
        /// <param name="columnIndex">Index of the column.</param>
        /// <param name="propertyName">Name of the property.</param>
        public ColumnInfo AddMapping(Type t, int columnIndex, string propertyName)
        {
            var typeMapper = TypeMapperFactory.Create(t);
            var prop = t.GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public);

            if (!typeMapper.ColumnsByIndex.ContainsKey(columnIndex))
                typeMapper.ColumnsByIndex.Add(columnIndex, new List<ColumnInfo>());

            var columnInfo = typeMapper.ColumnsByIndex[columnIndex].FirstOrDefault(ci => ci.Property.Name == prop.Name);
            if (columnInfo is null)
            {
                columnInfo = new ColumnInfo(prop);
                typeMapper.ColumnsByIndex[columnIndex].Add(columnInfo);
            }

            return columnInfo;
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

            typeMapper.ColumnsByName.Where(c => c.Value.Any(cc => cc.Property == prop))
                .ToList().ForEach(kvp => typeMapper.ColumnsByName.Remove(kvp.Key));
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

            typeMapper.ColumnsByName.Where(c => c.Value.Any(cc => cc.Property == prop))
                .ToList().ForEach(kvp => typeMapper.ColumnsByName.Remove(kvp.Key));
        }
    }
}