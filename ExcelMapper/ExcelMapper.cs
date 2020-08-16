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
using System.Text.Json;

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
        /// <returns>The objects read from the Excel file.</returns>
        public IEnumerable<T> Fetch<T>(string file, string sheetName) where T : new()
        {
            return Fetch(file, typeof(T), sheetName).OfType<T>();
        }

        /// <summary>
        /// Fetches objects from the specified sheet name.
        /// </summary>
        /// <param name="type">The type of objects the Excel file is mapped to.</param>
        /// <param name="file">The path to the Excel file.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <returns>The objects read from the Excel file.</returns>
        public IEnumerable Fetch(string file, Type type, string sheetName)
        {
            Workbook = WorkbookFactory.Create(file);
            return Fetch(type, sheetName);
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
            return Fetch(file, typeof(T), sheetIndex).OfType<T>();
        }

        /// <summary>
        /// Fetches objects from the specified sheet index.
        /// </summary>
        /// <param name="type">The type of objects the Excel file is mapped to.</param>
        /// <param name="file">The path to the Excel file.</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <returns>The objects read from the Excel file.</returns>
        public IEnumerable Fetch(string file, Type type, int sheetIndex)
        {
            Workbook = WorkbookFactory.Create(file);
            return Fetch(type, sheetIndex);
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
            return Fetch(stream, typeof(T), sheetName).OfType<T>();
        }

        /// <summary>
        /// Fetches objects from the specified sheet name.
        /// </summary>
        /// <param name="type">The type of objects the Excel file is mapped to.</param>
        /// <param name="stream">The stream the Excel file is read from.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <returns>The objects read from the Excel file.</returns>
        public IEnumerable Fetch(Stream stream, Type type, string sheetName)
        {
            Workbook = WorkbookFactory.Create(stream);
            return Fetch(type, sheetName);
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
            return Fetch(stream, typeof(T), sheetIndex).OfType<T>();
        }

        /// <summary>
        /// Fetches objects from the specified sheet index.
        /// </summary>
        /// <param name="type">The type of objects the Excel file is mapped to.</param>
        /// <param name="stream">The stream the Excel file is read from.</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <returns>The objects read from the Excel file.</returns>
        public IEnumerable Fetch(Stream stream, Type type, int sheetIndex)
        {
            Workbook = WorkbookFactory.Create(stream);
            return Fetch(type, sheetIndex);
        }

        /// <summary>
        /// Fetches objects from the specified sheet name.
        /// </summary>
        /// <typeparam name="T">The type of objects the Excel file is mapped to.</typeparam>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <returns>The objects read from the Excel file.</returns>
        /// <exception cref="System.ArgumentOutOfRangeException">Thrown when a sheet is not found</exception>
        public IEnumerable<T> Fetch<T>(string sheetName) where T : new()
        {
            return Fetch(typeof(T), sheetName).OfType<T>();
        }

        /// <summary>
        /// Fetches objects from the specified sheet name.
        /// </summary>
        /// <param name="type">The type of objects the Excel file is mapped to.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <returns>The objects read from the Excel file.</returns>
        /// <exception cref="System.ArgumentOutOfRangeException">Thrown when a sheet is not found</exception>
        public IEnumerable Fetch(Type type, string sheetName)
        {
            PrimitiveCheck(type);

            var sheet = Workbook.GetSheet(sheetName);
            if (sheet == null)
            {
                throw new ArgumentOutOfRangeException(nameof(sheetName), sheetName, "Sheet not found");
            }
            return Fetch(sheet, type);
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

        /// <summary>
        /// Fetches objects from the specified sheet index.
        /// </summary>
        /// <param name="type">The type of objects the Excel file is mapped to</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <returns>The objects read from the Excel file.</returns>
        public IEnumerable Fetch(Type type, int sheetIndex = 0)
        {
            PrimitiveCheck(type);

            var sheet = Workbook.GetSheetAt(sheetIndex);
            return Fetch(sheet, type);
        }

        IEnumerable<T> Fetch<T>(ISheet sheet) where T : new()
        {
            return Fetch(sheet, typeof(T)).OfType<T>();
        }

        IEnumerable Fetch(ISheet sheet, Type type)
        {
            var typeMapper = TypeMapperFactory.Create(type);
            var firstRow = sheet.GetRow(HeaderRow ? HeaderRowNumber : MinRowNumber);
            var cells = Enumerable.Range(0, firstRow.LastCellNum).Select(i => firstRow.GetCell(i, MissingCellPolicy.CREATE_NULL_AS_BLANK));
            var columns = cells
                .Where(c => !HeaderRow || (c.CellType == CellType.String && !string.IsNullOrWhiteSpace(c.StringCellValue)))
                .Select(c => new 
                { 
                    c.ColumnIndex, 
                    ColumnInfo = GetColumnInfo(typeMapper, c)
                })
                .Where(c => c.ColumnInfo != null)
                .ToDictionary(c => c.ColumnIndex, c => c.ColumnInfo);
            var i = MinRowNumber;
            IRow row = null;

            if (TrackObjects) Objects[sheet.SheetName] = new Dictionary<int, object>();

            while (i <= MaxRowNumber && (row = sheet.GetRow(i)) != null)
            {
                // optionally skip header row and blank rows
                if ((!HeaderRow || i != HeaderRowNumber) && (!SkipBlankRows || row.Cells.Any(c => !IsCellBlank(c))))
                {
                    var o = Activator.CreateInstance(type);

                    foreach (var col in columns)
                    {
                        var cell = row.GetCell(col.Key);

                        if (cell != null && (!SkipBlankRows || !IsCellBlank(cell)))
                        {
                            var cellValue = GetCellValue(cell, col.Value);

                            try
                            {
                                col.Value.SetProperty(o, cellValue);
                            }
                            catch (Exception e)
                            {
                                throw new ExcelMapperConvertException(cellValue, col.Value.PropertyType, i, col.Key, e);
                            }
                        }
                    }

                    if (TrackObjects) Objects[sheet.SheetName][i] = o;

                    yield return o;
                }

                i++;
            }
        }

        /// <summary>
        /// Fetches objects from the specified sheet name using async I/O.
        /// </summary>
        /// <typeparam name="T">The type of objects the Excel file is mapped to.</typeparam>
        /// <param name="file">The path to the Excel file.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <returns>The objects read from the Excel file.</returns>
        public async Task<IEnumerable<T>> FetchAsync<T>(string file, string sheetName) where T : new()
        {
            return (await FetchAsync(file, typeof(T), sheetName)).OfType<T>();
        }

        /// <summary>
        /// Fetches objects from the specified sheet name using async I/O.
        /// </summary>
        /// <param name="type">The type of objects the Excel file is mapped to.</param>
        /// <param name="file">The path to the Excel file.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <returns>The objects read from the Excel file.</returns>
        public async Task<IEnumerable> FetchAsync(string file, Type type, string sheetName)
        {
            using var ms = await ReadAsync(file);
            return Fetch(ms, type, sheetName);
        }

        /// <summary>
        /// Fetches objects from the specified sheet index using async I/O.
        /// </summary>
        /// <typeparam name="T">The type of objects the Excel file is mapped to.</typeparam>
        /// <param name="file">The path to the Excel file.</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <returns>The objects read from the Excel file.</returns>
        public async Task<IEnumerable<T>> FetchAsync<T>(string file, int sheetIndex = 0) where T : new()
        {
            using var ms = await ReadAsync(file);
            return Fetch(ms, typeof(T), sheetIndex).OfType<T>();
        }

        /// <summary>
        /// Fetches objects from the specified sheet index using async I/O.
        /// </summary>
        /// <param name="type">The type of objects the Excel file is mapped to.</param>
        /// <param name="file">The path to the Excel file.</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <returns>The objects read from the Excel file.</returns>
        public async Task<IEnumerable> FetchAsync(string file, Type type, int sheetIndex = 0)
        {
            using var ms = await ReadAsync(file);
            return Fetch(ms, type, sheetIndex);
        }

        /// <summary>
        /// Fetches objects from the specified sheet name using async I/O.
        /// </summary>
        /// <typeparam name="T">The type of objects the Excel file is mapped to.</typeparam>
        /// <param name="stream">The stream the Excel file is read from.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <returns>The objects read from the Excel file.</returns>
        public async Task<IEnumerable<T>> FetchAsync<T>(Stream stream, string sheetName) where T : new()
        {
            using var ms = await ReadAsync(stream);
            return Fetch(ms, typeof(T), sheetName).OfType<T>();
        }

        /// <summary>
        /// Fetches objects from the specified sheet name using async I/O.
        /// </summary>
        /// <param name="type">The type of objects the Excel file is mapped to.</param>
        /// <param name="stream">The stream the Excel file is read from.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <returns>The objects read from the Excel file.</returns>
        public async Task<IEnumerable> FetchAsync(Stream stream, Type type, string sheetName)
        {
            using var ms = await ReadAsync(stream);
            return Fetch(ms, type, sheetName);
        }

        /// <summary>
        /// Fetches objects from the specified sheet index using async I/O.
        /// </summary>
        /// <typeparam name="T">The type of objects the Excel file is mapped to.</typeparam>
        /// <param name="stream">The stream the Excel file is read from.</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <returns>The objects read from the Excel file.</returns>
        public async Task<IEnumerable<T>> FetchAsync<T>(Stream stream, int sheetIndex = 0) where T : new()
        {
            using var ms = await ReadAsync(stream);
            return Fetch(ms, typeof(T), sheetIndex).OfType<T>();
        }

        /// <summary>
        /// Fetches objects from the specified sheet index using async I/O.
        /// </summary>
        /// <param name="type">The type of objects the Excel file is mapped to.</param>
        /// <param name="stream">The stream the Excel file is read from.</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <returns>The objects read from the Excel file.</returns>
        public async Task<IEnumerable> FetchAsync(Stream stream, Type type, int sheetIndex = 0)
        {
            using var ms = await ReadAsync(stream);
            return Fetch(ms, type, sheetIndex);
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

        ColumnInfo GetColumnInfo(TypeMapper typeMapper, ICell cell)
        {
            var colByIndex = typeMapper.GetColumnByIndex(cell.ColumnIndex);
            if (!HeaderRow || colByIndex != null)
                return colByIndex;
            var name = cell.StringCellValue;
            var colByName = typeMapper.GetColumnByName(name);
            // map column by name only if it hasn't been mapped to another property by index
            if (colByName != null
                && !typeMapper.ColumnsByIndex.Any(c => c.Value.Property.Name.Equals(name, StringComparison.OrdinalIgnoreCase)))
                return colByName;
            return null;
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
            using var fs = File.Open(file, FileMode.Create, FileAccess.Write);
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
            using var fs = File.Open(file, FileMode.Create, FileAccess.Write);
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
            using var fs = File.Open(file, FileMode.Create, FileAccess.Write);
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
            using var fs = File.Open(file, FileMode.Create, FileAccess.Write);
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
                    col.Value.SetCellStyle(cell);
                    col.Value.SetCell(cell, col.Value.GetProperty(o.Value));
                }
            }

            Workbook.Write(stream);
        }

        void Save<T>(Stream stream, ISheet sheet, IEnumerable<T> objects)
        {
            var typeMapper = TypeMapperFactory.Create(typeof(T));
            var columnsByIndex = GetColumns(sheet, typeMapper);
            var i = MinRowNumber;

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
                    col.Value.SetCellStyle(cell);
                    col.Value.SetCell(cell, col.Value.GetProperty(o));
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

        /// <summary>
        /// Saves the specified objects to the specified Excel file using async I/O.
        /// </summary>
        /// <typeparam name="T">The type of objects to save.</typeparam>
        /// <param name="file">The path to the Excel file.</param>
        /// <param name="objects">The objects to save.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="xlsx">if set to <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
        public async Task SaveAsync<T>(string file, IEnumerable<T> objects, string sheetName, bool xlsx = true)
        {
            var ms = new MemoryStream();
            Save(ms, objects, sheetName, xlsx);
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
        public async Task SaveAsync<T>(string file, IEnumerable<T> objects, int sheetIndex = 0, bool xlsx = true)
        {
            var ms = new MemoryStream();
            Save(ms, objects, sheetIndex, xlsx);
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
        public async Task SaveAsync<T>(Stream stream, IEnumerable<T> objects, string sheetName, bool xlsx = true)
        {
            var ms = new MemoryStream();
            Save(ms, objects, sheetName, xlsx);
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
        public async Task SaveAsync<T>(Stream stream, IEnumerable<T> objects, int sheetIndex = 0, bool xlsx = true)
        {
            var ms = new MemoryStream();
            Save(ms, objects, sheetIndex, xlsx);
            await SaveAsync(stream, ms);
        }

        /// <summary>
        /// Saves tracked objects to the specified Excel file using async I/O.
        /// </summary>
        /// <param name="file">The path to the Excel file.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="xlsx">if set to <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
        public async Task SaveAsync(string file, string sheetName, bool xlsx = true)
        {
            var ms = new MemoryStream();
            Save(ms, sheetName, xlsx);
            await SaveAsync(file, ms.ToArray());
        }

        /// <summary>
        /// Saves tracked objects to the specified Excel file using async I/O.
        /// </summary>
        /// <param name="file">The path to the Excel file.</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <param name="xlsx">if set to <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
        public async Task SaveAsync(string file, int sheetIndex = 0, bool xlsx = true)
        {
            var ms = new MemoryStream();
            Save(ms, sheetIndex, xlsx);
            await SaveAsync(file, ms.ToArray());
        }

        /// <summary>
        /// Saves tracked objects to the specified stream using async I/O.
        /// </summary>
        /// <param name="stream">The stream to save the objects to.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="xlsx">if set to <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
        public async Task SaveAsync(Stream stream, string sheetName, bool xlsx = true)
        {
            var ms = new MemoryStream();
            Save(ms, sheetName, xlsx);
            await SaveAsync(stream, ms);
        }

        /// <summary>
        /// Saves tracked objects to the specified stream using async I/O.
        /// </summary>
        /// <param name="stream">The stream to save the objects to.</param>
        /// <param name="sheetIndex">Index of the sheet.</param>
        /// <param name="xlsx">if set to <c>true</c> saves in .xlsx format; otherwise, saves in .xls format.</param>
        public async Task SaveAsync(Stream stream, int sheetIndex = 0, bool xlsx = true)
        {
            var ms = new MemoryStream();
            Save(ms, sheetIndex, xlsx);
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

        static void SetColumnStyles(ISheet sheet, Dictionary<int, ColumnInfo> columnsByIndex)
        {
            foreach (var c in columnsByIndex)
                c.Value.SetColumnStyle(sheet, c.Key);
        }

        Dictionary<int, ColumnInfo> GetColumns(ISheet sheet, TypeMapper typeMapper)
        {
            var columnsByIndex = typeMapper.ColumnsByIndex;

            SetColumnStyles(sheet, columnsByIndex);

            if (HeaderRow)
            {
                var columnsByName = typeMapper.ColumnsByName;
                var headerRow = sheet.GetRow(HeaderRowNumber);
                var hasColumnsByIndex = columnsByIndex.Any();

                if (headerRow == null)
                {
                    var j = 0;
                    headerRow = sheet.CreateRow(HeaderRowNumber);

                    if (!hasColumnsByIndex)
                        columnsByIndex = new Dictionary<int, ColumnInfo>();

                    foreach (var getter in columnsByName)
                    {
                        var columnIndex = !hasColumnsByIndex ? j : columnsByIndex.First(c => c.Value.Property == getter.Value.Property).Key;
                        var cell = headerRow.GetCell(columnIndex, MissingCellPolicy.CREATE_NULL_AS_BLANK);

                        if (!hasColumnsByIndex)
                            columnsByIndex[j] = getter.Value;

                        cell.SetCellValue(getter.Key);

                        j++;
                    }
                }
                else if (!hasColumnsByIndex)
                {
                    columnsByIndex = headerRow.Cells
                        .Where(c => c.CellType == CellType.String && !string.IsNullOrWhiteSpace(c.StringCellValue))
                        .Select(c => new { c.ColumnIndex, ColumnInfo = typeMapper.GetColumnByName(c.StringCellValue) })
                        .Where(c => c.ColumnInfo != null)
                        .ToDictionary(c => c.ColumnIndex, c => c.ColumnInfo);
                }
            }

            return columnsByIndex;
        }

        static bool dateBugWorkaround = false;

        static object GetCellValue(ICell cell, ColumnInfo targetColumn)
        {
            var cellType = cell.CellType == CellType.Formula && (targetColumn.PropertyType != typeof(string) || targetColumn.FormulaResult) ? cell.CachedFormulaResultType : cell.CellType;

            switch (cellType)
            {
                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(cell))
                    {
                        // see https://github.com/mganss/ExcelMapper/issues/51
                        try
                        {
                            if (!dateBugWorkaround)
                                return cell.DateCellValue;
                            else
                                return DateUtil.GetJavaDate(cell.NumericCellValue);
                        }
                        catch (NullReferenceException)
                        {
                            dateBugWorkaround = true;
                            return DateUtil.GetJavaDate(cell.NumericCellValue);
                        }
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
        /// Adds a mapping from a column name to a property.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="columnName">Name of the column.</param>
        /// <param name="propertyExpression">The property expression.</param>
        public ColumnInfo AddMapping<T>(string columnName, Expression<Func<T, object>> propertyExpression)
        {
            var typeMapper = TypeMapperFactory.Create(typeof(T));
            var prop = GetPropertyInfo(propertyExpression);
            var columnInfo = new ColumnInfo(prop);
            typeMapper.ColumnsByName[columnName] = columnInfo;

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
            var columnInfo = new ColumnInfo(prop);
            typeMapper.ColumnsByIndex[columnIndex] = columnInfo;

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
            var columnInfo = new ColumnInfo(prop);
            typeMapper.ColumnsByName[columnName] = columnInfo;

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
            var columnInfo = new ColumnInfo(prop);
            typeMapper.ColumnsByIndex[columnIndex] = columnInfo;

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