# ExcelMapper

[![NuGet version](https://badge.fury.io/nu/ExcelMapper.svg)](http://badge.fury.io/nu/ExcelMapper)
[![Build status](https://ci.appveyor.com/api/projects/status/tyyg8905i24qv9pg/branch/master?svg=true)](https://ci.appveyor.com/project/mganss/excelmapper/branch/master)
[![codecov.io](https://codecov.io/github/mganss/ExcelMapper/coverage.svg?branch=master)](https://codecov.io/github/mganss/ExcelMapper?branch=master)
[![netstandard2.0](https://img.shields.io/badge/netstandard-2.0-brightgreen.svg)](https://img.shields.io/badge/netstandard-2.0-brightgreen.svg)
[![net461](https://img.shields.io/badge/net-461-brightgreen.svg)](https://img.shields.io/badge/net-461-brightgreen.svg)

A library to map [POCO](https://en.wikipedia.org/wiki/Plain_Old_CLR_Object) objects to Excel files.

## Features

* Read and write Excel files
* Uses the pure managed [NPOI](https://github.com/tonyqus/npoi) library instead of the [Jet](https://en.wikipedia.org/wiki/Microsoft_Jet_Database_Engine) database engine ([NPOI users group](https://t.me/npoidevs))
* Map to Excel files using header rows (column names) or column indexes (no header row)
* Map nested objects (parent/child objects)
* Optionally skip blank lines when reading
* Preserve formatting when saving back files
* Optionally let the mapper track objects
* Map columns to properties through convention, attributes or method calls
* Use custom or builtin data formats for numeric and DateTime columns
* Map formulas or formula results depending on property type
* Map JSON
* Fetch/Save dynamic objects
* Use records

## Read objects from an Excel file

```C#
var products = new ExcelMapper("products.xlsx").Fetch<Product>();
```

This expects the Excel file to contain a header row with the column names. Objects are read from the first worksheet. If the column names equal the property names (ignoring case) no other configuration is necessary. The format of the Excel file (xlsx or xls) is autodetected.

## Map to specific column names

```C#
public class Product
{
  public string Name { get; set; }
  [Column("Number")]
  public int NumberInStock { get; set; }
  public decimal Price { get; set; }
}
```

This maps the column named `Number` to the `NumberInStock` property.

## Map to column indexes

```C#
public class Product
{
    [Column(1)]
    public string Name { get; set; }
    [Column(Letter="C")]
    public int NumberInStock { get; set; }
    [Column(4)]
    public decimal Price { get; set; }
}

var products = new ExcelMapper("products.xlsx") { HeaderRow = false }.Fetch<Product>();
```

Note that column indexes don't need to be consecutive. When mapping to column indexes, every property needs to be explicitly mapped through the `ColumnAttribute` attribute or the `AddMapping()` method. You can combine column indexes with column names to specify an explicit column order while still using a header row.

## Map through method calls

```C#
var excel = new ExcelMapper("products.xls");
excel.AddMapping<Product>("Number", p => p.NumberInStock);
excel.AddMapping<Product>(1, p => p.NumberInStock);
excel.AddMapping(typeof(Product), "Number", "NumberInStock");
excel.AddMapping(typeof(Product), ExcelMapper.LetterToIndex("A"), "NumberInStock");
```

## Multiple mappings

You can map a single column to multiple properties but you need to be aware of what should happen when mapping back from objects to Excel. To specify the single property you want to map back to Excel, add `MappingDirections.ExcelToObject` in the `Column` attribute of all other properties that map to the same column. Alternatively, you can use the `FromExcelOnly()` method when mapping through method calls.

```c#
public class Product
{
    public decimal Price { get; set; }
    [Column("Price", MappingDirections.ExcelToObject)]
    public string PriceString { get; set; }
}

// or

excel.AddMapping<Product>("Price", p => p.PriceString).FromExcelOnly();
```

## Dynamic mapping

You don't have to specify a mapping to static types, you can also fetch a collection of dynamic objects.

```c#
var products = new ExcelMapper("products.xlsx").Fetch(); // -> IEnumerable<dynamic>
products.First().Price += 1.0;
```

The returned dynamic objects are instances of `ExpandoObject` with an extra property called `__indexes__` that is a dictionary specifying the mapping from property names to
column indexes. If you set the `HeaderRow` property to `false` on the `ExcelMapper` object, the property names of the returned dynamic objects will match the Excel "letter" column names, i.e. "A" for column 1 etc.

## Save objects

```C#
var products = new List<Product>
{
    new Product { Name = "Nudossi", NumberInStock = 60, Price = 1.99m },
    new Product { Name = "Halloren", NumberInStock = 33, Price = 2.99m },
    new Product { Name = "Filinchen", NumberInStock = 100, Price = 0.99m },
};

new ExcelMapper().Save("products.xlsx", products, "Products");
```

This saves to the worksheet named "Products". If you save objects after having previously read from an Excel file using the same instance of `ExcelMapper` the style of the workbook is preserved allowing use cases where an Excel template is filled with computed data.

## Track objects

```C#
var products = new ExcelMapper("products.xlsx").Fetch<Product>().ToList();
products[1].Price += 1.0m;
excel.Save("products.out.xlsx");
```

## Ignore properties

```C#
public class Product
{
    public string Name { get; set; }
    [Ignore]
    public int Number { get; set; }
    public decimal Price { get; set; }
}

// or

var excel = new ExcelMapper("products.xlsx");
excel.Ignore<Product>(p => p.Price);
```

## Use specific data formats

```C#
public class Product
{
    [DataFormat(0xf)]
    public DateTime Date { get; set; }

    [DataFormat("0%")]
    public decimal Number { get; set; }
}
```

You can use both [builtin formats](https://poi.apache.org/apidocs/org/apache/poi/ss/usermodel/BuiltinFormats.html) and [custom formats](https://support.office.com/en-nz/article/Create-or-delete-a-custom-number-format-78f2a361-936b-4c03-8772-09fab54be7f4). The default format for DateTime cells is 0x16 ("m/d/yy h:mm").

## Map formulas or results

Formula columns are mapped according to the type of the property they are mapped to: for string properties, the formula itself (e.g. "=A1+B1") is mapped, for other property types the formula result is mapped. If you need the formula result in a string property, use the `FormulaResult` attribute.

```C#
public class Product
{
    [FormulaResult]
    public string Result { get; set; }
}

// or

excel.AddMapping<Product>("Result" p => p.Result).AsFormulaResult();
```

## Custom mapping

If you have specific requirements for mapping between cells and objects, you can use custom conversion methods. Here, cells that contain the string "NULL" are mapped to null:

```C#
public class Product
{
    public DateTime? Date { get; set; }
}

excel.AddMapping<Product>("Date", p => p.Date)
    .SetCellUsing((c, o) =>
    {
        if (o == null) c.SetCellValue("NULL"); else c.SetCellValue((DateTime)o);
    })
    .SetPropertyUsing(v =>
    {
        if ((v as string) == "NULL") return null;
        return Convert.ChangeType(v, typeof(DateTime), CultureInfo.InvariantCulture);
    });
```

## Header row and data row range

You can specify the row number of the header row using the property `HeaderRowNumber` (default is 0). The range of rows that are considered rows that may contain data can be specified using the properties `MinRowNumber` (default is 0) and `MaxRowNumber` (default is `int.MaxValue`). The header row doesn't have to fall within this range, e.g. you can have the header row in row 5 and the data in rows 10-20.

## JSON

You can easily serialize to and from JSON formatted cells by specifying the `Json` attribute or `AsJson()` method.

```c#
public class ProductJson
{
    [Json]
    public Product Product { get; set; }
}

// or

var excel = new ExcelMapper("products.xls");
excel.AddMapping<ProductJson>("Product", p => p.Product).AsJson();
```

This also works with lists.

```c#
public class ProductJson
{
    [Json]
    public List<Product> Products { get; set; }
}
```

## Name normalization

If the header cell values are not uniform, perhaps because they contain varying amounts of whitespace, you can specify a normalization function that will be applied to header cell
values before mapping to property names. This can be done globally or for specific classes only.

```c#
excel.NormalizeUsing(n => Regex.Replace(n, "\w", ""));
```

This removes all whitespace so that columns with the string " First Name " map to a property named `FirstName`.

## Records

Records are supported. 
If the type has no default constructor (as is the case for positional records) the constructor with the highest number of arguments is used to initialize objects. 
This constructor must have a parameter for each of the mapped properties with the same name as the corresponding property (ignoring case). 
The remanining parameters will receive the default value of their type.

## Nested objects

Nested objects are supported and should work out of the box for most use cases. For example, if you have a sheet with columns Name, Street, City, Zip, Birthday, you can map
to the following class hierarchy without any configuration:

```c#
public class Person
{
    public string Name { get; set; }
    public DateTime Birthday { get; set; }
    public Address Address { get; set; }
}

public class Address
{
    public string Street { get; set; }
    public string City { get; set; }
    public string Zip { get; set; }
}

var customers = new ExcelMapper("customers.xlsx").Fetch<Person>();
```

This works with records, too:

```c#
public record Person(string Name, DateTime Birthday, Address Address);
public record Address(string Street, string City, string Zip);
```
