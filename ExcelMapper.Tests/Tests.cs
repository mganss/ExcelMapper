using NPOI.SS.UserModel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Ganss.Excel.Exceptions;
using System.IO;
using System.Threading.Tasks;
using NPOI.XSSF.UserModel;
using NPOI.OpenXmlFormats.Spreadsheet;
using System.Data.Common;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Text.RegularExpressions;
using NPOI.SS.Formula.Functions;
using System.Threading;

namespace Ganss.Excel.Tests
{
    public class Tests
    {
        [SetUp]
        public void Setup()
        {
            Environment.CurrentDirectory = TestContext.CurrentContext.TestDirectory;
        }

        private class ProductMultiColums
        {
            public string Name { get; set; }

            [Column("Number", MappingDirections.ExcelToObject)]// Read as "Number"
            [Column("NewNumber", MappingDirections.ObjectToExcel)]// Write as "NewNumber"
            public int NumberInStock { get; set; }

            [Column(MappingDirections.ExcelToObject)]// Read as "Price"
            [Column("NewPrice", MappingDirections.ObjectToExcel)]// Write as "NewPrice"
            public decimal Price { get; set; }

            [Column(MappingDirections.ExcelToObject)] // Read as "Value"
            [Column("NewValue", MappingDirections.ObjectToExcel)] // Write as "NewValue"
            public string Value { get; set; }
        }

        private class ProductMultiColumsReload
        {
            public string Name { get; set; }
            public int NewNumber { get; set; }
            public decimal NewPrice { get; set; }
            public string NewValue { get; set; }

            public override bool Equals(object obj) =>
                obj is ProductMultiColumsReload reload
                && Name == reload.Name
                && NewNumber == reload.NewNumber
                && NewPrice == reload.NewPrice
                && NewValue == reload.NewValue;

            public override int GetHashCode() =>
                $"{Name}{NewNumber}{NewPrice}{NewValue}".GetHashCode();
        }

        private class ProductDirection
        {
            [Column(MappingDirections.ExcelToObject)]
            public string Name { get; set; }

            [Column("Number", MappingDirections.ExcelToObject)]
            public int NumberInStock { get; set; }

            [Column(MappingDirections.ObjectToExcel)]
            public decimal Price { get; set; }

            [Column(MappingDirections.ObjectToExcel)]
            public string Value { get; set; }

            public override bool Equals(object obj) =>
                obj is ProductDirection o
                && o.Name == Name
                && o.NumberInStock == NumberInStock
                && o.Price == Price
                && o.Value == Value;

            public override int GetHashCode() =>
                $"{Name}{NumberInStock}{Price}{Value}".GetHashCode();
        }

        private class ProductFluent
        {
            public string Name { get; set; }
            public int Number { get; set; }
            public decimal Price { get; set; }
            public string Value { get; set; }

            public override bool Equals(object obj) =>
                obj is ProductFluent o
                && o.Name == Name
                && o.Number == Number
                && o.Price == Price
                && o.Value == Value;

            public override int GetHashCode() =>
                $"{Name}{Number}{Price}{Value}".GetHashCode();
        }
        private class ProductFluentResult : ProductFluent
        { }

        private class Product
        {
            public string Name { get; set; }
            [Column("Number")]
            public int NumberInStock { get; set; }
            public decimal Price { get; set; }
            public string Value { get; set; }

            public override bool Equals(object obj)
            {
                if (obj is not Product o) return false;
                return o.Name == Name && o.NumberInStock == NumberInStock && o.Price == Price && o.Value == Value;
            }

            public override int GetHashCode()
            {
                return (Name + NumberInStock + Price + Value).GetHashCode();
            }
        }

        private class ProductValue
        {
            public decimal Value { get; set; }
        }

        private class ProductValueString
        {
            public decimal Value { get; set; }
            public string ValueDefaultAsFormula { get; set; }
            [FormulaResult]
            public string ValueAsString { get; set; }
        }

        private class BeforeAfterMapping
        {
            public string Name { get; set; }
            public int Number { get; set; }
            public decimal Price { get; set; }
            public string Value { get; set; }

            public int Id { get; set; }
            public string Hash { get; set; }

            public override bool Equals(object obj) =>
                obj is BeforeAfterMapping o
                && o.Name == Name
                && o.Number == Number
                && o.Price == Price
                && o.Value == Value
                && o.Id == Id
                && o.Hash == Hash
                ;

            public override int GetHashCode() =>
                $"{Name}{Number}{Price}{Value}{Id}{Hash}".GetHashCode();
        }

        private class ProductDynamic
        {
            public string Name { get; set; }
            public int Number { get; set; }
            public decimal Price { get; set; }
            public bool Offer { get; set; }
            public DateTime OfferEnd { get; set; }
            public double Value { get; set; }

            public override bool Equals(object obj)
            {
                return obj is ProductDynamic dynamic &&
                       Name == dynamic.Name &&
                       Number == dynamic.Number &&
                       Price == dynamic.Price &&
                       Offer == dynamic.Offer &&
                       OfferEnd == dynamic.OfferEnd &&
                       Value == dynamic.Value;
            }

            public override int GetHashCode()
            {
                return HashCode.Combine(Name, Number, Price, Offer, OfferEnd, Value);
            }
        }
        private class ProductDynamicValueConvertSave : ProductDynamic
        {
            public override bool Equals(object obj)
            {
                return obj is ProductDynamicValueConvertSave save &&
                       Name == save.Name &&
                       Number == save.Number &&
                       Price == save.Price &&
                       Offer == save.Offer &&
                       OfferEnd == save.OfferEnd &&
                       Value == save.Value;
            }

            public override int GetHashCode()
            {
                int hashCode = 1336918815;
                hashCode = hashCode * -1521134295 + base.GetHashCode();
                hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(Name);
                hashCode = hashCode * -1521134295 + Number.GetHashCode();
                hashCode = hashCode * -1521134295 + Price.GetHashCode();
                hashCode = hashCode * -1521134295 + Offer.GetHashCode();
                hashCode = hashCode * -1521134295 + OfferEnd.GetHashCode();
                hashCode = hashCode * -1521134295 + Value.GetHashCode();
                return hashCode;
            }
        }
        private class ProductDynamicValueConvert
        {
            public string Name { get; set; }
            public string Number { get; set; }
            public string Price { get; set; }
            public bool Offer { get; set; }
            public DateTime OfferEnd { get; set; }
            public double Value { get; set; }

            public override bool Equals(object obj)
            {
                return obj is ProductDynamicValueConvert convert &&
                       Name == convert.Name &&
                       Number == convert.Number &&
                       Price == convert.Price &&
                       Offer == convert.Offer &&
                       OfferEnd == convert.OfferEnd &&
                       Value == convert.Value;
            }

            public override int GetHashCode()
            {
                return HashCode.Combine(Name, Number, Price, Offer, OfferEnd, Value);
            }
        }

        private static void CheckDynamicObjectsValueConvert(IEnumerable<dynamic> dynProducts)
        {
            var products = dynProducts.Select(p => new ProductDynamicValueConvert()
            {
                Number = p.Number,
                Price = p.Price,
                Name = p.Name,
                Value = p.Value,
                Offer = p.Offer,
                OfferEnd = p.OfferEnd,
            }).ToList();

            CollectionAssert.AreEqual(new List<ProductDynamicValueConvert>
            {
                new ProductDynamicValueConvert { Name = "-Nudossi-", Number = "3C", Price = "/1.99/", Value = 119.40, Offer = false, OfferEnd = new DateTime(1970, 01, 01) },
                new ProductDynamicValueConvert { Name = "-Halloren-", Number = "21", Price = "/2.99/", Value = 98.67, Offer = true, OfferEnd = new DateTime(2015, 12, 31) },
                new ProductDynamicValueConvert { Name = "-Filinchen-", Number = "64", Price = "/0.99/", Value = 99.00, Offer = false, OfferEnd = new DateTime(1970, 01, 01) },
            }, products);
        }

        [Test]
        public async Task FetchSaveValueConverterOverloadsTest()
        {
            var file = @"../../../xlsx/Products.xlsx";
            var excel = new ExcelMapper();

            object valueParser(string colname, object val)
            {
                return colname switch
                {
                    // Readers
                    "Number" when val is double dval => ((int)dval).ToString("X"),
                    "Price" when val is double dval => string.Format(CultureInfo.InvariantCulture, "/{0}/", dval),
                    "Name" when val is string sval => $"-{sval}-",
                    _ => val,
                };
            }

            object valueConverter(string colname, object val)
            {
                return colname switch
                {
                    // Writers
                    "Number" when val is string sval => int.Parse(sval, NumberStyles.HexNumber),
                    "Price" when val is string sval => decimal.Parse(sval.Replace("/", string.Empty), CultureInfo.InvariantCulture),
                    "Name" when val is string sval && sval[0] == '-' && sval[^1] == '-' => sval.Replace("-", string.Empty),
                    _ => val,
                };
            }

            var products = await excel.FetchAsync(file, "Tabelle1", valueParser);
            CheckDynamicObjectsValueConvert(products);

            products = await excel.FetchAsync(file, 0, valueParser);
            CheckDynamicObjectsValueConvert(products);

            var stream = new FileStream(file, FileMode.Open, FileAccess.Read);
            products = await excel.FetchAsync(stream, "Tabelle1", valueParser);
            stream.Close();
            CheckDynamicObjectsValueConvert(products);

            stream = new FileStream(file, FileMode.Open, FileAccess.Read);
            products = await excel.FetchAsync(stream, 0, valueParser);
            stream.Close();
            CheckDynamicObjectsValueConvert(products);

            // Save
            var expectedResult = new List<ProductDynamicValueConvertSave>
            {
                new ProductDynamicValueConvertSave { Name = "Nudossi", Number = 60, Price = 1.99m, Value = 119.40, Offer = false, OfferEnd = new DateTime(1970, 01, 01) },
                new ProductDynamicValueConvertSave { Name = "Halloren", Number = 33, Price = 2.99m, Value = 98.67, Offer = true, OfferEnd = new DateTime(2015, 12, 31) },
                new ProductDynamicValueConvertSave { Name = "Filinchen", Number = 100, Price = 0.99m, Value = 99.00, Offer = false, OfferEnd = new DateTime(1970, 01, 01) },
            };

            var filesave = "productssave_valueconverter.xlsx";

            await new ExcelMapper().SaveAsync(filesave, products, "Products", valueConverter: valueConverter);
            var productsFetched = new ExcelMapper(filesave).Fetch<ProductDynamicValueConvertSave>().ToList();
            CollectionAssert.AreEqual(expectedResult, productsFetched);

            await new ExcelMapper().SaveAsync(filesave, products, valueConverter: valueConverter);
            productsFetched = new ExcelMapper(filesave).Fetch<ProductDynamicValueConvertSave>().ToList();
            CollectionAssert.AreEqual(expectedResult, productsFetched);

            using (var fs = File.OpenWrite(filesave))
            {
                await new ExcelMapper().SaveAsync(fs, products, "Products", valueConverter: valueConverter);
            }
            productsFetched = new ExcelMapper(filesave).Fetch<ProductDynamicValueConvertSave>().ToList();
            CollectionAssert.AreEqual(expectedResult, productsFetched);

            using (var fs = File.OpenWrite(filesave))
            {
                await new ExcelMapper().SaveAsync(fs, products, valueConverter: valueConverter);
            }
            productsFetched = new ExcelMapper(filesave).Fetch<ProductDynamicValueConvertSave>().ToList();
            CollectionAssert.AreEqual(expectedResult, productsFetched);
        }

        [Test]
        public void BeforeAfterMappingTest()
        {
            var products = new ExcelMapper(@"../../../xlsx/Products.xlsx")
                // preparation before the mapping start
                .AddBeforeMapping<BeforeAfterMapping>((obj, idx) =>
                    obj.Id = idx + 1000
                )
                // Apply late mapping after every Excel to Object row mapping
                .AddAfterMapping<BeforeAfterMapping>((obj, idx) =>
                {
                    obj.Hash = $"{obj.Name}:{obj.Number}:{obj.Id}";
                })
                .Fetch<BeforeAfterMapping>().ToList();

            CollectionAssert.AreEqual(new List<BeforeAfterMapping>
            {
                new BeforeAfterMapping { Name = "Nudossi", Number = 60, Price = 1.99m, Value = "C2*D2"
                    , Id = 1000, Hash = $"Nudossi:60:1000"
                },
                new BeforeAfterMapping { Name = "Halloren", Number = 33, Price = 2.99m, Value = "C3*D3"
                    , Id = 1001, Hash = $"Halloren:33:1001"
                },
                new BeforeAfterMapping { Name = "Filinchen", Number = 100, Price = 0.99m, Value = "C5*D5"
                    , Id = 1002, Hash = $"Filinchen:100:1002"
                },
            }, products);
        }

        [Test]
        public void MultiDirectionalTest()
        {
            /// Reading using <see cref="MappingDirections.ExcelToObject"/> direction mapping
            var products = new ExcelMapper(@"../../../xlsx/Products.xlsx").Fetch<ProductMultiColums>().ToList();

            var file = "productssave_multicolums.xlsx";

            /// Saving using <see cref="MappingDirections.ObjectToExcel"/> direction mapping
            new ExcelMapper().Save(file, products, "Products");

            /// reload excel with <see cref="ProductMultiColumsReload"/> mapping instead of <see cref="ProductMultiColums"/>
            var reloaded = new ExcelMapper(file).Fetch<ProductMultiColumsReload>().ToList();

            CollectionAssert.AreEqual(new List<ProductMultiColumsReload>
            {
                new ProductMultiColumsReload { Name = "Nudossi", NewNumber = 60, NewPrice = 1.99m, NewValue = "C2*D2" },
                new ProductMultiColumsReload { Name = "Halloren", NewNumber = 33, NewPrice = 2.99m, NewValue = "C3*D3" },
                new ProductMultiColumsReload { Name = "Filinchen", NewNumber = 100, NewPrice = 0.99m, NewValue = "C5*D5" },
            }, reloaded);
        }

        [Test]
        public void FetchDynamicTest()
        {
            var products = new ExcelMapper(@"../../../xlsx/Products.xlsx").Fetch().ToList();

            var result = new List<ProductDynamic>();
            foreach (var p in products)
            {
                var dicoRow = p as IDictionary<string, object>;// Need underlying Dictionary for ColumnIndex value recovery
                var r = new ProductDynamic()
                {
                    // Using dynamic notation
                    Number = (int)p.Number,// Need explicit casting for CellType.Numeric when different from double
                    Price = (decimal)p.Price,

                    // using Dictionary notation
                    Name = dicoRow["Name"] as string,

                    Value = (double)dicoRow["Value"],

                    // No casting
                    Offer = p.Offer,
                    OfferEnd = p.OfferEnd,
                };
                result.Add(r);
            }

            CollectionAssert.AreEqual(new List<ProductDynamic>
            {
                new ProductDynamic { Name = "Nudossi", Number = 60, Price = 1.99m, Value = 119.40, Offer = false, OfferEnd = new DateTime(1970, 01, 01) },
                new ProductDynamic { Name = "Halloren", Number = 33, Price = 2.99m, Value = 98.67, Offer = true, OfferEnd = new DateTime(2015, 12, 31) },
                new ProductDynamic { Name = "Filinchen", Number = 100, Price = 0.99m, Value = 99.00, Offer = false, OfferEnd = new DateTime(1970, 01, 01) },
            }, result);
        }

        [Test]
        public void FetchDynamicIndexTest()
        {
            var products = new ExcelMapper(@"../../../xlsx/Products.xlsx") { HeaderRow = false, MinRowNumber = 1 }.Fetch().ToList();

            var result = new List<ProductDynamic>();
            foreach (var p in products)
            {
                var r = new ProductDynamic()
                {
                    Name = p.A,
                    Number = (int)p.C,
                    Price = (decimal)p.D,
                    Offer = p.E,
                    OfferEnd = p.F,
                    Value = p.G,
                };
                result.Add(r);
            }

            CollectionAssert.AreEqual(new List<ProductDynamic>
            {
                new ProductDynamic { Name = "Nudossi", Number = 60, Price = 1.99m, Value = 119.40, Offer = false, OfferEnd = new DateTime(1970, 01, 01) },
                new ProductDynamic { Name = "Halloren", Number = 33, Price = 2.99m, Value = 98.67, Offer = true, OfferEnd = new DateTime(2015, 12, 31) },
                new ProductDynamic { Name = "Filinchen", Number = 100, Price = 0.99m, Value = 99.00, Offer = false, OfferEnd = new DateTime(1970, 01, 01) },
            }, result);
        }

        [Test]
        public void FetchDynamicSaveTest()
        {
            var excel = new ExcelMapper(@"../../../xlsx/Products.xlsx");
            var dynProducts = excel.Fetch().ToList();
            var products = dynProducts.Select(p => new ProductDynamic()
            {
                Number = (int)p.Number,
                Price = (decimal)p.Price,
                Name = p.Name,
                Value = p.Value,
                Offer = p.Offer,
                OfferEnd = p.OfferEnd,
            }).ToList();

            CollectionAssert.AreEqual(new List<ProductDynamic>
            {
                new ProductDynamic { Name = "Nudossi", Number = 60, Price = 1.99m, Value = 119.40, Offer = false, OfferEnd = new DateTime(1970, 01, 01) },
                new ProductDynamic { Name = "Halloren", Number = 33, Price = 2.99m, Value = 98.67, Offer = true, OfferEnd = new DateTime(2015, 12, 31) },
                new ProductDynamic { Name = "Filinchen", Number = 100, Price = 0.99m, Value = 99.00, Offer = false, OfferEnd = new DateTime(1970, 01, 01) },
            }, products);

            dynProducts[0].Name += "Test";
            dynProducts[1].Price += 1.0;
            dynProducts[2].OfferEnd = new DateTime(2000, 1, 2);

            var file = @"productssavedynamic.xlsx";

            excel.Save(file);

            var productsFetched = new ExcelMapper(file).Fetch<ProductDynamic>().ToList();

            products[0].Name += "Test";
            products[1].Price += 1.0m;
            products[2].OfferEnd = new DateTime(2000, 1, 2);

            CollectionAssert.AreEqual(products, productsFetched);
        }

        [Test]
        public void FetchDynamicIndexSaveTest()
        {
            var excel = new ExcelMapper(@"../../../xlsx/Products.xlsx") { HeaderRow = false, MinRowNumber = 1 };
            var dynProducts = excel.Fetch().ToList();
            var products = dynProducts.Select(p => new ProductDynamic()
            {
                Name = p.A,
                Number = (int)p.C,
                Price = (decimal)p.D,
                Offer = p.E,
                OfferEnd = p.F,
                Value = p.G,
            }).ToList();

            CollectionAssert.AreEqual(new List<ProductDynamic>
            {
                new ProductDynamic { Name = "Nudossi", Number = 60, Price = 1.99m, Value = 119.40, Offer = false, OfferEnd = new DateTime(1970, 01, 01) },
                new ProductDynamic { Name = "Halloren", Number = 33, Price = 2.99m, Value = 98.67, Offer = true, OfferEnd = new DateTime(2015, 12, 31) },
                new ProductDynamic { Name = "Filinchen", Number = 100, Price = 0.99m, Value = 99.00, Offer = false, OfferEnd = new DateTime(1970, 01, 01) },
            }, products);

            dynProducts[0].A += "Test";
            dynProducts[1].D += 1.0;
            dynProducts[2].F = new DateTime(2000, 1, 2);

            // remove index map to test automatic letter mapping on save
            foreach (var d in dynProducts)
                d.__indexes__ = null;

            var file = @"productssavedynamicindex.xlsx";

            excel.Save(file);

            var productsFetched = new ExcelMapper(file).Fetch<ProductDynamic>().ToList();

            products[0].Name += "Test";
            products[1].Price += 1.0m;
            products[2].OfferEnd = new DateTime(2000, 1, 2);

            CollectionAssert.AreEqual(products, productsFetched);
        }

        [Test]
        public void FetchDynamicSaveObjectsTest()
        {
            var excel = new ExcelMapper(@"../../../xlsx/Products.xlsx");
            var dynProducts = excel.Fetch().ToList();
            var products = dynProducts.Select(p => new ProductDynamic()
            {
                Number = (int)p.Number,
                Price = (decimal)p.Price,
                Name = p.Name,
                Value = p.Value,
                Offer = p.Offer,
                OfferEnd = p.OfferEnd,
            }).ToList();

            CollectionAssert.AreEqual(new List<ProductDynamic>
            {
                new ProductDynamic { Name = "Nudossi", Number = 60, Price = 1.99m, Value = 119.40, Offer = false, OfferEnd = new DateTime(1970, 01, 01) },
                new ProductDynamic { Name = "Halloren", Number = 33, Price = 2.99m, Value = 98.67, Offer = true, OfferEnd = new DateTime(2015, 12, 31) },
                new ProductDynamic { Name = "Filinchen", Number = 100, Price = 0.99m, Value = 99.00, Offer = false, OfferEnd = new DateTime(1970, 01, 01) },
            }, products);

            dynProducts[0].Name += "Test";
            dynProducts[1].Price += 1.0;
            dynProducts[2].OfferEnd = new DateTime(2000, 1, 2);

            var file = @"productssavedynamic.xlsx";

            new ExcelMapper().Save(file, dynProducts);

            var productsFetched = new ExcelMapper(file).Fetch<ProductDynamic>().ToList();

            products[0].Name += "Test";
            products[1].Price += 1.0m;
            products[2].OfferEnd = new DateTime(2000, 1, 2);

            CollectionAssert.AreEqual(products, productsFetched);
        }

        [Test]
        public void FetchDynamicOverloadsTest()
        {
            var file = @"../../../xlsx/Products.xlsx";
            var excel = new ExcelMapper();

            var products = excel.Fetch(file, "Tabelle1");
            CheckDynamicObjects(products);

            products = excel.Fetch(file, 0);
            CheckDynamicObjects(products);

            var stream = new FileStream(file, FileMode.Open, FileAccess.Read);
            products = excel.Fetch(stream, "Tabelle1");
            stream.Close();
            CheckDynamicObjects(products);

            stream = new FileStream(file, FileMode.Open, FileAccess.Read);
            products = excel.Fetch(stream, 0);
            stream.Close();
            CheckDynamicObjects(products);

            excel = new ExcelMapper(file);
            products = excel.Fetch("Tabelle1");
            CheckDynamicObjects(products);

            excel = new ExcelMapper(file);
            products = excel.Fetch(0);
            CheckDynamicObjects(products);
        }

        [Test]
        public void FetchAsyncDynamicOverloadsTest()
        {
            var file = @"../../../xlsx/Products.xlsx";
            var excel = new ExcelMapper();

            var products = excel.FetchAsync(file, "Tabelle1").Result;
            CheckDynamicObjects(products);

            products = excel.FetchAsync(file, 0).Result;
            CheckDynamicObjects(products);

            var stream = new FileStream(file, FileMode.Open, FileAccess.Read);
            products = excel.FetchAsync(stream, "Tabelle1").Result;
            stream.Close();
            CheckDynamicObjects(products);

            stream = new FileStream(file, FileMode.Open, FileAccess.Read);
            products = excel.FetchAsync(stream, 0).Result;
            stream.Close();
            CheckDynamicObjects(products);
        }

        private static void CheckDynamicObjects(IEnumerable<dynamic> dynProducts)
        {
            var products = dynProducts.Select(p => new ProductDynamic()
            {
                Number = (int)p.Number,
                Price = (decimal)p.Price,
                Name = p.Name,
                Value = p.Value,
                Offer = p.Offer,
                OfferEnd = p.OfferEnd,
            }).ToList();

            CollectionAssert.AreEqual(new List<ProductDynamic>
            {
                new ProductDynamic { Name = "Nudossi", Number = 60, Price = 1.99m, Value = 119.40, Offer = false, OfferEnd = new DateTime(1970, 01, 01) },
                new ProductDynamic { Name = "Halloren", Number = 33, Price = 2.99m, Value = 98.67, Offer = true, OfferEnd = new DateTime(2015, 12, 31) },
                new ProductDynamic { Name = "Filinchen", Number = 100, Price = 0.99m, Value = 99.00, Offer = false, OfferEnd = new DateTime(1970, 01, 01) },
            }, products);
        }

        [Test]
        public void FromExcelOnlyTest()
        {
            var products = new ExcelMapper(@"../../../xlsx/Products.xlsx").Fetch<ProductDirection>().ToList();
            CollectionAssert.AreEqual(new List<ProductDirection>
            {
                new ProductDirection{ Name = "Nudossi", NumberInStock = 60, Price = 0, Value = null },
                new ProductDirection{ Name = "Halloren", NumberInStock = 33, Price = 0, Value = null },
                new ProductDirection{ Name = "Filinchen", NumberInStock = 100, Price = 0, Value = null },
            }, products);
        }

        [Test]
        public void ToExcelOnlyTest()
        {
            var src = new List<ProductDirection>
            {
                new ProductDirection {
                    // FromExcelOnly
                    Name = "Nudossi", NumberInStock = 60
                    // ToExcelOnly
                    , Price = 1.99m, Value = "C2*D2"
                },
                new ProductDirection { Name = "Halloren", NumberInStock = 33, Price = 2.99m, Value = "C3*D3" },
                new ProductDirection { Name = "Filinchen", NumberInStock = 100, Price = 0.99m, Value = "C5*D5" },
            };

            var file = "productssavetoexcelonly.xlsx";

            new ExcelMapper().Save(file, src, "Products");

            /// Read result with <see cref="Product"/> mapping instead of <see cref="ProductDirection"/>
            var productsFetched = new ExcelMapper(file).Fetch<Product>().ToList();

            CollectionAssert.AreEqual(new List<Product>
            {
                new Product {
                    // FromExcelOnly prevent excel saving
                    Name = null, NumberInStock = 0
                    // ToExcelOnly allow saving but prevent reading
                    , Price = 1.99m, Value = "C2*D2"
                },
                new Product { Name = null, NumberInStock = 0, Price = 2.99m, Value = "C3*D3" },
                new Product { Name = null, NumberInStock = 0, Price = 0.99m, Value = "C5*D5" },
            }, productsFetched);
        }

        [Test]
        public void ToExcelOnlyFluentTest()
        {
            var src = new List<ProductFluent>
            {
                new ProductFluent { Name = "Nudossi", Number = 60, Price = 1.99m, Value = "C2*D2" },
                new ProductFluent { Name = "Halloren", Number = 33, Price = 2.99m, Value = "C3*D3" },
                new ProductFluent { Name = "Filinchen", Number = 100, Price = 0.99m, Value = "C5*D5" },
            };

            var file = "productssavetoexcelonly_fluent.xlsx";
            var mapperwrite = new ExcelMapper();

            // Make mapper unable to read Name & Value from Excel. Only write.
            mapperwrite.AddMapping<ProductFluent>("Name", p => p.Name).ToExcelOnly();
            mapperwrite.AddMapping<ProductFluent>("Value", p => p.Value).ToExcelOnly();

            // Make mapper unable to write Number & Price from Excel. Only read.
            mapperwrite.AddMapping<ProductFluent>("Number", p => p.Number).FromExcelOnly();
            mapperwrite.AddMapping<ProductFluent>("Price", p => p.Price).FromExcelOnly();

            mapperwrite.Save(file, src, "Products");

            // Reload rows
            var productsFetched = new ExcelMapper(file).Fetch<ProductFluentResult>().ToList();

            CollectionAssert.AreEqual(new List<ProductFluentResult>
            {
                new ProductFluentResult { Name = "Nudossi", Number = 0, Price = 0, Value = "C2*D2" },
                new ProductFluentResult { Name = "Halloren", Number = 0, Price = 0, Value = "C3*D3" },
                new ProductFluentResult { Name = "Filinchen", Number = 0, Price = 0, Value = "C5*D5" },
            }, productsFetched);
        }

        [Test]
        public void FetchTest()
        {
            var products = new ExcelMapper(@"../../../xlsx/Products.xlsx").Fetch<Product>().ToList();
            CollectionAssert.AreEqual(new List<Product>
            {
                new Product { Name = "Nudossi", NumberInStock = 60, Price = 1.99m, Value = "C2*D2" },
                new Product { Name = "Halloren", NumberInStock = 33, Price = 2.99m, Value = "C3*D3" },
                new Product { Name = "Filinchen", NumberInStock = 100, Price = 0.99m, Value = "C5*D5" },
            }, products);
        }

        [Test]
        public void FetchWithTypeTest()
        {
            var products = new ExcelMapper(@"../../../xlsx/Products.xlsx").Fetch(typeof(Product));
            CollectionAssert.AreEqual(new List<Product>
            {
                new Product { Name = "Nudossi", NumberInStock = 60, Price = 1.99m, Value = "C2*D2" },
                new Product { Name = "Halloren", NumberInStock = 33, Price = 2.99m, Value = "C3*D3" },
                new Product { Name = "Filinchen", NumberInStock = 100, Price = 0.99m, Value = "C5*D5" },
            }, products);
        }

        [Test]
        public void FetchWithStreamAndIndexTest()
        {
            var stream = new FileStream(@"../../../xlsx/Products.xlsx", FileMode.Open, FileAccess.Read);
            var products = new ExcelMapper().Fetch<Product>(stream, 0);

            CollectionAssert.AreEqual(new List<Product>
            {
                new Product { Name = "Nudossi", NumberInStock = 60, Price = 1.99m, Value = "C2*D2" },
                new Product { Name = "Halloren", NumberInStock = 33, Price = 2.99m, Value = "C3*D3" },
                new Product { Name = "Filinchen", NumberInStock = 100, Price = 0.99m, Value = "C5*D5" },
            }, products);
            stream.Close();
        }

        [Test]
        public void FetchWithTypeUsingStreamAndIndexTest()
        {
            var stream = new FileStream(@"../../../xlsx/Products.xlsx", FileMode.Open, FileAccess.Read);
            var products = new ExcelMapper().Fetch(stream, typeof(Product), 0);

            CollectionAssert.AreEqual(new List<Product>
            {
                new Product { Name = "Nudossi", NumberInStock = 60, Price = 1.99m, Value = "C2*D2" },
                new Product { Name = "Halloren", NumberInStock = 33, Price = 2.99m, Value = "C3*D3" },
                new Product { Name = "Filinchen", NumberInStock = 100, Price = 0.99m, Value = "C5*D5" },
            }, products);
            stream.Close();
        }

        [Test]
        public void FetchWithStreamAndSheetNameTest()
        {
            var stream = new FileStream(@"../../../xlsx/Products.xlsx", FileMode.Open, FileAccess.Read);
            var products = new ExcelMapper().Fetch<Product>(stream, "Tabelle1");

            CollectionAssert.AreEqual(new List<Product>
            {
                new Product { Name = "Nudossi", NumberInStock = 60, Price = 1.99m, Value = "C2*D2" },
                new Product { Name = "Halloren", NumberInStock = 33, Price = 2.99m, Value = "C3*D3" },
                new Product { Name = "Filinchen", NumberInStock = 100, Price = 0.99m, Value = "C5*D5" },
            }, products);
            stream.Close();
        }

        [Test]
        public void FetchWithTypeUsingStreamAndSheetNameTest()
        {
            var stream = new FileStream(@"../../../xlsx/Products.xlsx", FileMode.Open, FileAccess.Read);
            var products = new ExcelMapper().Fetch(stream, typeof(Product), "Tabelle1");

            CollectionAssert.AreEqual(new List<Product>
            {
                new Product { Name = "Nudossi", NumberInStock = 60, Price = 1.99m, Value = "C2*D2" },
                new Product { Name = "Halloren", NumberInStock = 33, Price = 2.99m, Value = "C3*D3" },
                new Product { Name = "Filinchen", NumberInStock = 100, Price = 0.99m, Value = "C5*D5" },
            }, products);
            stream.Close();

        }

        [Test]
        public void FetchWithFileAndSheetNameTest()
        {
            var products = new ExcelMapper().Fetch<Product>(@"../../../xlsx/Products.xlsx", "Tabelle1");

            CollectionAssert.AreEqual(new List<Product>
            {
                new Product { Name = "Nudossi", NumberInStock = 60, Price = 1.99m, Value = "C2*D2" },
                new Product { Name = "Halloren", NumberInStock = 33, Price = 2.99m, Value = "C3*D3" },
                new Product { Name = "Filinchen", NumberInStock = 100, Price = 0.99m, Value = "C5*D5" },
            }, products);
        }

        [Test]
        public void FetchWithFileAndIndexTest()
        {
            var products = new ExcelMapper().Fetch<Product>(@"../../../xlsx/Products.xlsx", 0);

            CollectionAssert.AreEqual(new List<Product>
            {
                new Product { Name = "Nudossi", NumberInStock = 60, Price = 1.99m, Value = "C2*D2" },
                new Product { Name = "Halloren", NumberInStock = 33, Price = 2.99m, Value = "C3*D3" },
                new Product { Name = "Filinchen", NumberInStock = 100, Price = 0.99m, Value = "C5*D5" },
            }, products);
        }

        [Test]
        public void FetchWithTypeThrowsExceptionWithPrimitivesTest()
        {
            var excel = new ExcelMapper(@"../../../xlsx/Products.xlsx");
            Assert.Throws<ArgumentException>(() => excel.Fetch(typeof(string)));
            Assert.Throws<ArgumentException>(() => excel.Fetch(typeof(object)));
            Assert.Throws<ArgumentException>(() => excel.Fetch(typeof(int)));
            Assert.Throws<ArgumentException>(() => excel.Fetch(typeof(double?)));
        }

        [Test]
        public void FetchValueTest()
        {
            var products = new ExcelMapper(@"../../../xlsx/Products.xlsx").Fetch<ProductValue>().ToList();
            CollectionAssert.AreEqual(new List<decimal> { 119.4m, 98.67m, 99m }, products.Select(p => p.Value).ToList());
        }

        [Test]
        public void FetchValueWithTypeTest()
        {
            var products = new ExcelMapper(@"../../../xlsx/Products.xlsx").Fetch(typeof(ProductValue))
                                                                     .OfType<ProductValue>()
                                                                     .ToList();
            CollectionAssert.AreEqual(new List<decimal> { 119.4m, 98.67m, 99m }, products.Select(p => p.Value).ToList());
        }

        private class ProductException : Product
        {
            public bool Offer { get; set; }
        }

        [Test]
        public void FetchEmptyTest()
        {
            var products = new ExcelMapper(@"../../../xlsx/ProductsExceptionEmpty.xlsx").Fetch<ProductException>().ToList();
            CollectionAssert.AreEqual(new List<ProductException>
            {
                new ProductException { Name = "Nudossi", NumberInStock = 60, Price = 0m },
            }, products);
        }

        [Test]
        public void FetchExceptionWhenEmptyTest()
        {
            var ex = Assert.Throws<ExcelMapperConvertException>(() => new ExcelMapper(@"../../../xlsx/ProductsExceptionEmpty.xlsx") { SkipBlankCells = false }.Fetch<ProductException>().ToList());
            Assert.That(ex.Message.Contains("<EMPTY>"));
            Assert.That(ex.Message.Contains("[L:1]:[C:2]"));
        }

        [Test]
        public void FetchWithTypeExceptionWhenEmptyTest()
        {
            var ex = Assert.Throws<ExcelMapperConvertException>(() => new ExcelMapper(@"../../../xlsx/ProductsExceptionEmpty.xlsx") { SkipBlankCells = false }.Fetch(typeof(ProductException))
                                                                                                                              .OfType<ProductException>()
                                                                                                                              .ToList());
            Assert.That(ex.Message.Contains("<EMPTY>"));
            Assert.That(ex.Message.Contains("[L:1]:[C:2]"));
        }

        [Test]
        public void FetchExceptionWhenFieldTooBigTest()
        {
            var ex = Assert.Throws<ExcelMapperConvertException>(() => new ExcelMapper(@"../../../xlsx/ProductsExceptionTooBig.xlsx").Fetch<ProductException>().ToList());
            //2147483649 is Int.MaxValue + 1
            Assert.That(ex.Message.Contains("2147483649"));
            Assert.That(ex.Message.Contains("[L:1]:[C:1]"));
        }

        [Test]
        public void FetchExceptionWhenFieldInvalidTest()
        {
            var ex = Assert.Throws<ExcelMapperConvertException>(() => new ExcelMapper(@"../../../xlsx/ProductsExceptionInvalid.xlsx").Fetch<ProductException>().ToList());
            Assert.That(ex.Message.Contains("FALSEd"));
            Assert.That(ex.Message.Contains("[L:1]:[C:3]"));
        }

        [Test]
        public void FetchEventErrorWhenFieldInvalidTest()
        {
            int numberOfErrors = 0;
            var excelMapper = new ExcelMapper(@"../../../xlsx/ProductsEventsExceptionInvalid.xlsx");
            excelMapper.ErrorParsingCell += (sender, e) =>
            {
                Assert.IsInstanceOf<ParsingErrorEventArgs>(e);
                Assert.IsInstanceOf<ExcelMapperConvertException>(e.Error);
                Assert.IsNotNull(e.Error);
                Assert.That(e.Error.Message.Contains("FALSEd"));
                numberOfErrors++;

                e.Cancel = true;
            };

            List<ProductException> listOfProducts = null;
            Assert.DoesNotThrow(() => listOfProducts = excelMapper.Fetch<ProductException>().ToList());
            Assert.That(numberOfErrors == 5);
            Assert.That(listOfProducts.Count == 6);
        }

        [Test]
        public void FetchEventExceptionExplicitlyDisabledWhenFieldInvalidTest()
        {
            var excelMapper = new ExcelMapper(@"../../../xlsx/ProductsEventsExceptionInvalid.xlsx");
            excelMapper.ErrorParsingCell += (sender, e) =>
            {
                e.Cancel = false;
            };

            List<ProductException> listOfProducts = null;
            var ex = Assert.Throws<ExcelMapperConvertException>(() => listOfProducts = excelMapper.Fetch<ProductException>().ToList());
            Assert.IsNull(listOfProducts);
        }

        [Test]
        public void FetchExceptionWhenSheetDoesNotExists()
        {
            var ex = Assert.Throws<ArgumentOutOfRangeException>(() => new ExcelMapper(@"../../../xlsx/ProductsExceptionInvalid.xlsx").Fetch<ProductException>("this sheet does not exist").ToList());
            Assert.That(ex.Message.Contains("Sheet not found"));
        }

        [Test]
        public void FetchWithTypeThrowsExceptionWhenSheetDoesNotExists()
        {
            var ex = Assert.Throws<ArgumentOutOfRangeException>(() => new ExcelMapper(@"../../../xlsx/ProductsExceptionInvalid.xlsx").Fetch(typeof(ProductException), "This is not a exist")
                                                                                                                                .OfType<ProductException>()
                                                                                                                                .ToList());
            Assert.That(ex.Message.Contains("Sheet not found"));
        }

        [Test]
        public void FetchSheetNamesTest()
        {
            var excel = new ExcelMapper(@"../../../xlsx/Products.xlsx");
            var sheetNames = excel.FetchSheetNames().ToList();

            CollectionAssert.AreEqual(new List<string>
            {
                "Tabelle1",
                "Tabelle2",
                "Tabelle3",
            }, sheetNames);
        }

        [Test]
        public void FetchSheetNamesHiddenTest()
        {
            var excel = new ExcelMapper(@"../../../xlsx/ProductsHidden.xlsx");
            var sheetNames = excel.FetchSheetNames(ignoreHidden: true).ToList();

            CollectionAssert.AreEqual(new List<string>
            {
                "Tabelle2",
                "Tabelle3",
            }, sheetNames);
        }

        [Test]
        public void FetchSheetNamesEmptyTest()
        {
            var excel = new ExcelMapper();
            CollectionAssert.IsEmpty(excel.FetchSheetNames());
        }

        private class ProductNoHeader
        {
            [Column(1)]
            public string Name { get; set; }
            [Column(3)]
            public int NumberInStock { get; set; }
            [Column(4)]
            public decimal Price { get; set; }

            public override bool Equals(object obj)
            {
                if (obj is not ProductNoHeader o) return false;
                return o.Name == Name && o.NumberInStock == NumberInStock && o.Price == Price;
            }

            public override int GetHashCode()
            {
                return (Name + NumberInStock + Price).GetHashCode();
            }
        }

        [Test]
        public void FetchNoHeaderTest()
        {
            var products = new ExcelMapper(@"../../../xlsx/ProductsNoHeader.xlsx") { HeaderRow = false }.Fetch<ProductNoHeader>("Products").ToList();
            CollectionAssert.AreEqual(new List<ProductNoHeader>
            {
                new ProductNoHeader { Name = "Nudossi", NumberInStock = 60, Price = 1.99m },
                new ProductNoHeader { Name = "Halloren", NumberInStock = 33, Price = 2.99m },
                new ProductNoHeader { Name = "Filinchen", NumberInStock = 100, Price = 0.99m },
            }, products);
        }

        [Test]
        public void FetchWithTypeNoHeaderTest()
        {
            var products = new ExcelMapper(@"../../../xlsx/ProductsNoHeader.xlsx") { HeaderRow = false }.Fetch(typeof(ProductNoHeader), "Products")
                                                                                                   .OfType<ProductNoHeader>()
                                                                                                   .ToList();
            CollectionAssert.AreEqual(new List<ProductNoHeader>
            {
                new ProductNoHeader { Name = "Nudossi", NumberInStock = 60, Price = 1.99m },
                new ProductNoHeader { Name = "Halloren", NumberInStock = 33, Price = 2.99m },
                new ProductNoHeader { Name = "Filinchen", NumberInStock = 100, Price = 0.99m },
            }, products);
        }

        private class ProductNoHeaderManual
        {
            public string NameX { get; set; }
            public int NumberInStockX { get; set; }
            public decimal PriceX { get; set; }

            public override bool Equals(object obj)
            {
                if (obj is not ProductNoHeaderManual o) return false;
                return o.NameX == NameX && o.NumberInStockX == NumberInStockX && o.PriceX == PriceX;
            }

            public override int GetHashCode()
            {
                return (NameX + NumberInStockX + PriceX).GetHashCode();
            }
        }

        [Test]
        public void FetchNoHeaderManualTest()
        {
            var excel = new ExcelMapper(@"../../../xlsx/ProductsNoHeader.xlsx") { HeaderRow = false };

            excel.AddMapping<ProductNoHeaderManual>(1, p => p.NameX);
            excel.AddMapping<ProductNoHeaderManual>(ExcelMapper.LetterToIndex("C"), p => p.NumberInStockX);
            excel.AddMapping(typeof(ProductNoHeaderManual), 4, "PriceX");

            var products = excel.Fetch<ProductNoHeaderManual>("Products").ToList();

            CollectionAssert.AreEqual(new List<ProductNoHeaderManual>
            {
                new ProductNoHeaderManual { NameX = "Nudossi", NumberInStockX = 60, PriceX = 1.99m },
                new ProductNoHeaderManual { NameX = "Halloren", NumberInStockX = 33, PriceX = 2.99m },
                new ProductNoHeaderManual { NameX = "Filinchen", NumberInStockX = 100, PriceX = 0.99m },
            }, products);
        }

        [Test]
        public void SaveTest()
        {
            var products = new List<Product>
            {
                new Product { Name = "Nudossi", NumberInStock = 60, Price = 1.99m, Value = "C2*D2" },
                new Product { Name = "Halloren", NumberInStock = 33, Price = 2.99m, Value = "C3*D3" },
                new Product { Name = "Filinchen", NumberInStock = 100, Price = 0.99m, Value = "C4*D4" },
            };

            var file = "productssave.xlsx";
            var excelMapper = new ExcelMapper();

            excelMapper.Saving += (s, e) =>
            {
                var cols = e.Sheet.GetRow(excelMapper.HeaderRowNumber).LastCellNum;

                for (int i = 0; i < cols; i++)
                    e.Sheet.AutoSizeColumn(i);
            };

            excelMapper.Save(file, products, "Products");

            var productsFetched = new ExcelMapper(file).Fetch<Product>().ToList();

            CollectionAssert.AreEqual(products, productsFetched);
        }

        [Test]
        public void SaveNoHeaderSaveTest()
        {
            var products = new List<ProductNoHeader>
            {
                new ProductNoHeader { Name = "Nudossi", NumberInStock = 60, Price = 1.99m },
                new ProductNoHeader { Name = "Halloren", NumberInStock = 33, Price = 2.99m },
                new ProductNoHeader { Name = "Filinchen", NumberInStock = 100, Price = 0.99m },
            };

            var file = "productsnoheadersave.xlsx";

            new ExcelMapper() { HeaderRow = false }.Save(file, products, "Products");

            var productsFetched = new ExcelMapper(file) { HeaderRow = false }.Fetch<ProductNoHeader>().ToList();

            CollectionAssert.AreEqual(products, productsFetched);
        }

        [Test]
        public void SaveFetchedTest()
        {
            var excel = new ExcelMapper(@"../../../xlsx/Products.xlsx");
            var products = excel.Fetch<Product>().ToList();

            products[2].Price += 1.0m;

            var file = @"productssavefetched.xlsx";

            excel.Save(file, products);

            var productsFetched = new ExcelMapper(file).Fetch<Product>().ToList();

            CollectionAssert.AreEqual(products, productsFetched);
        }

        private class ProductMapped
        {
            public string NameX { get; set; }
            public int NumberX { get; set; }
            public decimal PriceX { get; set; }

            public override bool Equals(object obj)
            {
                if (obj is not ProductMapped o) return false;
                return o.NameX == NameX && o.NumberX == NumberX && o.PriceX == PriceX;
            }

            public override int GetHashCode()
            {
                return (NameX + NumberX + PriceX).GetHashCode();
            }
        }

        [Test]
        public void SaveTrackedObjectsTest()
        {
            var excel = new ExcelMapper(@"../../../xlsx/Products.xlsx") { TrackObjects = true };

            excel.AddMapping(typeof(ProductMapped), "Name", "NameX");
            excel.AddMapping<ProductMapped>("Number", p => p.NumberX);
            excel.AddMapping<ProductMapped>("Price", p => p.PriceX);

            var products = excel.Fetch<ProductMapped>().ToList();

            products[1].PriceX += 1.0m;

            var file = @"productssavetracked.xlsx";

            excel.Save(file);

            var productsFetched = new ExcelMapper(file).Fetch<ProductMapped>().ToList();

            CollectionAssert.AreEqual(products, productsFetched);
        }

        private class GetterSetterProduct
        {
            public string Name { get; set; }
            public string RedName { get; set; }
            public DateTime? OfferEnd { get; set; }
            public string OfferEndToString { get; set; }
            public long OfferEndToLong { get; set; }

            public override bool Equals(object obj)
            {
                if (obj is not GetterSetterProduct o) return false;
                return o.Name == Name && o.OfferEnd == OfferEnd;
            }

            public override int GetHashCode()
            {
                return (Name + OfferEnd).GetHashCode();
            }
        }

        [Test]
        public void GetterSetterTest()
        {
            var excel = new ExcelMapper(@"../../../xlsx/ProductsConvert.xlsx") { TrackObjects = true };

            excel.AddMapping<GetterSetterProduct>("Name", p => p.Name);
            excel.AddMapping<GetterSetterProduct>("Name", p => p.RedName)
                .FromExcelOnly()
                .SetPropertyUsing((v, c) =>
                {
                    return c.CellStyle.FillForegroundColorColor switch
                    {
                        XSSFColor color when color.ARGBHex == "FFFF0000" => v,
                        _ => null,
                    };
                });
            excel.AddMapping<GetterSetterProduct>("OfferEnd", p => p.OfferEnd)
                .SetCellUsing((c, o) =>
                {
                    if (o == null) c.SetCellValue("NULL"); else c.SetCellValue((DateTime)o);
                })
                .SetPropertyUsing(v =>
                {
                    if ((v as string) == "NULL") return null;
                    return Convert.ChangeType(v, typeof(DateTime), CultureInfo.InvariantCulture);
                });

            // Multi "Excel to Object" unidirectional mapping
            excel.AddMapping<GetterSetterProduct>("OfferEnd", p => p.OfferEndToString)
                .FromExcelOnly()
                .SetPropertyUsing(v =>
                {
                    if ((v as string) == "NULL") return "IS_NULL";
                    var dt = (DateTime)Convert.ChangeType(v, typeof(DateTime), CultureInfo.InvariantCulture);
                    return dt.ToLongDateString();
                });

            excel.AddMapping<GetterSetterProduct>("OfferEnd", p => p.OfferEndToLong)
                .FromExcelOnly()
                .SetPropertyUsing((v, c) =>
                {
                    if ((v as string) == "NULL") return 0L;
                    var dt = (DateTime)Convert.ChangeType(v, typeof(DateTime), CultureInfo.InvariantCulture);
                    return dt.ToBinary();
                });

            var products = excel.Fetch<GetterSetterProduct>().ToList();

            Assert.Null(products[0].RedName);
            Assert.AreEqual("Halloren", products[1].RedName);

            var file = @"productsconverttracked.xlsx";

            excel.Save(file);

            var productsFetched = new ExcelMapper(file).Fetch<GetterSetterProduct>().ToList();

            CollectionAssert.AreEqual(products, productsFetched);
        }

        private class IgnoreProduct
        {
            public string Name { get; set; }
            [Ignore]
            public int Number { get; set; }
            public decimal Price { get; set; }
            public bool Offer { get; set; }
            public DateTime OfferEnd { get; set; }
            public string Value { get; set; }

            public override bool Equals(object obj)
            {
                if (obj is not IgnoreProduct o) return false;
                return o.Name == Name && o.Number == Number && o.Price == Price && o.Offer == Offer && o.OfferEnd == OfferEnd;
            }

            public override int GetHashCode()
            {
                return (Name + Number + Price + Offer + OfferEnd).GetHashCode();
            }
        }

        [Test]
        public void IgnoreTest()
        {
            var excel = new ExcelMapper(@"../../../xlsx/Products.xlsx");
            excel.Ignore<IgnoreProduct>(p => p.Price);
            excel.Ignore(typeof(IgnoreProduct), "Value");
            var products = excel.Fetch<IgnoreProduct>().ToList();

            var nudossi = products[0];
            Assert.AreEqual("Nudossi", nudossi.Name);
            Assert.AreEqual(0, nudossi.Number);
            Assert.AreEqual(0m, nudossi.Price);
            Assert.IsFalse(nudossi.Offer);
            Assert.IsNull(nudossi.Value);

            var halloren = products[1];
            Assert.IsTrue(halloren.Offer);
            Assert.AreEqual(new DateTime(2015, 12, 31), halloren.OfferEnd);
            Assert.IsNull(halloren.Value);

            var file = "productsignored.xlsx";

            new ExcelMapper().Save(file, products, "Products");

            var productsFetched = new ExcelMapper(file).Fetch<IgnoreProduct>().ToList();

            CollectionAssert.AreEqual(products, productsFetched);
        }

        private class NullableProduct
        {
            public string Name { get; set; }
            public int? Number { get; set; }
            public decimal? Price { get; set; }
            public bool? Offer { get; set; }
            public DateTime? OfferEnd { get; set; }

            public override bool Equals(object obj)
            {
                if (obj is not NullableProduct o) return false;
                return o.Name == Name && o.Number == Number && o.Price == Price && o.Offer == Offer && o.OfferEnd == OfferEnd;
            }

            public override int GetHashCode()
            {
                return (Name + Number + Price + Offer + OfferEnd).GetHashCode();
            }
        }

        [Test]
        public void NullableTest()
        {
            var workbook = WorkbookFactory.Create(@"../../../xlsx/Products.xlsx");
            var excel = new ExcelMapper(workbook);
            var products = excel.Fetch<NullableProduct>().ToList();

            var nudossi = products[0];
            Assert.AreEqual("Nudossi", nudossi.Name);
            Assert.AreEqual(60, nudossi.Number);
            Assert.AreEqual(1.99m, nudossi.Price);
            Assert.IsFalse(nudossi.Offer.Value);
            nudossi.OfferEnd = null;

            var halloren = products[1];
            Assert.IsTrue(halloren.Offer.Value);
            Assert.AreEqual(new DateTime(2015, 12, 31), halloren.OfferEnd);
            halloren.Number = null;
            halloren.Offer = null;

            var file = "productsnullable.xlsx";

            new ExcelMapper().Save(file, products, "Products");

            var productsFetched = new ExcelMapper(file).Fetch<NullableProduct>().ToList();

            CollectionAssert.AreEqual(products, productsFetched);
        }

        [Test]
        public void NullableDynamicTest()
        {
            var workbook = WorkbookFactory.Create(@"../../../xlsx/Products.xlsx");
            var excel = new ExcelMapper(workbook) { SkipBlankCells = false };
            var products = excel.Fetch().ToList();
            var nudossi = products[0];
            Assert.AreEqual("Nudossi", nudossi.Name);
            Assert.AreEqual(60, nudossi.Number);
            Assert.AreEqual(1.99m, nudossi.Price);
            Assert.IsFalse(nudossi.Offer);
            Assert.IsNotNull(nudossi.OfferEnd);
            nudossi.OfferEnd = null; //set to null to test it

            var halloren = products[1];
            Assert.IsTrue(halloren.Offer);
            Assert.AreEqual(new DateTime(2015, 12, 31), halloren.OfferEnd);
            Assert.IsNotNull(halloren.Number);
            halloren.Number = null; //set to null to test it
            Assert.IsNotNull(halloren.Offer);
            halloren.Offer = null; //set to null to test it

            var file = "productsnullabledynamic.xlsx";

            new ExcelMapper().Save(file, products, "Products");

            var productsFetched = new ExcelMapper(file) { SkipBlankCells = false }.Fetch(0, (colnum, value) =>
            {
                //convert an empty string to null
                if (value is string && value.ToString().Length == 0 && new string[] { "OfferEnd", "Number", "Offer" }.Contains(colnum))
                {
                    return null;
                }
                return value;
            }).ToList();

            Assert.IsNull(productsFetched[0].OfferEnd);
            Assert.IsNull(productsFetched[1].Number);
            Assert.IsNull(productsFetched[1].Offer);
        }

        private class DataFormatProduct
        {
            [DataFormat(0xf)]
            public DateTime Date { get; set; }

            [DataFormat("0%")]
            public decimal Number { get; set; }
        }

        [Test]
        public void DataFormatTest()
        {
            var p = new DataFormatProduct { Date = new DateTime(2015, 12, 31), Number = 0.47m };
            var file = "productsdataformat.xlsx";
            new ExcelMapper().Save(file, new[] { p });
            var pfs = new ExcelMapper(file).Fetch<DataFormatProduct>().ToList();

            Assert.AreEqual(1, pfs.Count);
            var pf = pfs[0];
            Assert.AreEqual(p.Date, pf.Date);
            Assert.AreEqual(p.Number, pf.Number);
        }

        private class DataItem
        {
            [Ignore]
            public string Bql { get; set; }

            [Ignore]
            public int Id { get; set; }

            [Column(2)]
            public string OriginalBql { get; set; }

            [Column(1)]
            public string Title { get; set; }

            [Column(3)]
            public string TranslatedBql { get; set; }

            public override bool Equals(object obj)
            {
                if (obj is not DataItem o) return false;
                return o.Bql == Bql && o.Id == Id && o.OriginalBql == OriginalBql && o.Title == Title && o.TranslatedBql == TranslatedBql;
            }

            public override int GetHashCode()
            {
                return (Bql + Id + OriginalBql + Title + TranslatedBql).GetHashCode();
            }
        }

        [Test]
        public void ColumnTest()
        {
            var excel = new ExcelMapper(@"../../../xlsx/DataItems.xlsx") { HeaderRow = false };
            var items = excel.Fetch<DataItem>().ToList();

            var trackedFile = "dataitemstracked.xlsx";
            excel.Save(trackedFile, "DataItems");
            var itemsTracked = excel.Fetch<DataItem>(trackedFile, "DataItems").ToList();
            CollectionAssert.AreEqual(items, itemsTracked);

            var saveFile = "dataitemssave.xlsx";
            new ExcelMapper().Save(saveFile, items, "DataItems");
            var itemsSaved = new ExcelMapper().Fetch<DataItem>(saveFile, "DataItems").ToList();
            CollectionAssert.AreEqual(items, itemsSaved);
        }

        [Test]
        public void ColumnTestUsingFetchWithType()
        {
            var excel = new ExcelMapper(@"../../../xlsx/DataItems.xlsx") { HeaderRow = false };
            var items = excel.Fetch(typeof(DataItem)).OfType<DataItem>().ToList();

            var trackedFile = "dataitemstracked1.xlsx";
            excel.Save(trackedFile, "DataItems");
            var itemsTracked = excel.Fetch(trackedFile, typeof(DataItem), "DataItems").OfType<DataItem>().ToList();
            CollectionAssert.AreEqual(items, itemsTracked);

            var saveFile = "dataitemssave1.xlsx";
            new ExcelMapper().Save(saveFile, items, "DataItems");
            var itemsSaved = new ExcelMapper().Fetch(saveFile, typeof(DataItem), "DataItems").OfType<DataItem>().ToList();
            CollectionAssert.AreEqual(items, itemsSaved);
        }

        [Test]
        public void FetchMinMaxTest()
        {
            var products = new ExcelMapper(@"../../../xlsx/ProductsMinMaxRow.xlsx")
            {
                HeaderRowNumber = 2,
                MinRowNumber = 6,
                MaxRowNumber = 9,
            }.Fetch<Product>().ToList();
            CollectionAssert.AreEqual(new List<Product>
            {
                new Product { Name = "Nudossi", NumberInStock = 60, Price = 1.99m, Value = "C7*D7" },
                new Product { Name = "Halloren", NumberInStock = 33, Price = 2.99m, Value = "C8*D8" },
                new Product { Name = "Filinchen", NumberInStock = 100, Price = 0.99m, Value = "C10*D10" },
            }, products);
        }

        [Test]
        public void SaveMinMaxTest()
        {
            var products = new List<Product>
            {
                new Product { Name = "Nudossi", NumberInStock = 60, Price = 1.99m, Value = "C2*D2" },
                new Product { Name = "Halloren", NumberInStock = 33, Price = 2.99m, Value = "C3*D3" },
                new Product { Name = "Filinchen", NumberInStock = 100, Price = 0.99m, Value = "C4*D4" },
            };

            var file = "productsminmaxsave.xlsx";

            new ExcelMapper
            {
                HeaderRowNumber = 1,
                MinRowNumber = 3
            }.Save(file, products, "Products");

            var productsFetched = new ExcelMapper(file)
            {
                HeaderRowNumber = 1,
                MinRowNumber = 3
            }.Fetch<Product>().ToList();

            CollectionAssert.AreEqual(products, productsFetched);
        }

        [Test]
        public void FormulaResultAttributeTest()
        {
            var products = new ExcelMapper(@"../../../xlsx/ProductsAsString.xlsx").Fetch<ProductValueString>().ToList();
            CollectionAssert.AreEqual(new List<string> { "119.4", "98.67", "99" }, products.Select(p => p.ValueAsString).ToList());
        }

        private class ProductFormulaMapped
        {
            public decimal Result { get; set; }
            public string Formula { get; set; }
            public string ResultString { get; set; }

            public override bool Equals(object obj)
            {
                if (obj is not ProductFormulaMapped o) return false;
                return o.Result == Result && o.Formula == Formula && o.ResultString == ResultString;
            }

            public override int GetHashCode()
            {
                return (Formula + ResultString + Result).GetHashCode();
            }
        }

        [Test]
        public void FormulaResultMappedTest()
        {
            var excel = new ExcelMapper(@"../../../xlsx/ProductsAsString.xlsx");

            excel.AddMapping<ProductFormulaMapped>("Value", p => p.Result);
            excel.AddMapping<ProductFormulaMapped>("ValueDefaultAsFormula", p => p.Formula);
            excel.AddMapping<ProductFormulaMapped>("ValueAsString", p => p.ResultString).AsFormulaResult();

            var products = excel.Fetch<ProductFormulaMapped>().ToList();
            var expectedProducts = new List<ProductFormulaMapped>
            {
                new ProductFormulaMapped { Result = 119.4m, Formula = "C2*D2", ResultString = "119.4" },
                new ProductFormulaMapped { Result = 98.67m, Formula = "C3*D3", ResultString = "98.67" },
                new ProductFormulaMapped { Result = 99m, Formula = "C5*D5", ResultString = "99" },
            };

            Assert.AreEqual(expectedProducts[0], products[0]);

            CollectionAssert.AreEqual(expectedProducts, products);
        }

        [Test]
        public void TestExcelMapperConvertException()
        {
            ExcelMapperConvertException ex = new("cellvalue", typeof(string), 12, 34);

            // Sanity check: Make sure custom properties are set before serialization
            Assert.AreEqual(12, ex.Line);
            Assert.AreEqual(34, ex.Column);

            // Round-trip the exception: Serialize and de-serialize
            var serializer = new DataContractJsonSerializer(typeof(ExcelMapperConvertException));
            using (var ms = new MemoryStream())
            {
                // "Save" object state
                serializer.WriteObject(ms, ex);

                // Re-use the same stream for de-serialization
                ms.Seek(0, 0);

                // Replace the original exception with de-serialized one
                ex = (ExcelMapperConvertException)serializer.ReadObject(ms);
            }

            // Make sure custom properties are preserved after serialization
            Assert.AreEqual(12, ex.Line);
            Assert.AreEqual(34, ex.Column);

            Assert.Throws<ArgumentNullException>(() => ex.GetObjectData(null, new System.Runtime.Serialization.StreamingContext()));
        }

        private class ProductIndex
        {
            [Column(1)]
            public string Price { get; set; }
            [Column(Letter = "C")]
            public string Name { get; set; }
            [Column(4)]
            public string Number { get; set; }

            public override bool Equals(object obj)
            {
                if (obj is not ProductIndex o) return false;
                return o.Name == Name && o.Number == Number && o.Price == Price;
            }

            public override int GetHashCode()
            {
                return (Name + Number + Price).GetHashCode();
            }
        }

        [Test]
        public void FetchIndexTest()
        {
            var products = new ExcelMapper(@"../../../xlsx/Products.xlsx").Fetch<ProductIndex>().ToList();
            CollectionAssert.AreEqual(new List<ProductIndex>
            {
                new ProductIndex { Price = "Nudossi", Name = "60", Number = "1.99" },
                new ProductIndex { Price = "Halloren", Name = "33", Number = "2.99" },
                new ProductIndex { Price = "Filinchen", Name = "100", Number = "0.99" },
            }, products);
        }

        private class ProductDoubleMap
        {
            [Column(1)]
            public string Price { get; set; }
            public string Name { get; set; }

            [Column("Number")]
            public string OtherNumber { get; set; }
            public string Number { get; set; }

            public override bool Equals(object obj)
            {
                if (obj is not ProductDoubleMap o) return false;
                return o.Name == Name && o.Number == Number && o.Price == Price && o.OtherNumber == OtherNumber;
            }

            public override int GetHashCode()
            {
                return (Name + Number + Price + OtherNumber).GetHashCode();
            }
        }

        [Test]
        public void FetchDoubleMap()
        {
            // https://github.com/mganss/ExcelMapper/issues/50
            var products = new ExcelMapper(@"../../../xlsx/Products.xlsx").Fetch<ProductDoubleMap>().ToList();
            CollectionAssert.AreEqual(new List<ProductDoubleMap>
            {
                new ProductDoubleMap { Price = "Nudossi", OtherNumber = "60" },
                new ProductDoubleMap { Price = "Halloren", OtherNumber = "33" },
                new ProductDoubleMap { Price = "Filinchen", OtherNumber = "100" },
            }, products);
        }

        static void AssertProducts(List<Product> products)
        {
            CollectionAssert.AreEqual(new List<Product>
            {
                new Product { Name = "Nudossi", NumberInStock = 60, Price = 1.99m, Value = "C2*D2" },
                new Product { Name = "Halloren", NumberInStock = 33, Price = 2.99m, Value = "C3*D3" },
                new Product { Name = "Filinchen", NumberInStock = 100, Price = 0.99m, Value = "C5*D5" },
            }, products);
        }

        [Test]
        public async Task FetchAsyncTest()
        {
            var path = @"../../../xlsx/Products.xlsx";

            var products = (await new ExcelMapper().FetchAsync<Product>(path)).ToList();
            AssertProducts(products);

            products = (await new ExcelMapper().FetchAsync<Product>(path, "Tabelle1")).ToList();
            AssertProducts(products);

            products = (await new ExcelMapper().FetchAsync(path, typeof(Product), "Tabelle1")).OfType<Product>().ToList();
            AssertProducts(products);

            products = (await new ExcelMapper().FetchAsync(path, typeof(Product))).OfType<Product>().ToList();
            AssertProducts(products);

            products = (await new ExcelMapper().FetchAsync<Product>(File.OpenRead(path))).ToList();
            AssertProducts(products);

            products = (await new ExcelMapper().FetchAsync<Product>(File.OpenRead(path), "Tabelle1")).ToList();
            AssertProducts(products);

            products = (await new ExcelMapper().FetchAsync(File.OpenRead(path), typeof(Product), "Tabelle1")).OfType<Product>().ToList();
            AssertProducts(products);

            products = (await new ExcelMapper().FetchAsync(File.OpenRead(path), typeof(Product))).OfType<Product>().ToList();
            AssertProducts(products);
        }

        [Test]
        public async Task SaveAsyncTest()
        {
            var products = new List<Product>
            {
                new Product { Name = "Nudossi", NumberInStock = 60, Price = 1.99m, Value = "C2*D2" },
                new Product { Name = "Halloren", NumberInStock = 33, Price = 2.99m, Value = "C3*D3" },
                new Product { Name = "Filinchen", NumberInStock = 100, Price = 0.99m, Value = "C4*D4" },
            };

            var file = "productssave.xlsx";

            await new ExcelMapper().SaveAsync(file, products, "Products");
            var productsFetched = new ExcelMapper(file).Fetch<Product>().ToList();
            CollectionAssert.AreEqual(products, productsFetched);

            await new ExcelMapper().SaveAsync(file, products);
            productsFetched = new ExcelMapper(file).Fetch<Product>().ToList();
            CollectionAssert.AreEqual(products, productsFetched);

            var fs = File.OpenWrite(file);
            await new ExcelMapper().SaveAsync(fs, products, "Products");
            fs.Close();
            productsFetched = new ExcelMapper(file).Fetch<Product>().ToList();
            CollectionAssert.AreEqual(products, productsFetched);

            fs = File.OpenWrite(file);
            await new ExcelMapper().SaveAsync(fs, products);
            fs.Close();
            productsFetched = new ExcelMapper(file).Fetch<Product>().ToList();
            CollectionAssert.AreEqual(products, productsFetched);

            var path = @"../../../xlsx/Products.xlsx";

            var mapper = new ExcelMapper() { TrackObjects = true };
            var tracked = (await mapper.FetchAsync<Product>(path)).ToList();
            await mapper.SaveAsync(file, "Tabelle1");
            productsFetched = new ExcelMapper(file).Fetch<Product>().ToList();
            CollectionAssert.AreEqual(tracked, productsFetched);

            mapper = new ExcelMapper() { TrackObjects = true };
            tracked = (await mapper.FetchAsync<Product>(path)).ToList();
            await mapper.SaveAsync(file);
            productsFetched = new ExcelMapper(file).Fetch<Product>().ToList();
            CollectionAssert.AreEqual(tracked, productsFetched);

            mapper = new ExcelMapper() { TrackObjects = true };
            tracked = (await mapper.FetchAsync<Product>(path)).ToList();
            fs = File.OpenWrite(file);
            await mapper.SaveAsync(fs, "Tabelle1");
            fs.Close();
            productsFetched = new ExcelMapper(file).Fetch<Product>().ToList();
            CollectionAssert.AreEqual(tracked, productsFetched);

            mapper = new ExcelMapper() { TrackObjects = true };
            tracked = (await mapper.FetchAsync<Product>(path)).ToList();
            fs = File.OpenWrite(file);
            await mapper.SaveAsync(fs);
            fs.Close();
            productsFetched = new ExcelMapper(file).Fetch<Product>().ToList();
            CollectionAssert.AreEqual(tracked, productsFetched);
        }

        private class Course
        {
            public string Mode { get; set; }
            public string AdEnrollSchedId { get; set; }
            public string SyStudentId { get; set; }
            public string StuNum { get; set; }
            public string AdEnrollId { get; set; }
            public string SyCampusId { get; set; }
            public string AdCourseId { get; set; }
            public string CourseCode { get; set; }
            public string AdClassSchedId { get; set; }
            public string SectionCode { get; set; }
            public string CourseStartDate { get; set; }
            public string CourseEndDate { get; set; }
            public string AdTermId { get; set; }
            public string TermCode { get; set; }
            public string State { get; set; }
        }

        [Test]
        public void DateTest()
        {
            // see https://github.com/mganss/ExcelMapper/issues/51
            var mapper = new ExcelMapper(@"../../../xlsx/DateTest.xlsx") { HeaderRow = true };

            var courses = mapper.Fetch<Course>().ToList();

            Assert.AreEqual("00:00.0", courses.First().CourseStartDate);
        }
        private class ProductJson
        {
            [Json]
            public Product Product { get; set; }
        }

        private class ProductJsonMapped
        {
            [Json]
            public Product Product { get; set; }
        }

        [Test]
        public void JsonTest()
        {
            var products = new ExcelMapper(@"../../../xlsx/ProductsJson.xlsx").Fetch<ProductJson>().ToList();

            CollectionAssert.AreEqual(new List<Product>
            {
                new Product { Name = "Nudossi", NumberInStock = 60, Price = 1.99m, Value = "C2*D2" },
                new Product { Name = "Halloren", NumberInStock = 33, Price = 2.99m, Value = "C3*D3" },
                new Product { Name = "Filinchen", NumberInStock = 100, Price = 0.99m, Value = "C4*D4" },
            }, products.Select(p => p.Product));

            var file = "productsjsonsave.xlsx";

            new ExcelMapper().Save(file, products, "Products");

            var productsFetched = new ExcelMapper(file).Fetch<ProductJson>().ToList();

            CollectionAssert.AreEqual(products.Select(p => p.Product), productsFetched.Select(p => p.Product));
        }

        [Test]
        public void JsonMappedTest()
        {
            var excel = new ExcelMapper(@"../../../xlsx/ProductsJson.xlsx");

            excel.AddMapping<ProductJsonMapped>("Product", p => p.Product).AsJson();

            var products = excel.Fetch<ProductJsonMapped>().ToList();

            CollectionAssert.AreEqual(new List<Product>
            {
                new Product { Name = "Nudossi", NumberInStock = 60, Price = 1.99m, Value = "C2*D2" },
                new Product { Name = "Halloren", NumberInStock = 33, Price = 2.99m, Value = "C3*D3" },
                new Product { Name = "Filinchen", NumberInStock = 100, Price = 0.99m, Value = "C4*D4" },
            }, products.Select(p => p.Product));
        }

        private class ProductJsonList
        {
            [Json]
            public List<Product> Products { get; set; }
        }

        [Test]
        public void JsonListTest()
        {
            var products = new ExcelMapper(@"../../../xlsx/ProductsJsonList.xlsx").Fetch<ProductJsonList>().ToList();

            CollectionAssert.AreEqual(new List<Product>
            {
                new Product { Name = "Nudossi", NumberInStock = 60, Price = 1.99m, Value = "C2*D2" },
                new Product { Name = "Halloren", NumberInStock = 33, Price = 2.99m, Value = "C3*D3" },
                new Product { Name = "Filinchen", NumberInStock = 100, Price = 0.99m, Value = "C4*D4" },
            }, products.First().Products);

            var file = "productsjsonlistsave.xlsx";

            new ExcelMapper().Save(file, products, "Products");

            var productsFetched = new ExcelMapper(file).Fetch<ProductJsonList>().ToList();

            CollectionAssert.AreEqual(products.First().Products, productsFetched.First().Products);
        }

        private class Sample2
        {
            [Column(37)]
            public DateTime? CaptDate { get; set; }

            [Column(38)]
            public string Identifier { get; set; }

            [Column(39)]
            public int Value { get; set; }

            [Column(40)]
            public int Disposed { get; set; }
        }

        [Test]
        public void Number74Test()
        {
            const int N = 3;

            for (var i = 0; i < N; i++)
            {
                using var f2 = File.OpenRead(@"../../../xlsx/SampleExcel.xlsx");
                var s2 = FetchWaterCaptationComplementsAsync(f2).Result;
            }
        }

        private static async Task<List<Sample2>> FetchWaterCaptationComplementsAsync(Stream file)
        {
            var excelMapper = new ExcelMapper { HeaderRow = false };
            var samples = (await excelMapper.FetchAsync<Sample2>(file)).ToList();
            return samples;
        }

        [Test]
        public void NormalizeTest()
        {
            var excel = new ExcelMapper(@"../../../xlsx/Products.xlsx");

            excel.NormalizeUsing(n => n + "X");

            var products = excel.Fetch<ProductMapped>().ToList();

            CollectionAssert.AreEqual(new List<ProductMapped>
            {
                new ProductMapped { NameX = "Nudossi", NumberX = 60, PriceX = 1.99m },
                new ProductMapped { NameX = "Halloren", NumberX = 33, PriceX = 2.99m },
                new ProductMapped { NameX = "Filinchen", NumberX = 100, PriceX = 0.99m },
            }, products);
        }

        [Test]
        public void NormalizeTypeTest()
        {
            var excel = new ExcelMapper(@"../../../xlsx/Products.xlsx");

            excel.AddMapping<ProductMapped>("NameY", p => p.NameX);
            excel.AddMapping<ProductMapped>("NumberY", p => p.NumberX);
            excel.AddMapping<ProductMapped>("PriceY", p => p.PriceX);

            excel.NormalizeUsing<ProductMapped>(n => n + "Y");

            var products = excel.Fetch<ProductMapped>().ToList();

            CollectionAssert.AreEqual(new List<ProductMapped>
            {
                new ProductMapped { NameX = "Nudossi", NumberX = 60, PriceX = 1.99m },
                new ProductMapped { NameX = "Halloren", NumberX = 33, PriceX = 2.99m },
                new ProductMapped { NameX = "Filinchen", NumberX = 100, PriceX = 0.99m },
            }, products);
        }

        [Test]
        public void NormalizeType2Test()
        {
            var excel = new ExcelMapper(@"../../../xlsx/Products.xlsx");

            excel.AddMapping<ProductMapped>("NameY", p => p.NameX);
            excel.AddMapping<ProductMapped>("NumberY", p => p.NumberX);
            excel.AddMapping<ProductMapped>("PriceY", p => p.PriceX);

            excel.NormalizeUsing(typeof(ProductMapped), n => n + "Y");

            var products = excel.Fetch<ProductMapped>().ToList();

            CollectionAssert.AreEqual(new List<ProductMapped>
            {
                new ProductMapped { NameX = "Nudossi", NumberX = 60, PriceX = 1.99m },
                new ProductMapped { NameX = "Halloren", NumberX = 33, PriceX = 2.99m },
                new ProductMapped { NameX = "Filinchen", NumberX = 100, PriceX = 0.99m },
            }, products);
        }

        [Test]
        public void LetterConversionTest()
        {
            Assert.AreEqual(1, ExcelMapper.LetterToIndex("A"));
            Assert.AreEqual(1, ExcelMapper.LetterToIndex("$A"));
            Assert.AreEqual(649, ExcelMapper.LetterToIndex("XY"));
            Assert.AreEqual(649, ExcelMapper.LetterToIndex("$XY"));
            Assert.AreEqual(649, ExcelMapper.LetterToIndex("xy"));
            Assert.AreEqual("AB", ExcelMapper.IndexToLetter(28));
            Assert.AreEqual("A", ExcelMapper.IndexToLetter(1));
            Assert.AreEqual("XY", ExcelMapper.IndexToLetter(649));

            Assert.Throws<ArgumentException>(() => ExcelMapper.LetterToIndex(null));
            Assert.Throws<ArgumentException>(() => ExcelMapper.LetterToIndex("???"));
            Assert.Throws<ArgumentException>(() => ExcelMapper.LetterToIndex("A$"));
            Assert.Throws<ArgumentException>(() => ExcelMapper.IndexToLetter(-1));
        }

        [Test]
        public void ColumnSkipTest()
        {
            // see https://github.com/mganss/ExcelMapper/issues/90
            var products = new ExcelMapper(@"../../../xlsx/ProductsExceptionEmpty.xlsx") { SkipBlankCells = false }.Fetch().ToList();
            Assert.AreEqual(1, products.Count);
            var p = products[0];
            Assert.IsEmpty(p.Price);
            Assert.AreEqual("Nudossi", p.Name);
            Assert.AreEqual(60, p.Number);
        }

        class NullProduct
        {
            public string Number { get; set; }
            public string Color { get; set; }
        }

        [Test]
        public void NullTest()
        {
            // see https://github.com/mganss/ExcelMapper/issues/96
            var products = new ExcelMapper(@"../../../xlsx/null_test.xlsx").Fetch<NullProduct>().ToList();
            Assert.AreEqual(20, products.Count);
        }

        private record ProductRecord
        {
            public string Name { get; }
            [Column("Number")]
            public int NumberInStock { get; }
            public decimal Price { get; }
            public string Value { get; }

            public ProductRecord(string name, int numberinstock, decimal price, string value) => (Name, NumberInStock, Price, Value) = (name, numberinstock, price, value);
        }

        [Test]
        public void RecordFetchTest()
        {
            var products = new ExcelMapper(@"../../../xlsx/Products.xlsx").Fetch<ProductRecord>().ToList();
            CollectionAssert.AreEqual(new List<ProductRecord>
            {
                new ProductRecord("Nudossi", 60, 1.99m, "C2*D2"),
                new ProductRecord("Halloren", 33, 2.99m, "C3*D3"),
                new ProductRecord("Filinchen", 100, 0.99m, "C5*D5"),
            }, products);
        }

        [Test]
        public void SaveFetchedRecordTest()
        {
            var excel = new ExcelMapper(@"../../../xlsx/Products.xlsx");
            var products = excel.Fetch<ProductRecord>().ToList();

            var file = @"productssavefetchedrecord.xlsx";

            excel.Save(file, products);

            var productsFetched = new ExcelMapper(file).Fetch<ProductRecord>().ToList();

            CollectionAssert.AreEqual(products, productsFetched);
        }

        private record ProductRecordNoHeaderManual
        {
            [Ignore]
            public bool Offer { get; set; }
            public string NameX { get; }
            public int NumberInStockX { get; }
            public string Unmapped { get; set; }
            public decimal PriceX { get; }

            public ProductRecordNoHeaderManual(bool offer, string namex, int numberinstockx, string unmapped, decimal pricex) =>
                (Offer, NameX, NumberInStockX, Unmapped, PriceX) = (offer, namex, numberinstockx, unmapped, pricex);
        }

        [Test]
        public void FetchRecordNoHeaderManualTest()
        {
            var excel = new ExcelMapper(@"../../../xlsx/ProductsNoHeader.xlsx") { HeaderRow = false };

            excel.AddMapping<ProductRecordNoHeaderManual>(1, p => p.NameX);
            excel.AddMapping<ProductRecordNoHeaderManual>(ExcelMapper.LetterToIndex("C"), p => p.NumberInStockX);
            excel.AddMapping(typeof(ProductRecordNoHeaderManual), 4, "PriceX");

            var products = excel.Fetch<ProductRecordNoHeaderManual>("Products").ToList();

            CollectionAssert.AreEqual(new List<ProductRecordNoHeaderManual>
            {
                new ProductRecordNoHeaderManual(false, "Nudossi", 60, null, 1.99m),
                new ProductRecordNoHeaderManual(false, "Halloren", 33, null, 2.99m),
                new ProductRecordNoHeaderManual(false, "Filinchen", 100, null, 0.99m),
            }, products);
        }

        private record ProductPosRecord(int Number, string Name, decimal Price, string Value);

        [Test]
        public void PosRecordFetchTest()
        {
            var products = new ExcelMapper(@"../../../xlsx/Products.xlsx").Fetch<ProductPosRecord>().ToList();
            CollectionAssert.AreEqual(new List<ProductPosRecord>
            {
                new ProductPosRecord(60, "Nudossi", 1.99m, "C2*D2"),
                new ProductPosRecord(33, "Halloren", 2.99m, "C3*D3"),
                new ProductPosRecord(100, "Filinchen", 0.99m, "C5*D5"),
            }, products);
        }

        private record CustomProduct
        {
            [Column(Letter = "B")]
            public string Name { get; set; }
            [Column("Number", Letter = "C")]
            public int NumberInStock { get; set; }
            [Column(Letter = "D")]
            public decimal Price { get; set; }
        }

        [Test]
        public void SaveMissingHeadersTest()
        {
            var products = new List<CustomProduct>
            {
                new CustomProduct { Name = "Nudossi", NumberInStock = 60, Price = 1.99m },
                new CustomProduct { Name = "Halloren", NumberInStock = 33, Price = 2.99m },
                new CustomProduct { Name = "Filinchen", NumberInStock = 100, Price = 0.99m },
            };

            var excelMapper = new ExcelMapper(@"../../../xlsx/ProductsMissingHeaders.xlsx")
            {
                HeaderRowNumber = 2,
                MinRowNumber = 3,
                CreateMissingHeaders = true
            };

            var file = "ProductsMissingHeaders.xlsx";

            excelMapper.Save(file, products, "PROD");

            var productsFetched = new ExcelMapper(file)
            {
                HeaderRowNumber = 2,
                MinRowNumber = 3,
                CreateMissingHeaders = true
            }.Fetch<CustomProduct>().ToList();

            CollectionAssert.AreEqual(products, productsFetched);
        }

        [Test]
        public void LongRowsTest()
        {
            var rows = new ExcelMapper(@"../../../xlsx/JaggedRows.xlsx") { HeaderRow = false, SkipBlankCells = false }.Fetch().ToList();

            Assert.AreEqual(2, rows.Count);
            Assert.AreEqual(13, ((IDictionary<string, object>)rows[0]).Count);
            Assert.AreEqual("TestL", rows[1].L);
        }

        private class OfferDetails
        {
            [Column("Offer")]
            public bool IsOffer { get; set; }
            [Column("OfferEnd")]
            public DateTime End { get; set; }

            public override bool Equals(object obj)
            {
                if (obj is not OfferDetails o) return false;
                return o.IsOffer == IsOffer && o.End == End;
            }

            public override int GetHashCode()
            {
                return HashCode.Combine(IsOffer, End);
            }

            public OfferDetails(bool isOffer, DateTime end)
            {
                IsOffer = isOffer;
                End = end;
            }

            public OfferDetails() { }
        }

        private class NestedProduct
        {
            public string Name { get; set; }
            public int Number { get; set; }
            public decimal Price { get; set; }
            public OfferDetails Offer { get; set; } = new();

            public override bool Equals(object obj)
            {
                if (obj is not NestedProduct o) return false;
                return o.Name == Name && o.Number == Number && o.Price == Price && o.Offer.Equals(Offer);
            }

            public override int GetHashCode()
            {
                return HashCode.Combine(Name, Number, Price, Offer);
            }

            public NestedProduct(string name, int number, decimal price, OfferDetails offer)
            {
                Name = name;
                Number = number;
                Price = price;
                Offer = offer;
            }

            public NestedProduct() { }
        }

        [Test]
        public void NestedTest()
        {
            var excel = new ExcelMapper(@"../../../xlsx/Products.xlsx");
            var products = excel.Fetch<NestedProduct>().ToList();

            Assert.AreEqual(3, products.Count);
            Assert.AreEqual(false, products[0].Offer.IsOffer);
            Assert.AreEqual(new DateTime(1970, 1, 1), products[0].Offer.End);
            Assert.AreEqual(true, products[1].Offer.IsOffer);
            Assert.AreEqual(new DateTime(2015, 12, 31), products[1].Offer.End);
        }

        [Test]
        public void IgnoreNestedTest()
        {
            var excel = new ExcelMapper(@"../../../xlsx/Products.xlsx") { IgnoreNestedTypes = true };
            var products = excel.Fetch<NestedProduct>().ToList();

            Assert.AreEqual(3, products.Count);

            CollectionAssert.AreEqual(new List<NestedProduct>
            {
                new NestedProduct("Nudossi", 60, 1.99m, new OfferDetails()),
                new NestedProduct("Halloren", 33, 2.99m, new OfferDetails()),
                new NestedProduct("Filinchen", 100, 0.99m, new OfferDetails()),
            }, products);
        }

        [Test]
        public void NestedSaveTest()
        {
            var products = new List<NestedProduct>
            {
                new NestedProduct("Nudossi", 60, 1.99m, new OfferDetails(false, new DateTime(1970, 01, 01))),
                new NestedProduct("Halloren", 33, 2.99m, new OfferDetails(true, new DateTime(2015, 12, 31))),
                new NestedProduct("Filinchen", 100, 0.99m, new OfferDetails(false, new DateTime(1970, 01, 01))),
            };

            var file = "nestedsave.xlsx";

            new ExcelMapper().Save(file, products, "Products");

            var excel = new ExcelMapper(file);
            var productsFetched = excel.Fetch<NestedProduct>().ToList();

            CollectionAssert.AreEqual(products, productsFetched);

            productsFetched[0].Name = "Nudossi2";
            productsFetched[0].Offer.End = new DateTime(2021, 4, 21);

            excel.Save(file);

            var productsFetched2 = excel.Fetch<NestedProduct>().ToList();

            CollectionAssert.AreEqual(productsFetched, productsFetched2);
        }

        [Test]
        public void IgnoreNestedSaveTest()
        {
            var products = new List<NestedProduct>
            {
                new NestedProduct("Nudossi", 60, 1.99m, new OfferDetails(false, new DateTime(1970, 01, 01))),
                new NestedProduct("Halloren", 33, 2.99m, new OfferDetails(true, new DateTime(2015, 12, 31))),
                new NestedProduct("Filinchen", 100, 0.99m, new OfferDetails(false, new DateTime(1970, 01, 01))),
            };

            var file = "ignorenestedsave.xlsx";

            new ExcelMapper() { IgnoreNestedTypes = true }.Save(file, products, "Products");

            var excel = new ExcelMapper(file) { IgnoreNestedTypes = true };
            var productsFetched = excel.Fetch<NestedProduct>().ToList();

            CollectionAssert.AreEqual(new List<NestedProduct>
            {
                new NestedProduct("Nudossi", 60, 1.99m, new OfferDetails()),
                new NestedProduct("Halloren", 33, 2.99m, new OfferDetails()),
                new NestedProduct("Filinchen", 100, 0.99m, new OfferDetails()),
            }, productsFetched);
        }

        private record OfferDetailsRecord(bool Offer, DateTime OfferEnd);
        private record NestedRecord(string Name, int Number, decimal Price, decimal Value, OfferDetailsRecord OfferDetails);

        [Test]
        public void NestedRecordTest()
        {
            var products = new ExcelMapper(@"../../../xlsx/Products.xlsx").Fetch<NestedRecord>().ToList();

            CollectionAssert.AreEqual(new List<NestedRecord>
            {
                new NestedRecord("Nudossi", 60, 1.99m, 119.40m, new OfferDetailsRecord(false, new DateTime(1970, 01, 01))),
                new NestedRecord("Halloren", 33, 2.99m, 98.67m, new OfferDetailsRecord(true, new DateTime(2015, 12, 31))),
                new NestedRecord("Filinchen", 100, 0.99m, 99.00m, new OfferDetailsRecord(false, new DateTime(1970, 01, 01))),
            }, products);
        }

        [Test]
        public void NestedRecordSaveTest()
        {
            var products = new List<NestedRecord>
            {
                new NestedRecord("Nudossi", 60, 1.99m, 119.40m, new OfferDetailsRecord(false, new DateTime(1970, 01, 01))),
                new NestedRecord("Halloren", 33, 2.99m, 98.67m, new OfferDetailsRecord(true, new DateTime(2015, 12, 31))),
                new NestedRecord("Filinchen", 100, 0.99m, 99.00m, new OfferDetailsRecord(false, new DateTime(1970, 01, 01))),
            };

            var file = "nestedrecordssave.xlsx";

            new ExcelMapper().Save(file, products, "Products");

            var excel = new ExcelMapper(file);
            var productsFetched = excel.Fetch<NestedRecord>().ToList();

            CollectionAssert.AreEqual(products, productsFetched);
        }

        [Test]
        public void NestedRecordSaveMissingHeadersTest()
        {
            var products = new List<NestedRecord>
            {
                new NestedRecord("Nudossi", 60, 1.99m, 119.40m, new OfferDetailsRecord(false, new DateTime(1970, 01, 01))),
                new NestedRecord("Halloren", 33, 2.99m, 98.67m, new OfferDetailsRecord(true, new DateTime(2015, 12, 31))),
                new NestedRecord("Filinchen", 100, 0.99m, 99.00m, new OfferDetailsRecord(false, new DateTime(1970, 01, 01))),
            };

            var excelMapper = new ExcelMapper(@"../../../xlsx/ProductsMissingHeaders.xlsx")
            {
                CreateMissingHeaders = true,
                HeaderRowNumber = 2,
                MinRowNumber = 3,
            };

            var file = "nestedrecordssavemissingheaders.xlsx";

            excelMapper.Save(file, products, "PROD");

            var productsFetched = new ExcelMapper(file)
            {
                HeaderRowNumber = 2,
                MinRowNumber = 3
            }.Fetch<NestedRecord>().ToList();

            CollectionAssert.AreEqual(products, productsFetched);
        }

        private class NestedOfferMapped
        {
            public bool O { get; set; }
            public DateTime E { get; set; }
            public NestedProductMapped Cycle { get; set; }

            public override bool Equals(object obj)
            {
                if (obj is not NestedOfferMapped o) return false;
                return o.O == O && o.E == E;
            }

            public override int GetHashCode()
            {
                return HashCode.Combine(O, E);
            }
        }

        private class NestedProductMapped
        {
            public string N { get; set; }
            public int Num { get; set; }
            public decimal P { get; set; }
            public NestedOfferMapped O { get; set; }

            public override bool Equals(object obj)
            {
                if (obj is not NestedProductMapped o) return false;
                return o.N == N && o.Num == Num && o.P == P && o.O.Equals(O);
            }

            public override int GetHashCode()
            {
                return HashCode.Combine(N, Num, P, O);
            }
        }

        [Test]
        public void NestedProductMappedTest()
        {
            var excel = new ExcelMapper(@"../../../xlsx/Products.xlsx");

            excel.AddMapping<NestedProductMapped>("Name", p => p.N);
            excel.AddMapping<NestedProductMapped>("Number", p => p.Num);
            excel.AddMapping<NestedProductMapped>("Price", p => p.P);
            excel.AddMapping<NestedOfferMapped>("Offer", p => p.O);
            excel.AddMapping<NestedOfferMapped>("OfferEnd", p => p.E);

            var products = excel.Fetch<NestedProductMapped>().ToList();

            var expectedResult = new List<NestedProductMapped>
            {
                new NestedProductMapped { N = "Nudossi", Num = 60, P = 1.99m, O = new NestedOfferMapped { O = false, E = new DateTime(1970, 01, 01) } },
                new NestedProductMapped { N = "Halloren", Num = 33, P = 2.99m, O = new NestedOfferMapped { O = true, E = new DateTime(2015, 12, 31) } },
                new NestedProductMapped { N = "Filinchen", Num = 100, P = 0.99m, O = new NestedOfferMapped { O = false, E = new DateTime(1970, 01, 01) } },
            };

            CollectionAssert.AreEqual(expectedResult, products);
        }

        [Test]
        public void NestedProductIndexMappedTest()
        {
            var excel = new ExcelMapper(@"../../../xlsx/Products.xlsx")
            {
                HeaderRow = false,
                MinRowNumber = 1
            };

            excel.AddMapping<NestedProductMapped>(ExcelMapper.LetterToIndex("A"), p => p.N);
            excel.AddMapping<NestedProductMapped>(ExcelMapper.LetterToIndex("C"), p => p.Num);
            excel.AddMapping<NestedProductMapped>(ExcelMapper.LetterToIndex("D"), p => p.P);
            excel.AddMapping<NestedOfferMapped>(ExcelMapper.LetterToIndex("E"), p => p.O);
            excel.AddMapping<NestedOfferMapped>(ExcelMapper.LetterToIndex("F"), p => p.E);

            var products = excel.Fetch<NestedProductMapped>().ToList();

            var expectedResult = new List<NestedProductMapped>
            {
                new NestedProductMapped { N = "Nudossi", Num = 60, P = 1.99m, O = new NestedOfferMapped { O = false, E = new DateTime(1970, 01, 01) } },
                new NestedProductMapped { N = "Halloren", Num = 33, P = 2.99m, O = new NestedOfferMapped { O = true, E = new DateTime(2015, 12, 31) } },
                new NestedProductMapped { N = "Filinchen", Num = 100, P = 0.99m, O = new NestedOfferMapped { O = false, E = new DateTime(1970, 01, 01) } },
            };

            CollectionAssert.AreEqual(expectedResult, products);

            var file = "nestedindexmapped.xlsx";

            new ExcelMapper()
            {
                HeaderRow = false
            }.Save(file, products, "Products");

            excel = new ExcelMapper(file)
            {
                HeaderRow = false
            };

            var productsFetched = excel.Fetch<NestedProductMapped>().ToList();

            CollectionAssert.AreEqual(expectedResult, productsFetched);
        }

        class ProductStringArray
        {
            public string[] Products { get; set; }
        }

        [Test]
        public void StringArrayTest()
        {
            var excel = new ExcelMapper("../../../xlsx/ProductsJson.xlsx");
            excel.AddMapping<ProductStringArray>("Product", p => p.Products)
                .SetCellUsing((c, o) =>
                {
                    if (o is string[] s) c.SetCellValue(string.Join(",", s));
                    else c.SetCellValue((string)null);
                })
                .SetPropertyUsing(v =>
                {
                    if (v is string s) return s.Split(',');
                    else return Array.Empty<string>();
                });

            var ps = excel.Fetch<ProductStringArray>().ToList();

            Assert.AreEqual(3, ps.Count);
            Assert.True(ps.All(p => p.Products.Length == 4));
        }

        public record GuidProduct(Guid? Id, string Name, int NumberInStock, decimal Price);

        [Test]
        public void GuidTest()
        {
            var excel = new ExcelMapper("../../../xlsx/ProductsGuid.xlsx");

            var productsFetched = excel.Fetch<GuidProduct>().ToList();

            var expectedProducts = new List<GuidProduct>
            {
              new GuidProduct(new Guid("{6bba457c-00ee-4dc7-8002-967c760b428c}"), "Nudossi", 60, 1.99m),
              new GuidProduct(new Guid("{f833727a-63dd-483f-8ec0-4a667e707ebb}"), "Halloren", 33, 2.99m),
              new GuidProduct(new Guid("{eb966e28-b9b4-4dcc-b7b1-1521b53cb37f}"), "Filinchen", 100, 0.99m),
            };

            CollectionAssert.AreEqual(expectedProducts, productsFetched);

            var file = "guid.xlsx";

            new ExcelMapper().Save(file, expectedProducts, "Products");

            excel = new ExcelMapper(file);

            productsFetched = excel.Fetch<GuidProduct>().ToList();

            CollectionAssert.AreEqual(expectedProducts, productsFetched);
        }

        class ProductBase
        {
            public string Name { get; set; }
        }

        class ProductDerived : ProductBase
        {
            public int Number { get; set; }
        }

        [Test]
        public void DerivedSaveTest()
        {
            var products = new List<ProductDerived>
            {
                new ProductDerived { Name = "Name1", Number = 15 },
                new ProductDerived { Name = "Name2", Number = 41 }
            };

            var file = "derived.xlsx";

            new ExcelMapper().Save(file, products, "Products");

            var excel = new ExcelMapper(file);

            var productsFetched = excel.Fetch<ProductDerived>().ToList();

            Assert.AreEqual(products.Count, productsFetched.Count);

            for (int i = 0; i < products.Count; i++)
            {
                Assert.AreEqual(products[i].Name, productsFetched[i].Name);
                Assert.AreEqual(products[i].Number, productsFetched[i].Number);
            }
        }

        interface INumberInterface
        {
            int Number { get; set; }
        }

        class NumberClass : INumberInterface
        {
            public int Number { get; set; }
        }

        class InterfaceProduct
        {
            public string Name { get; set; }
            public INumberInterface Num { get; set; }
        }

        [Test]
        public void InterfaceTest()
        {
            var excel = new ExcelMapper("../../../xlsx/Products.xlsx");

            var productsFetched = excel.Fetch<InterfaceProduct>().ToList();

            var products = new List<InterfaceProduct>
            {
                new InterfaceProduct { Name = "Nudossi", Num = null },
                new InterfaceProduct { Name = "Halloren", Num = null },
                new InterfaceProduct { Name = "Filinchen", Num = null },
            };

            Assert.AreEqual(products.Count, productsFetched.Count);

            for (int i = 0; i < products.Count; i++)
            {
                Assert.AreEqual(products[i].Name, productsFetched[i].Name);
                Assert.AreEqual(products[i].Num, productsFetched[i].Num);
            }
        }

        [Test]
        public void InterfaceFactoryTest()
        {
            var excel = new ExcelMapper("../../../xlsx/Products.xlsx");

            excel.CreateInstance<INumberInterface>(() => new NumberClass());

            var productsFetched = excel.Fetch<InterfaceProduct>().ToList();

            var products = new List<InterfaceProduct>
            {
                new InterfaceProduct { Name = "Nudossi", Num = new NumberClass { Number = 60 } },
                new InterfaceProduct { Name = "Halloren", Num = new NumberClass { Number = 33 } },
                new InterfaceProduct { Name = "Filinchen", Num = new NumberClass { Number = 100 } },
            };

            Assert.AreEqual(products.Count, productsFetched.Count);

            for (int i = 0; i < products.Count; i++)
            {
                Assert.AreEqual(products[i].Name, productsFetched[i].Name);
                Assert.AreEqual(products[i].Num.Number, productsFetched[i].Num.Number);
            }
        }

        [Test]
        public void AttachWithPathUsingFetchTest()
        {
            var mapper = new ExcelMapper();

            mapper.Attach(@"../../../xlsx/Products.xlsx");

            var products = mapper.Fetch<Product>();

            CollectionAssert.AreEqual(new List<Product>
            {
                new Product { Name = "Nudossi", NumberInStock = 60, Price = 1.99m, Value = "C2*D2" },
                new Product { Name = "Halloren", NumberInStock = 33, Price = 2.99m, Value = "C3*D3" },
                new Product { Name = "Filinchen", NumberInStock = 100, Price = 0.99m, Value = "C5*D5" },
            }, products);
        }

        [Test]
        public void AttachWithStreamUsingFetchTest()
        {
            var stream = new FileStream(@"../../../xlsx/Products.xlsx", FileMode.Open, FileAccess.Read);
            var mapper = new ExcelMapper();

            mapper.Attach(stream);

            var products = mapper.Fetch<Product>();

            CollectionAssert.AreEqual(new List<Product>
            {
                new Product { Name = "Nudossi", NumberInStock = 60, Price = 1.99m, Value = "C2*D2" },
                new Product { Name = "Halloren", NumberInStock = 33, Price = 2.99m, Value = "C3*D3" },
                new Product { Name = "Filinchen", NumberInStock = 100, Price = 0.99m, Value = "C5*D5" },
            }, products);
            stream.Close();
        }

        [Test]
        public void AttachWithWorkbookUsingFetchTest()
        {
            var workbook = WorkbookFactory.Create(@"../../../xlsx/Products.xlsx");
            var mapper = new ExcelMapper();

            mapper.Attach(workbook);

            var products = mapper.Fetch<Product>();

            CollectionAssert.AreEqual(new List<Product>
            {
                new Product { Name = "Nudossi", NumberInStock = 60, Price = 1.99m, Value = "C2*D2" },
                new Product { Name = "Halloren", NumberInStock = 33, Price = 2.99m, Value = "C3*D3" },
                new Product { Name = "Filinchen", NumberInStock = 100, Price = 0.99m, Value = "C5*D5" },
            }, products);
        }

        private enum NameEnum
        {
            Nudossi,
            Halloren,
            Filinchen
        }

        private record EnumProduct
        {
            public NameEnum Name { get; }

            public EnumProduct(NameEnum name) => Name = name;
        }

        [Test]
        public void EnumTest()
        {
            var excel = new ExcelMapper("../../../xlsx/Products.xlsx");
            var products = excel.Fetch<EnumProduct>().ToList();

            CollectionAssert.AreEqual(new List<EnumProduct>
            {
                new EnumProduct(NameEnum.Nudossi),
                new EnumProduct(NameEnum.Halloren),
                new EnumProduct(NameEnum.Filinchen),
            }, products);

            var file = "enumsave.xlsx";

            excel.Save(file, products, "Products");

            var productsFetched = new ExcelMapper(file).Fetch<EnumProduct>().ToList();

            CollectionAssert.AreEqual(products, productsFetched);
        }

        [Test]
        public void EnumTestException()
        {
            var excel = new ExcelMapper("../../../xlsx/ProductsExceptionEnum.xlsx");
            var exception = Assert.Throws<ExcelMapperConvertException > (() => excel.Fetch<EnumProduct>().ToList());

            Assert.AreEqual(@"Unable to convert ""FilinchenError"" from [L:1]:[C:0] to Ganss.Excel.Tests.Tests+NameEnum.", exception.Message);
            Assert.AreEqual("Did not find a matching enum name.", exception.InnerException.Message);
        }

        private record BytesData
        {
            public byte[] TextData1 { get; set; }
            public byte[] TextData2 { get; set; }
            public byte[] RowVersion { get; set; }
        }

        [Test]
        public void BytesTest()
        {
            var excel = new ExcelMapper();
            var datas = new List<BytesData>
            {
                new BytesData(){TextData1 = Encoding.UTF8.GetBytes("ABC"), TextData2 = Encoding.UTF8.GetBytes("DEF"), RowVersion = new byte[]{1, 0, 0, 0}},
                new BytesData(){TextData1 = Encoding.UTF8.GetBytes("GHI"), TextData2 =                          null, RowVersion = new byte[]{2, 0, 0, 0}},
                new BytesData(){TextData1 =                          null, TextData2 = Encoding.UTF8.GetBytes("JKL"), RowVersion = new byte[]{3, 0, 0, 0}},
                new BytesData(){TextData1 = Encoding.UTF8.GetBytes("MNO"), TextData2 = Encoding.UTF8.GetBytes("PQR"), RowVersion = null}
            };

            var file = "bytesdata.xlsx";

            excel.Save(file, datas, "data", true, (colnum, value) =>
            {
                if (value != null)
                {
                    switch (colnum)
                    {
                        case "TextData1":
                            return Encoding.UTF8.GetString(value as byte[]);
                        case "RowVersion":
                            return BitConverter.ToInt32(value as byte[]);

                    }
                }
                return value;
            });

            var productsFetched = new ExcelMapper(file).Fetch<BytesData>(0, (colnum, value) =>
            {
                if (value != null && !string.IsNullOrEmpty(value.ToString()))
                {
                    switch (colnum)
                    {
                        case "TextData1":
                            return Encoding.UTF8.GetBytes(value.ToString());
                        case "RowVersion":
                            return BitConverter.GetBytes(Convert.ToInt32(value.ToString()));

                    }
                }
                return value;
            }).ToList();

            for (var index = 0; index < datas.Count; index++)
            {
                Assert.AreEqual(datas[index].TextData1, productsFetched[index].TextData1);
                Assert.AreEqual(datas[index].TextData2, productsFetched[index].TextData2);
                Assert.AreEqual(datas[index].RowVersion, productsFetched[index].RowVersion);
            }
        }

        private record MixedRecordProduct
        {
            public string Name { get; }

            public int Number { get; set; }

            public MixedRecordProduct(string name) => Name = name;
        }

        [Test]
        public void MixedRecordTest()
        {
            var excel = new ExcelMapper("../../../xlsx/Products.xlsx");
            var products = excel.Fetch<MixedRecordProduct>().ToList();

            CollectionAssert.AreEqual(new List<MixedRecordProduct>
            {
                new MixedRecordProduct("Nudossi") { Number = 60 },
                new MixedRecordProduct("Halloren") { Number = 33 },
                new MixedRecordProduct("Filinchen") { Number = 100 },
            }, products);
        }

        private record LinkRow
        {
            [Formula]
            public string Url { get; set; }
        }

        public static void LinkTest<T>(IEnumerable<T> rows)
        {
            var file = "linksave.xlsx";

            new ExcelMapper().Save(file, rows, "Links");

            var workbook = WorkbookFactory.Create(file);
            var excel = new ExcelMapper(workbook);
            var links = excel.Fetch<T>().ToList();

            CollectionAssert.AreEqual(rows, links);

            var sheet = workbook.GetSheetAt(0);

            foreach (var rownum in new[] { 1, 2 })
            {
                var row = sheet.GetRow(rownum);
                Assert.AreEqual(CellType.Formula, row.Cells.First().CellType);
            }
        }

        [Test]
        public void FormulaAttributeTest()
        {
            var rows = new[] { new LinkRow { Url = "HYPERLINK(\"https://www.google.com/\")" }, new LinkRow { Url = "HYPERLINK(\"https://www.microsoft.com/\")" } };
            LinkTest(rows);
        }

        private record LinkMappedRow
        {
            public string Url { get; set; }
        }

        [Test]
        public void FormulaMappedTest()
        {
            new ExcelMapper().AddMapping<LinkMappedRow>("Url", l => l.Url).AsFormula();
            var rows = new[] { new LinkMappedRow { Url = "HYPERLINK(\"https://www.google.com/\")" }, new LinkMappedRow { Url = "HYPERLINK(\"https://www.microsoft.com/\")" } };
            LinkTest(rows);
        }

        record ExcelRow
        {
            [Column(1)]
            public string Column1 { get; set; }
            [Column(2)]
            public string Column2 { get; set; }
            [Column(3)]
            public string Column3 { get; set; }
        }

        [Test]
        public void NumericColumnTest()
        {
            var excel = new ExcelMapper("../../../xlsx/numeric-header.xlsx");
            var rows = excel.Fetch<ExcelRow>().ToList();

            CollectionAssert.AreEqual(new List<ExcelRow>
            {
                new ExcelRow { Column1 = "value1", Column2 = "value2", Column3 = "value3" },
                new ExcelRow { Column1 = "value4", Column2 = "value5", Column3 = "value6" },
            }, rows);
        }

        public class BaseClass
        {
            [Column("Text base")]
            public virtual string Text { get; set; }
        }

        public class ChildClass : BaseClass
        {
            [Column("Text new")]
            public override string Text { get; set; }
        }

        [Test]
        public void VirtualTest()
        {
            var tf = new TypeMapperFactory();
            var tm = tf.Create(typeof(ChildClass));
            var ccs = new ExcelMapper("../../../xlsx/virtual.xlsx").Fetch<ChildClass>().ToList();

            Assert.AreEqual(1, ccs.Count);
            Assert.AreEqual("new", ccs[0].Text);
        }

        record VirtualSaveTestRecord
        {
            [Column("Text base")]
            public string TextBase { get; set; }
            [Column("Text new")]
            public string TextNew { get; set; }
        }

        [Test]
        public void VirtualSaveTest()
        {
            var tf = new TypeMapperFactory();
            var tm = tf.Create(typeof(ChildClass));

            new ExcelMapper().Save("virtualsave.xlsx", new[] { new ChildClass { Text = "test" } });

            var ccs = new ExcelMapper("virtualsave.xlsx").Fetch<VirtualSaveTestRecord>().ToList();

            Assert.AreEqual(1, ccs.Count);
            Assert.AreEqual("test", ccs[0].TextBase);
            Assert.AreEqual("test", ccs[0].TextNew);

            ccs = new ExcelMapper("../../../xlsx/virtual.xlsx").Fetch<VirtualSaveTestRecord>().ToList();

            Assert.AreEqual(1, ccs.Count);
            Assert.AreEqual("base", ccs[0].TextBase);
            Assert.AreEqual("new", ccs[0].TextNew);
        }

        class NoInheritBase
        {
            [Column("Text base", Inherit = false)]
            public virtual string Text { get; set; }
        }

        class NoInheritChild : NoInheritBase
        {
            [Column("Text new")]
            public override string Text { get; set; }
        }

        [Test]
        public void VirtualNoInheritTest()
        {
            var tf = new TypeMapperFactory();
            var tm = tf.Create(typeof(NoInheritChild));
            var ccs = new ExcelMapper("../../../xlsx/virtual.xlsx").Fetch<NoInheritChild>().ToList();

            Assert.AreEqual(1, ccs.Count);
            Assert.AreEqual("new", ccs[0].Text);
        }

        [Test]
        public void VirtualNoInheritSaveTest()
        {
            var tf = new TypeMapperFactory();
            var tm = tf.Create(typeof(NoInheritChild));

            new ExcelMapper().Save("virtualnoinheritsave.xlsx", new[] { new NoInheritChild { Text = "test" } });

            var ccs = new ExcelMapper("virtualnoinheritsave.xlsx").Fetch<VirtualSaveTestRecord>().ToList();

            Assert.AreEqual(1, ccs.Count);
            Assert.IsNull(ccs[0].TextBase);
            Assert.AreEqual("test", ccs[0].TextNew);
        }

        public record RowDef : RowDefInner
        {
            public bool Ignore { get; set; }
        }

        public record RowDefInner
        {
            public string Value { get; set; }

            public Func<dynamic, string> CustomOutput { get; set; } = null;
        }

        [Test]
        public void IgnoreInnerTest()
        {
            var mapper = new ExcelMapper(@"../../../xlsx/IgnoreInner.xlsx");
            mapper.Ignore<RowDef>(i => i.CustomOutput);
            var rows = mapper.Fetch<RowDef>().ToList();

            CollectionAssert.AreEqual(new List<RowDef>
            {
                new RowDef { Value = "A", Ignore = true },
                new RowDef { Value = "B", Ignore = true },
                new RowDef { Value = "C", Ignore = false },
            }, rows);
        }

        record InvalidDate(DateTime Date);

        [Test]
        public void InvalidDateTest()
        {
            var mapper = new ExcelMapper(@"../../../xlsx/InvalidDate.xlsx");
            Assert.Throws<ExcelMapperConvertException>(() => mapper.Fetch<InvalidDate>().ToList(),
                "Unable to convert \"55555555\" from [L:1]:[C:0] to System.DateTime.");
        }

        record InvalidJson
        {
            [Json]
            public string Json { get; set; }
        }

        [Test]
        public void InvalidJsonTest()
        {
            var mapper = new ExcelMapper(@"../../../xlsx/InvalidJson.xlsx");
            Assert.Throws<ExcelMapperConvertException>(() => mapper.Fetch<InvalidJson>().ToList(),
                @"Unable to convert ""{ ""key"": }"" from [L:1]:[C:0] to System.String.");
        }

        [Test]
        public void ParallelTest()
        {
            // see #208

            const int numThreads = 16;
            const int numRuns = 10;

            for (int i = 0; i < numRuns; i++)
            {
                var allGo = new ManualResetEvent(false);
                Exception firstException = null;
                var failures = 0;
                var waiting = numThreads;
                var threads = Enumerable.Range(0, numThreads)
                    .Take(numThreads)
                    .Select(m => new Thread(() =>
                    {
                        try
                        {
                            if (Interlocked.Decrement(ref waiting) == 0) allGo.Set();
                            var products = new ExcelMapper(@"../../../xlsx/Products.xlsx").Fetch<Product>().ToList();
                        }
                        catch (Exception ex)
                        {
                            Interlocked.CompareExchange(ref firstException, ex, null);
                            Interlocked.Increment(ref failures);
                        }
                    })).ToList();

                foreach (var thread in threads)
                    thread.Start();
                foreach (var thread in threads)
                    thread.Join();

                Assert.Null(firstException);
                Assert.AreEqual(0, failures);
            }
        }

        record DateTimeOffsetProduct(DateTimeOffset OfferEnd);

        [Test]
        public void DateTimeOffsetTest()
        {
            var products = new ExcelMapper(@"../../../xlsx/Products.xlsx").Fetch<DateTimeOffsetProduct>().ToList();

            void AssertProducts(IEnumerable<DateTimeOffsetProduct> products)
            {
                CollectionAssert.AreEqual(new List<DateTimeOffsetProduct>
                {
                    new DateTimeOffsetProduct(new DateTime(1970, 01, 01)),
                    new DateTimeOffsetProduct(new DateTime(2015, 12, 31)),
                    new DateTimeOffsetProduct(new DateTime(1970, 01, 01)),
                }, products);
            }

            AssertProducts(products);

            var file = "DateTimeOffsetProducts.xlsx";

            new ExcelMapper().Save(file, products);

            var savedProducts = new ExcelMapper(file).Fetch<DateTimeOffsetProduct>().ToList();

            AssertProducts(savedProducts);
        }

        record Customer(int Id, string Phone);

        [Test]
        public void ErrorTest()
        {
            var customers = new ExcelMapper(@"../../../xlsx/Error.xlsx").Fetch<Customer>().ToList();

            CollectionAssert.AreEqual(new List<Customer>
            {
                new Customer(1, "3001333"),
                new Customer(2, null),
                new Customer(3, "10031"),
            }, customers);
        }
    }
}
